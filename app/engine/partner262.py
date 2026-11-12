"""Frozen, human-verified employer relationships. No network or database access."""
import hashlib
import json
import re
from datetime import date, datetime
from pathlib import Path
from urllib.parse import urlsplit

REGISTRY_PATH = 'amendments/rule262/registry.json'


def canonical_hash(value):
    return hashlib.sha256(json.dumps(value, ensure_ascii=False, sort_keys=True,
                                     separators=(',', ':')).encode()).hexdigest()


def snapshot_path(root, relative):
    relative = Path(str(relative))
    if relative.is_absolute() or '..' in relative.parts or not relative.parts:
        raise ValueError('Недопустимый путь снимка amendment')
    path = root / relative
    if not path.resolve().is_relative_to(root.resolve()) or path.is_symlink():
        raise ValueError('Снимок amendment выходит из каталога')
    if not path.is_file():
        raise ValueError('Отсутствует снимок amendment: ' + str(relative))
    return path


def validate_registry(path, expected_as_of=None):
    path = Path(path)
    registry = json.loads(path.read_text('utf-8-sig'))
    if registry.get('schema_version') != 1:
        raise ValueError('Неизвестная схема реестра 262')
    policy = registry.get('amendment', {})
    if policy.get('rule_row') != 262 or policy.get('months') != 6:
        raise ValueError('Amendment должен уточнять строку 262 и срок 6 месяцев')
    for field in ('id', 'approved_by', 'user_instruction', 'approved_at', 'as_of'):
        if not str(policy.get(field, '')).strip():
            raise ValueError('Не заполнено поле amendment: ' + field)
    as_of = date.fromisoformat(policy['as_of'])
    datetime.fromisoformat(policy['approved_at'].replace('Z', '+00:00'))
    if expected_as_of and as_of.isoformat() != expected_as_of:
        raise ValueError('Дата amendment отличается от даты запуска')
    domains = policy.get('approved_domains', [])
    if not domains or any(not re.fullmatch(r'[a-z0-9.-]+', x) for x in domains):
        raise ValueError('Нужны проверенные официальные домены amendment')
    supplementary = registry.get('supplementary_sources', [])
    if not isinstance(supplementary, list):
        raise ValueError('supplementary_sources должен быть списком')
    for proof in supplementary:
        if not isinstance(proof, dict) or not proof.get('snapshot') or not proof.get('sha256'):
            raise ValueError('Не заполнен дополнительный снимок registry')
        source = snapshot_path(path.parent, proof['snapshot'])
        if hashlib.sha256(source.read_bytes()).hexdigest() != proof['sha256']:
            raise ValueError('Изменён хеш дополнительного снимка registry')
    seen = set()
    entities = registry.get('entities')
    if not isinstance(entities, list):
        raise ValueError('Нужен список entities реестра 262')
    for entity in entities:
        key = entity.get('entity_id')
        if not isinstance(key, str) or not key or key in seen:
            raise ValueError('Повторный или пустой entity_id')
        seen.add(key)
        if not str(entity.get('canonical_name', '')).strip():
            raise ValueError('Не указано юридическое имя')
        aliases = entity.get('aliases')
        if not isinstance(aliases, list) or not aliases or any(not isinstance(x, str) or not x.strip() for x in aliases):
            raise ValueError('Нужны точные проверенные aliases')
        if entity.get('certainty', 'confirmed') not in ('confirmed', 'probable'):
            raise ValueError('Неизвестная достоверность связи')
        if not isinstance(entity.get('auto_match', False), bool):
            raise ValueError('auto_match должен быть логическим')
        if entity.get('relation') not in ('official_partner', 'resident'):
            raise ValueError('Не подтвержден вид связи с ОЭЗ')
        if not str(entity.get('temporal_basis', '')).strip():
            raise ValueError('Нужна применимость источника к дате проверки')
        for field, size in (('inn', (10, 12)), ('ogrn', (13, 15))):
            value = entity.get(field)
            if value is not None and (not isinstance(value, str) or not value.isdigit() or len(value) not in size):
                raise ValueError('Некорректный ' + field)
        start = date.fromisoformat(entity['valid_from']) if entity.get('valid_from') else None
        end = date.fromisoformat(entity['valid_to']) if entity.get('valid_to') else None
        if start and end and start > end:
            raise ValueError('Обратный интервал связи с ОЭЗ')
        evidence = entity.get('evidence')
        if not isinstance(evidence, list) or not evidence:
            raise ValueError('Нужен официальный первоисточник связи с ОЭЗ')
        for proof in evidence:
            for field in ('url', 'title', 'excerpt', 'snapshot', 'sha256', 'checked_at', 'reviewer'):
                if not str(proof.get(field, '')).strip():
                    raise ValueError('Не заполнено доказательство: ' + field)
            url = urlsplit(proof['url'])
            if (url.scheme != 'https' or url.username or url.password or
                    not url.hostname or not any(url.hostname == d or url.hostname.endswith('.' + d) for d in domains)):
                raise ValueError('Источник не относится к проверенному официальному домену')
            date.fromisoformat(proof['checked_at'])
            source = snapshot_path(path.parent, proof['snapshot'])
            raw = source.read_bytes()
            if hashlib.sha256(raw).hexdigest() != proof['sha256']:
                raise ValueError('Изменён хеш официального снимка')
            if proof.get('raw_snapshot') or proof.get('raw_sha256'):
                raw_source = snapshot_path(path.parent, proof.get('raw_snapshot', ''))
                if hashlib.sha256(raw_source.read_bytes()).hexdigest() != proof.get('raw_sha256'):
                    raise ValueError('Изменён хеш исходного официального HTML')
            if proof['excerpt'] not in raw.decode('utf-8-sig'):
                raise ValueError('Цитата не найдена в официальном снимке')
    return registry


def load_registry(run):
    path = Path(run) / 'inputs' / REGISTRY_PATH
    if not path.exists():
        return None
    manifest = json.loads((Path(run) / 'manifest.json').read_text('utf-8-sig'))
    return validate_registry(path, manifest['as_of'])


def verify_amendment_run(run, card=None, deep=True):
    """Fail closed on altered provenance, without weakening engine hash checks."""
    run = Path(run)
    meta = json.loads((run/'portable.json').read_text('utf-8-sig'))
    receipt_path = run/'amendment_receipt.json'
    if hashlib.sha256(receipt_path.read_bytes()).hexdigest() != meta.get('amendment_receipt_sha256'):
        raise ValueError('Изменён receipt amendment-версии')
    receipt = json.loads(receipt_path.read_text('utf-8-sig'))
    if (receipt.get('amendment_key') != meta.get('amendment_key') or
            receipt.get('new_rules_hash') != meta.get('rules_hash') or
            receipt.get('new_implementation_hash') != meta.get('implementation_hash')):
        raise ValueError('Receipt не соответствует версии amendment')
    if hashlib.sha256((run/'archive_manifest.json').read_bytes()).hexdigest() != receipt['archive_manifest_sha256']:
        raise ValueError('Изменён архивный scope amendment')
    manifest = json.loads((run/'manifest.json').read_text('utf-8-sig'))
    if manifest.get('input_hashes') != receipt['input_hashes']:
        raise ValueError('Изменён manifest amendment')
    for relative, fingerprint in receipt['input_hashes'].items():
        path = snapshot_path(run/'inputs', relative)
        if hashlib.sha256(path.read_bytes()).hexdigest() != fingerprint:
            raise ValueError('Изменён снимок источника amendment: ' + relative)
    load_registry(run)
    if deep:
        audit = run/'audit'
        for name, key in [('parent_state.sqlite','parent_database_snapshot_sha256'),
                          ('parent_portable.json','parent_portable_sha256'),
                          ('parent_manifest.json','parent_manifest_sha256')]:
            if hashlib.sha256((audit/name).read_bytes()).hexdigest() != receipt[key]:
                raise ValueError('Изменён родительский снимок amendment: '+name)
        actual_engine = {p.name:hashlib.sha256(p.read_bytes()).hexdigest() for p in (audit/'parent_engine').glob('*.py')}
        if actual_engine != receipt['parent_engine_files']:
            raise ValueError('Изменён снимок родительского движка amendment')
    if card and str(card.get('id')) not in {r['id'] for r in receipt['candidates']}:
        raise ValueError('ID вне архивного scope amendment')
    if card and card.get('amendment_provenance'):
        proof = card['amendment_provenance']
        record = next((r for r in receipt['candidates'] if r['id'] == str(card['id'])), None)
        if (not record or not record.get('parent_card_sha256') or
                proof.get('parent_card_sha256') != record['parent_card_sha256'] or
                proof.get('parent_result_sha256') != record['parent_result_sha256'] or
                proof.get('parent_run') != receipt['parent_run'] or
                proof.get('amendment_key') != receipt['amendment_key'] or
                card.get('reviewed_text_sha256') != record['text_sha256'] or
                not record['full_receipt_coverage']):
            raise ValueError('Недействительное происхождение перенесённой карточки')
        original = json.loads((run/'audit/parent_fact_cards'/(str(card['id'])+'.json')).read_text('utf-8-sig'))
        if canonical_hash(original) != record['parent_card_sha256']:
            raise ValueError('Изменена исходная карточка amendment')
        allowed = {'amendment_provenance', 'employer_relation_evidence', 'employer_relation_nonmatches'}
        if ({k:v for k,v in card.items() if k not in allowed} !=
                {k:v for k,v in original.items() if k not in allowed}):
            raise ValueError('Перенесённые факты/баллы изменены; требуется отдельное новое прочтение')
    return receipt


def prepare_relation_evidence(card, text, registry):
    """Suggest exact identities; opt-in automation only for curated names."""
    from .ban_engine import company_name, norm, normalized_positions
    result = json.loads(json.dumps(card, ensure_ascii=False))
    suggestions = []
    bindings = list(result.get('employer_relation_evidence', []))
    for index, job in enumerate(result.get('jobs', [])):
        exact_name = norm(job[0]);name = company_name(job[0])
        matches = [e for e in registry['entities'] if exact_name and exact_name in {norm(a) for a in e['aliases']}]
        for entity in matches:
            item = {'job_index':index,'entity_id':entity['entity_id'],
                    'registry_record_sha256':canonical_hash(entity),'requires_review':True}
            if entity.get('auto_match') and len(matches) == 1:
                normalized, position = normalized_positions(text)
                match = re.search(r'\b'+re.escape(exact_name)+r'\b', normalized)
                if match:
                    anchor = text[position(match.start()):position(match.end()-1)+len(exact_name.split()[-1])]
                    item['requires_review'] = False
                    if not any(p.get('job_index')==index and p.get('entity_id')==entity['entity_id'] for p in bindings):
                        bindings.append({k:v for k,v in item.items() if k!='requires_review'} |
                            {'source_cv_anchor':anchor,'certainty':entity.get('certainty','confirmed')})
            suggestions.append(item)
    if bindings:result['employer_relation_evidence']=bindings
    return result, suggestions


def relationship_findings(run, card, text, rule, finding):
    """Require a reviewed job, exact entity and frozen official + CV evidence."""
    from .ban_engine import company_name, norm
    registry = load_registry(run)
    bindings = card.get('employer_relation_evidence', [])
    if not isinstance(bindings, list):
        raise ValueError('employer_relation_evidence должен быть списком')
    if not registry:
        if bindings:raise ValueError('Нет frozen реестра для связи работодателя с ОЭЗ')
        return []
    if rule.get('row') != 262 or rule.get('months') != 6:
        raise ValueError('Строка 262 не соответствует согласованным 6 месяцам')
    entities = {e['entity_id']:e for e in registry['entities']}
    findings = [];seen=set()
    as_of = date.fromisoformat(registry['amendment']['as_of'])
    for proof in bindings:
        if not isinstance(proof, dict):raise ValueError('Неверная привязка работодателя к ОЭЗ')
        index=proof.get('job_index');entity=entities.get(proof.get('entity_id'))
        if (not isinstance(index, int) or isinstance(index,bool) or index<0 or
                index>=len(card.get('jobs', [])) or not entity):
            raise ValueError('Неизвестная работа или entity_id связи с ОЭЗ')
        if (index,entity['entity_id']) in seen:raise ValueError('Повторная привязка к строке 262')
        seen.add((index,entity['entity_id']))
        if proof.get('registry_record_sha256') != canonical_hash(entity):
            raise ValueError('Изменён хеш записи реестра работодателей')
        job=card['jobs'][index]
        if len(job)<6:raise ValueError('Не полная трудовая запись')
        name=company_name(job[0]);aliases={norm(a) for a in entity['aliases']}
        if not name or norm(job[0]) not in aliases:
            raise ValueError('Работодатель не совпадает с проверенными aliases юридического лица')
        anchor=proof.get('source_cv_anchor')
        if not isinstance(text,str) or not isinstance(anchor,str) or not anchor or anchor not in text:
            raise ValueError('Не подтверждена дословная CV-цитата работодателя')
        if not re.search(r'\b'+re.escape(norm(job[0]))+r'\b', norm(anchor)):
            raise ValueError('CV-цитата не содержит подтвержденное имя работодателя')
        certainty=proof.get('certainty','confirmed')
        if certainty not in ('confirmed','probable'):raise ValueError('Неизвестная достоверность привязки')
        # Explicit historical bounds can exclude a connection. Unknown bounds
        # never invent a starting date or erase a probable verified connection.
        if entity.get('valid_from') and date.fromisoformat(entity['valid_from'])>as_of:continue
        if entity.get('valid_to') and date.fromisoformat(entity['valid_to'])<as_of:continue
        if entity.get('certainty','confirmed')=='probable':certainty='probable'
        relation_label = {'official_partner': 'официальный партнёр', 'resident': 'резидент ОЭЗ'}[entity['relation']]
        reason=('Проверенная связь конкретного работодателя с ОЭЗ Алабуга: '+relation_label+
                '; '+entity['temporal_basis'])
        f=finding(rule,job,certainty,reason)
        f['source_cv_anchor']=anchor
        f['relationship_evidence']={
            'entity_id':entity['entity_id'],'canonical_name':entity['canonical_name'],
            'relation':entity['relation'],'registry_record_sha256':canonical_hash(entity),
            'registry_path':'inputs/'+REGISTRY_PATH,
            'registry_sha256':hashlib.sha256((Path(run)/'inputs'/REGISTRY_PATH).read_bytes()).hexdigest(),
            'temporal_basis':entity['temporal_basis'],'evidence':entity['evidence'],
            'amendment':registry['amendment'],'job_index':index}
        findings.append(f)
    return findings
