"""Read-only reference reconciliation and isolated replays of completed reviews.

Run inside Docker; tools are included in the image. This checks calculation from
saved fact cards, not the reproducibility of a fresh semantic reading by a model.
No production claim, receipt, review, active pointer or workbook is written.
"""
import argparse
import copy
import csv
import hashlib
import json
import os
import shutil
import sqlite3
import subprocess
import sys
import tempfile
from pathlib import Path

from openpyxl import load_workbook

from app.common import DATA, read, sha, stamp


FIELDS = ('ai', 'department', 'comment', 'ban', 'ban_comment')


def readonly(path):
    connection = sqlite3.connect(path.resolve().as_uri() + '?mode=ro', uri=True)
    connection.row_factory = sqlite3.Row
    connection.execute('PRAGMA query_only=ON')
    return connection


def normalized(value):
    return '' if value is None else str(value)


def output(result):
    value = {key: normalized(result.get(key)) for key in FIELDS}
    value['clusters'] = [
        {'cluster': item['cluster'], 'score': item['score']}
        for item in result.get('clusters', [])
    ]
    return value


def digest(value):
    raw = json.dumps(value, ensure_ascii=False, sort_keys=True).encode()
    return hashlib.sha256(raw).hexdigest()


def sample_ids(baseline):
    selected = {}
    for ident, result in baseline.items():
        categories = ['ban'] if result['ban'] == 'да' else []
        if result['ban'] == 'нет' and not result['clusters']:
            categories.append('no_clusters')
        categories.extend(item['cluster'] for item in result['clusters'] if item['cluster'] in ('Завод', 'СААЖ', 'УЦТ', '100 Лидеров', 'КБ'))
        for category in categories:
            selected.setdefault(category, ident)
    return list(dict.fromkeys(selected.values()))


def protected_hashes(reference, source):
    paths = [DATA / 'state' / 'active.json', reference]
    paths.extend(path for path in (DATA / 'state' / 'runs').glob('*/state.sqlite*')
                 if not path.name.endswith('-shm') and (not path.name.endswith('-wal') or path.stat().st_size))
    paths.extend((source / name) for name in ('manifest.json', 'portable.json'))
    paths.extend(path for path in (source / 'inputs').rglob('*') if path.is_file())
    paths.extend(path for path in (DATA/'state').glob('*reviews.json') if path.is_file())
    paths.extend(path for path in (DATA/'state/source_blocks').glob('*.json'))
    return {str(path): sha(path) for path in sorted(paths)}


def reference_rows(path):
    workbook = load_workbook(path, read_only=True, data_only=False)
    try:
        sheet=workbook.worksheets[0]
        append=str(sheet.cell(1,1).value).strip()=='ID'
        id_column=0 if append else 1
        columns=(38,39,40,41,42) if append else (0,29,30,31,32)
        rows = sheet.iter_rows(min_row=2, values_only=True)
        result = {}
        for row in rows:
            if len(row) <= max(columns) or row[id_column] is None:
                raise ValueError('Reference row lacks an ID or result columns')
            if append and normalized(row[38])!='да':continue
            value=row[id_column]
            ident = str(int(value)) if isinstance(value, float) and value.is_integer() else str(value)
            if ident in result:
                raise ValueError('Duplicate reference ID: ' + ident)
            result[ident] = dict(zip(FIELDS, (normalized(row[index]) for index in columns)))
        return result
    finally:
        workbook.close()


def clone_run(source, destination, ids=()):
    target = destination / 'state' / 'runs' / source.name
    target.mkdir(parents=True)
    for name in ('inputs', 'engine', 'candidate_source_blocks', 'audit/parent_fact_cards'):
        if (source / name).exists():
            shutil.copytree(source / name, target / name)
    for path in source.glob('*.json'):
        shutil.copy2(path,target/path.name)
    for name in ('image_reviews.json','page_reviews.json','source_blocks'):
        path=DATA/'state'/name
        if path.is_dir():shutil.copytree(path,destination/'state'/name)
        elif path.is_file():shutil.copy2(path,destination/'state'/name)
    old = readonly(source / 'state.sqlite')
    new = sqlite3.connect(target / 'state.sqlite')
    try:
        old.backup(new)
        for ident in ids:
            row=old.execute('SELECT files FROM candidates WHERE id=?',(ident,)).fetchone()
            if not row:continue
            for file in json.loads(row['files']):
                rel=Path(file['path'])
                if rel.is_absolute() or '..' in rel.parts:raise ValueError('Unsafe source path')
                original=DATA/'cvs'/rel
                if not original.is_file():continue  # preview reports missing originals
                if sha(original)!=file['sha256']:raise ValueError('Changed original: '+str(rel))
                copy=destination/'cvs'/rel;copy.parent.mkdir(parents=True,exist_ok=True)
                shutil.copy2(original,copy)
    finally:
        new.close()
        old.close()
    (destination / 'state' / 'active.json').write_text(json.dumps({'run': source.name}))
    (destination / '.repeatability-sandbox.json').write_text(json.dumps({'purpose': 'isolated-review-replay'}))


def replay(ids, cards_file=None):
    marker = DATA / '.repeatability-sandbox.json'
    if not marker.is_file() or read(marker).get('purpose') != 'isolated-review-replay':
        raise ValueError('Replay worker requires an isolated audit sandbox; production receipts must not be written')
    # HR_DATA points to a throwaway copy before app modules are imported.
    from app.common import active
    from app.service import preview, show

    replacements = {str(card['id']): card for card in read(cards_file)} if cards_file else {}
    connection = readonly(active() / 'state.sqlite')
    results = []
    try:
        for ident in ids:
            try:
                row = connection.execute('SELECT * FROM candidates WHERE id=?', (ident,)).fetchone()
                if not row or row['review_status'] != 'reviewed':
                    raise ValueError('Replay requires a completed review')
                previous = json.loads(row['result'])
                card = copy.deepcopy(replacements[ident] if cards_file else previous['fact_card'])
                actual_hash = hashlib.sha256(row['text'].encode()).hexdigest()
                if cards_file and card.get('full_text_review_complete') and not card.get('reviewed_text_sha256'):
                    card['reviewed_text_sha256'] = actual_hash
                if card.get('reviewed_text_sha256') != actual_hash:
                    raise ValueError('Saved fact card does not match the stored full text')
                # Register only sandbox receipts. This is a technical replay of
                # an existing completed review, not a new reading claim.
                offset = 0
                while True:
                    part = show(ident, offset, 10000)
                    if part['next_offset'] is None:
                        break
                    offset = part['next_offset']
                card['full_text_review_complete'] = True
                value = output(preview(card))
                results.append({'id': ident, 'output': value, 'sha256': digest(value)})
            except Exception as error:
                results.append({'id': ident, 'error': type(error).__name__ + ': ' + str(error)})
    finally:
        connection.close()
    return results


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--run', default='full_20261016_132123')
    parser.add_argument('--reference', type=Path)
    parser.add_argument('--ids', nargs='+', help='Default: representative completed cases selected by result category')
    parser.add_argument('--repeats', type=int, default=3)
    parser.add_argument('--report', type=Path)
    parser.add_argument('--cards-file', type=Path, help='Explicit audit cards from a separate full-text review; never submitted')
    parser.add_argument('--worker', action='store_true', help=argparse.SUPPRESS)
    args = parser.parse_args()
    if args.worker:
        print(json.dumps(replay(args.ids, args.cards_file), ensure_ascii=False))
        return
    if not args.reference or not args.report or args.repeats < 2:
        parser.error('--reference, --report and --repeats >= 2 are required')
    if Path(args.run).name != args.run:
        parser.error('--run must be a run name')
    source = DATA / 'state' / 'runs' / args.run
    before = protected_hashes(args.reference, source)
    reference = reference_rows(args.reference)
    connection = readonly(source / 'state.sqlite')
    try:
        rows = connection.execute("SELECT id,result,review_status FROM candidates WHERE review_status='reviewed' ORDER BY source_row").fetchall()
        baseline = {str(row['id']): output(json.loads(row['result'])) for row in rows}
        from app.books import export_comments
        projected=export_comments(source,rows) if read(source/'portable.json').get('export_layout')=='append_results' else rows
        excel_baseline={str(row['id']):output(json.loads(row['result'])) for row in projected}
        integrity = connection.execute('PRAGMA integrity_check').fetchone()[0]
    finally:
        connection.close()
    if integrity != 'ok':
        raise ValueError('Source integrity check failed: ' + integrity)
    if not args.ids:
        args.ids = sample_ids(baseline)
    if not args.ids:
        raise ValueError('No completed fact cards are available for replay')
    differences = []
    for ident in sorted(set(reference) | set(baseline)):
        if ident not in reference or ident not in baseline:
            differences.append({'id': ident, 'field': 'presence', 'excel': ident in reference, 'sqlite': ident in baseline})
            continue
        for field in FIELDS:
            if reference[ident][field] != excel_baseline[ident][field]:
                differences.append({'id': ident, 'field': field, 'excel': reference[ident][field], 'sqlite': excel_baseline[ident][field]})
    passes = []
    with tempfile.TemporaryDirectory(prefix='hrhive-repeatability-') as scratch:
        for index in range(args.repeats):
            sandbox = Path(scratch) / str(index + 1)
            clone_run(source, sandbox, args.ids)
            env = dict(os.environ, HR_DATA=str(sandbox), PYTHONHASHSEED=str(index + 1))
            command = [sys.executable, '-m', 'tools.audit_repeatability', '--worker', '--ids', *args.ids]
            if args.cards_file:
                command.extend(['--cards-file', str(args.cards_file.resolve())])
            child = subprocess.run(command, env=env, capture_output=True, text=True)
            if child.returncode:
                raise RuntimeError(child.stderr)
            values = json.loads(child.stdout)
            passes.append(values)
            print(json.dumps({'pass': index + 1, 'cases': len(values), 'errors': sum('error' in item for item in values)}, ensure_ascii=False), flush=True)
    comparisons = []
    for index, ident in enumerate(args.ids):
        values = [items[index] for items in passes]
        successful = all('output' in value for value in values)
        previous = baseline.get(ident)
        first = values[0].get('output') if successful else None
        old_clusters = (previous or {}).get('clusters', [])
        new_clusters = (first or {}).get('clusters', [])
        old_scores = {item['cluster']: item['score'] for item in old_clusters}
        new_scores = {item['cluster']: item['score'] for item in new_clusters}
        comparisons.append({
            'id': ident,
            'baseline': previous,
            'passes': values,
            'repeatable': successful and len({value['sha256'] for value in values}) == 1,
            'matches_baseline': successful and all(value['output'] == previous for value in values),
            'different_fields': [key for key in (*FIELDS, 'clusters') if successful and first[key] != previous[key]],
            'cluster_set_matches': successful and set(old_scores) == set(new_scores),
            'cluster_order_matches': successful and list(old_scores) == list(new_scores),
            'score_deltas': {name: new_scores[name] - old_scores[name] for name in sorted(set(old_scores) & set(new_scores))},
            'ban_matches': successful and first['ban'] == previous['ban'],
        })
    after = protected_hashes(args.reference, source)
    unchanged = before == after
    report = {
        'created_utc': stamp(), 'run': args.run,
        'as_of': read(source / 'manifest.json')['as_of'],
        'reference': str(args.reference), 'reference_sha256': before[str(args.reference)],
        'reference_rows': len(reference), 'sqlite_rows': len(baseline),
        'reference_differences': differences,
        'repeats': args.repeats, 'cases': comparisons,
        'production_unchanged': unchanged, 'protected_hashes': before,
        'cards_source': str(args.cards_file) if args.cards_file else 'Saved completed fact cards',
        'cards_sha256': sha(args.cards_file) if args.cards_file else None,
        'method': 'Fresh Python processes with distinct hash seeds; actual preview on fixed audit cards and full texts in throwaway SQLite snapshots. No submit or production receipts.',
        'limitation': 'Each pass reuses fixed cards. Custom audit cards represent one separate full-text review, not three independent model reviews. This does not establish repeatability of fresh model-generated facts or scores.',
    }
    args.report.parent.mkdir(parents=True, exist_ok=True)
    args.report.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
    csv_path = args.report.with_suffix('.csv')
    with csv_path.open('w', encoding='utf-8-sig', newline='') as stream:
        writer = csv.writer(stream, delimiter=';')
        writer.writerow(['ID', 'Сохранённый результат', *['Прогон ' + str(i + 1) for i in range(args.repeats)], 'Повторяемость', 'Совпадение с историей', 'Набор кластеров совпадает', 'Порядок совпадает', 'Изменение баллов', 'Различающиеся поля', 'Прежний комментарий', 'Новый комментарий'])
        for case in comparisons:
            writer.writerow([case['id'], (case['baseline'] or {}).get('department', ''), *[value.get('output', {}).get('department', value.get('error', '')) for value in case['passes']], 'да' if case['repeatable'] else 'нет', 'да' if case['matches_baseline'] else 'нет', 'да' if case['cluster_set_matches'] else 'нет', 'да' if case['cluster_order_matches'] else 'нет', json.dumps(case['score_deltas'], ensure_ascii=False), ', '.join(case['different_fields']), (case['baseline'] or {}).get('comment', ''), case['passes'][0].get('output', {}).get('comment', '')])
    print(json.dumps({
        'reference_rows': len(reference), 'reference_differences': len(differences),
        'repeats': args.repeats, 'cases': len(comparisons),
        'repeatable': sum(case['repeatable'] for case in comparisons),
        'matching_baseline': sum(case['matches_baseline'] for case in comparisons),
        'production_unchanged': unchanged, 'report': str(args.report),
    }, ensure_ascii=False, indent=2))
    if not unchanged or differences or any(not case['repeatable'] or not case['matches_baseline'] for case in comparisons):
        sys.exit(1)


if __name__ == '__main__':
    main()
