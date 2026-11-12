"""Amendment regression tests use synthetic temporary data exclusively."""
import copy
import hashlib
import json
import shutil
import tempfile
import unittest
from datetime import date
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

from openpyxl import Workbook
from app import batch, books, common, documents, service, amendments
from app.engine.pilot150_bans import StructuredBans


class AmendmentTest(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.patches = [patch.object(m, 'DATA', self.root) for m in
                        (batch, books, common, documents, service)]
        for p in self.patches:p.start()
        common.layout()
        self.registry_dir = self.root/'registry'
        (self.registry_dir/'sources').mkdir(parents=True)
        source = self.registry_dir/'sources/official.txt'
        source.write_text('Официальный партнер: ООО «Точный Партнер»',encoding='utf-8')
        raw = self.registry_dir/'sources/raw.html'
        raw.write_text('<h2>ООО «Точный Партнер»</h2>',encoding='utf-8')
        self.registry = {'schema_version':1,'amendment':{'id':'test-262','rule_row':262,
            'months':6,'as_of':'2026-10-22','approved_by':'synthetic coordinator',
            'user_instruction':'Все официальные партнеры ОЭЗ',
            'approved_at':'2026-10-23','approved_domains':['official.example']},
            'entities':[{'entity_id':'exact','canonical_name':'ООО «Точный Партнер»',
                'aliases':['ООО «Точный Партнер»'],'relation':'official_partner',
                'certainty':'probable','temporal_basis':'Наблюдение 23.10; применимость 22.10 требует уточнения.',
                'evidence':[{'url':'https://official.example/partners','title':'Официальный список',
                    'excerpt':'ООО «Точный Партнер»','snapshot':'sources/official.txt',
                    'sha256':common.sha(source),'checked_at':'2026-10-23','reviewer':'synthetic',
                    'raw_snapshot':'sources/raw.html','raw_sha256':common.sha(raw)}]}]}
        extra=self.registry_dir/'sources/index.html';extra.write_text('<h1>Полный каталог</h1>',encoding='utf-8')
        self.registry['supplementary_sources']=[{'snapshot':'sources/index.html','sha256':common.sha(extra)}]
        self.registry_file = self.registry_dir/'registry.json'
        common.dump(self.registry_file,self.registry)

    def tearDown(self):
        for p in reversed(self.patches):p.stop()
        self.temp.cleanup()

    def make_parent(self):
        parent = self.root/'state/runs/parent'
        (parent/'inputs').mkdir(parents=True)
        book=Workbook();sheet=book.active;sheet.title='Source'
        sheet.append(['ID']+['Поле']*37)
        for ident in ('101','102'):sheet.append([ident]+['Исходное']*37)
        book.save(parent/'inputs/template.xlsx');book.close()
        original=self.root/'cvs/101.txt';original.write_text('ООО «Точный Партнер»: монтаж оборудования.',encoding='utf-8')
        other=self.root/'cvs/102.txt';other.write_text('Нераспознанный документ',encoding='utf-8')
        files=[{'id':i,'path':i+'.txt','sha256':common.sha(self.root/'cvs'/(i+'.txt')),
                'size':(self.root/'cvs'/(i+'.txt')).stat().st_size}for i in ('101','102')]
        common.dump(parent/'archive_manifest.json',{'ids':['101','102'],'files':files})
        common.dump(parent/'manifest.json',{'as_of':'2026-10-22','input_hashes':{
            'template.xlsx':common.sha(parent/'inputs/template.xlsx')}})
        common.dump(parent/'portable.json',{'batch_key':'parent-archive','import_complete':True,
            'implementation_hash':service.implementation_hash(),'rules_hash':service.rules_hash(parent/'inputs'),
            'source_book':'inputs/template.xlsx','sheet':'Source','id_column':1,'source_column_count':38,
            'export_layout':'append_results','output_name':'parent.xlsx','output_path':str(self.root/'exports/parent.xlsx')})
        c=common.connect(parent/'state.sqlite')
        c.executescript('''CREATE TABLE candidates(id TEXT PRIMARY KEY,source_row INTEGER,files TEXT,fingerprint TEXT,extraction_status TEXT,text TEXT,documents TEXT,error TEXT,review_status TEXT,result TEXT,reviewed_utc TEXT,exported_revision INTEGER);
            CREATE TABLE fact_cards(id TEXT PRIMARY KEY,card TEXT,text_sha256 TEXT,updated_utc TEXT);
            CREATE TABLE documents(hash TEXT PRIMARY KEY,path TEXT,kind TEXT,status TEXT,pages TEXT,error TEXT,updated_utc TEXT);''')
        text=original.read_text();th=hashlib.sha256(text.encode()).hexdigest()
        self.card={'id':'101','name':'Синтетический Кандидат','phones':[],
            'jobs':[['ООО «Точный Партнер»','2025-01','2026-08','Монтаж оборудования','', 'private']],
            'tracks':{},'reviews':[],'exclusions':[],'comment':'Подтвержден монтаж оборудования',
            'reviewed_text_sha256':th,'full_text_review_complete':True}
        self.result={'id':'101','ai':'да','ban':'нет','department':'Подходящих кластеров не выявлено',
            'fact_card':self.card,'rules_hash':'old','comment':self.card['comment']}
        doc={'sha256':files[0]['sha256'],'complete':True,'pages':[{'page':1,'method':'text','text':text}]}
        with c:
            for idx,i in enumerate(('101','102')):
                reviewed=i=='101'
                c.execute('INSERT INTO candidates VALUES(?,?,?,?,?,?,?,?,?,?,?,?)',
                    (i,idx+2,json.dumps([files[idx]],ensure_ascii=False),files[idx]['sha256'],
                     'ready' if reviewed else 'error',text if reviewed else None,
                     json.dumps([doc])if reviewed else None,None if reviewed else 'Плохой скан',
                     'reviewed' if reviewed else 'pending',json.dumps(self.result,ensure_ascii=False)if reviewed else None,
                     'previous' if reviewed else None,1 if reviewed else None))
        c.close();service.prepare(parent)
        c=common.connect(parent/'state.sqlite')
        with c:
            # Two real intervals prove reading, never a synthetic replacement.
            c.execute('INSERT INTO portable_receipts VALUES(?,?,?,?)',('101',th,0,10))
            c.execute('INSERT INTO portable_receipts VALUES(?,?,?,?)',('101',th,10,len(text)))
            c.execute('INSERT INTO portable_claims VALUES(?,?,?)',('101','old-worker','previous'))
            c.execute('INSERT INTO portable_claims VALUES(?,?,?)',('102','pending-worker','previous'))
            c.execute('INSERT INTO portable_revisions VALUES(?,?,?,?,?)',(1,'101','completed',json.dumps(self.result),'previous'))
        c.close()
        common.dump(self.root/'state/active.json',{'run':'parent'})
        self.snapshot=self.root/'engine-snapshot';shutil.copytree(Path(service.__file__).parent/'engine',self.snapshot)
        return parent

    def test_registry_rejects_missing_or_tampered_primary_source(self):
        amendments.validate_registry(self.registry_file,'2026-10-22')
        (self.registry_dir/'sources/raw.html').write_text('Подмена',encoding='utf-8')
        with self.assertRaisesRegex(ValueError,'HTML'):
            amendments.validate_registry(self.registry_file,'2026-10-22')
        for field,value in [('snapshot','../escape'),('url','https://unofficial.example/list')]:
            changed=copy.deepcopy(self.registry);changed['entities'][0]['evidence'][0][field]=value
            common.dump(self.registry_file,changed)
            with self.assertRaises(ValueError):amendments.validate_registry(self.registry_file)

    def test_supplementary_catalog_tamper_is_rejected(self):
        (self.registry_dir/'sources/index.html').write_text('Подмена каталога',encoding='utf-8')
        with self.assertRaisesRegex(ValueError,'дополнительного'):
            amendments.validate_registry(self.registry_file)

    def test_migration_preserves_parent_history_scope_errors_and_real_receipts(self):
        parent=self.make_parent();before={p.relative_to(parent).as_posix():common.sha(p) for p in parent.rglob('*')if p.is_file() and not p.name.endswith(('-wal','-shm'))}
        report=batch.amend_archive('parent',self.registry_file,self.snapshot)
        child=common.active();meta=common.read(child/'portable.json')
        self.assertEqual(meta['export_layout'],'append_results');self.assertEqual(meta['id_column'],1)
        self.assertNotEqual(meta['output_path'],common.read(parent/'portable.json')['output_path'])
        self.assertEqual(before,{p.relative_to(parent).as_posix():common.sha(p) for p in parent.rglob('*')if p.is_file() and not p.name.endswith(('-wal','-shm'))})
        self.assertTrue((child/'inputs/amendments/rule262/sources/raw.html').is_file())
        self.assertTrue((child/'inputs/amendments/rule262/sources/index.html').is_file())
        c=common.connect(child/'state.sqlite')
        self.assertEqual([tuple(r)for r in c.execute('SELECT id,review_status,result FROM candidates ORDER BY id')],
            [('101','pending',None),('102','pending',None)])
        self.assertEqual(c.execute("SELECT error FROM candidates WHERE id='102'").fetchone()[0],'Плохой скан')
        self.assertEqual([tuple(r)for r in c.execute('SELECT start,end FROM portable_receipts ORDER BY start')],[(0,10),(10,len(self.card_text()))])
        self.assertEqual(c.execute("SELECT worker FROM portable_claims WHERE id='101'").fetchone()[0],'amendment')
        self.assertEqual(c.execute("SELECT worker FROM portable_claims WHERE id='102'").fetchone()[0],'pending-worker')
        c.close()
        migrated=json.loads((child/'migrated_cards/101.jsonl').read_text())
        self.assertEqual(migrated['reviews'],self.card['reviews'])
        self.assertIn('amendment_provenance',migrated)
        amendments.verify_amendment_run(child,migrated)
        self.assertTrue(batch.amend_archive('parent',self.registry_file,self.snapshot)['resumed'])
        self.assertEqual(common.active(),child)
        changed=dict(migrated,comment='Новые произвольные факты')
        with self.assertRaisesRegex(ValueError,'факты/баллы'):amendments.verify_amendment_run(child,changed)

    def card_text(self):
        return (self.root/'cvs/101.txt').read_text()

    def test_migration_rejects_wrong_engine_missing_reading_and_scope(self):
        parent=self.make_parent();(self.snapshot/'bans.py').write_text('changed',encoding='utf-8')
        with self.assertRaisesRegex(ValueError,'движка'):batch.amend_archive('parent',self.registry_file,self.snapshot)
        shutil.rmtree(self.snapshot);shutil.copytree(Path(service.__file__).parent/'engine',self.snapshot)
        c=common.connect(parent/'state.sqlite')
        with c:c.execute('DELETE FROM portable_receipts WHERE start=10')
        c.close()
        with self.assertRaisesRegex(ValueError,'чтение'):batch.amend_archive('parent',self.registry_file,self.snapshot)
        self.assertEqual(common.active(),parent)
        manifest=common.read(parent/'archive_manifest.json');manifest['ids'].append('999');common.dump(parent/'archive_manifest.json',manifest)
        with self.assertRaisesRegex(ValueError,'манифестом'):batch.amend_archive('parent',self.registry_file,self.snapshot)

    def bans(self,run):
        rule={'row':262,'name':'Резиденты ОЭЗ и партнёры нашей компании','stage':'6 месяцев',
            'months':6,'generic':True,'aliases':[]}
        base=SimpleNamespace(run=run,as_of=date(2026,10,22),company_rules=[rule],company_patterns={},
            people=[],meta={'input_hashes':{'banlist/Компании.csv':'synthetic-rule-sha'}},
            identities=lambda *args:([],set(),[]))
        with patch('app.engine.pilot150_bans.BanEngine',return_value=base):
            return StructuredBans(run)

    def test_verified_relationship_262_is_probable_and_never_guessed_from_title_or_client(self):
        parent=self.make_parent();batch.amend_archive('parent',self.registry_file,self.snapshot);run=common.active()
        card=copy.deepcopy(self.card);text=self.card_text();entity=self.registry['entities'][0]
        proof={'job_index':0,'entity_id':'exact','registry_record_sha256':amendments.canonical_hash(entity),
            'source_cv_anchor':'ООО «Точный Партнер»','certainty':'probable'}
        self.assertEqual(self.bans(run).assess(card,[None]*38,text)['ban'],'нет')
        card['employer_relation_evidence']=[proof]
        results=[self.bans(run).assess(card,[None]*38,text)for _ in range(3)]
        self.assertEqual(results[0],results[1]);self.assertEqual(results[1],results[2])
        finding=results[0]['active'][0]
        self.assertEqual((finding['row'],finding['months'],finding['expires'],finding['certainty']),
            (262,6,'2027-02-28','probable'))
        self.assertEqual(finding['relationship_evidence']['entity_id'],'exact')
        old=copy.deepcopy(card);old['jobs'][0][2]='2025-01'
        self.assertEqual(self.bans(run).assess(old,[None]*38,text)['ban'],'нет')
        unknown=copy.deepcopy(card);unknown['jobs'][0][2]=''
        self.assertEqual(self.bans(run).assess(unknown,[None]*38,text)['active'][0]['certainty'],'probable')
        for mutation in ('hash','job','anchor','client'):
            wrong=copy.deepcopy(card)
            if mutation=='hash':wrong['employer_relation_evidence'][0]['registry_record_sha256']='stale'
            if mutation=='job':wrong['employer_relation_evidence'][0]['job_index']=99
            if mutation=='anchor':wrong['employer_relation_evidence'][0]['source_cv_anchor']='Нет в CV'
            if mutation=='client':wrong['jobs'][0][0]='Совсем другой работодатель'
            with self.subTest(mutation=mutation),self.assertRaises(ValueError):self.bans(run).assess(wrong,[None]*38,text)
        (run/'inputs/amendments/rule262/sources/raw.html').write_text('changed')
        with self.assertRaises(ValueError):self.bans(run).assess(card,[None]*38,text)


    def test_new_preview_submit_is_repeatable_and_stricter_rule_wins(self):
        import re
        parent=self.make_parent();batch.amend_archive('parent',self.registry_file,self.snapshot);run=common.active()
        card=json.loads((run/'migrated_cards/101.jsonl').read_text())
        entity=self.registry['entities'][0]
        card['employer_relation_evidence']=[{'job_index':0,'entity_id':'exact',
            'registry_record_sha256':amendments.canonical_hash(entity),
            'source_cv_anchor':'ООО «Точный Партнер»'}]
        bans=self.bans(run)
        with patch('app.service.review_context',return_value=(service.source_rows(run),bans,{})):
            previews=[service.preview(card,'amendment')for _ in range(3)]
            self.assertEqual(previews[0],previews[1]);self.assertEqual(previews[1],previews[2])
            self.assertEqual(previews[0]['ban'],'да')
            self.assertEqual(previews[0]['implementation_hash'],common.read(run/'portable.json')['implementation_hash'])
            self.assertIn('amendment_receipt_sha256',previews[0])
            self.assertIn('source_excerpt',previews[0]['ban_assessment']['active'][0])
            service.submit(card,'amendment')
            self.assertTrue(service.submit(card,'amendment')['already_completed'])
            with patch('app.service.implementation_matches',return_value=False):
                with self.assertRaisesRegex(ValueError,'Движок'):service.preview(card,'amendment')
        c=common.connect(run/'state.sqlite')
        self.assertEqual(c.execute("SELECT COUNT(*) FROM portable_revisions WHERE kind='completed'").fetchone()[0],1)
        c.close()
        indefinite={'row':999,'name':'Точный Партнер','stage':'100%','months':None,
                    'generic':False,'aliases':['точный партнер']}
        bans.base.company_rules.append(indefinite);bans.rules[999]=indefinite
        bans.base.company_patterns[999]=[('точный партнер',re.compile(r'\bточный партнер\b'))]
        strongest=bans.assess(card,[None]*38,self.card_text())['active']
        self.assertEqual(strongest[0]['row'],999)
        self.assertTrue(any(f['row']==262 for f in strongest))


if __name__=='__main__':unittest.main()
