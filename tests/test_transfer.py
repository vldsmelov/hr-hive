"""Portable backups and explicit criteria use synthetic files only."""
import json
import sqlite3
import tarfile
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook
from app import archives, batch, books, common, service
from tools import transfer_audit, audit_repeatability
import test_archive_batch


class CriteriaVersionTest(unittest.TestCase):
    setUp=test_archive_batch.ArchiveBatchTest.setUp
    tearDown=test_archive_batch.ArchiveBatchTest.tearDown
    write_archive=test_archive_batch.ArchiveBatchTest.write_archive
    add_candidate=test_archive_batch.ArchiveBatchTest.add_candidate

    def criteria(self):
        root=self.root/'criteria-v2'
        (root/'clusters').mkdir(parents=True);(root/'banlist').mkdir()
        book=Workbook();sheet=book.active;sheet.title='Кого берём'
        sheet.append(['ID','Модуль','Кластер/проект','Профиль','Уровень','Образование','Опыт','Навыки','Стоп','Комментарий'])
        for number in range(2,342):
            sheet.append([number,'Модуль','КБ','Профиль '+str(number),'middle','Любое','Любой','Инженерные задачи',None,None])
        book.save(root/'clusters/profiles.xlsx');book.close()
        (root/'banlist/ФИО.csv').write_text('ФИО;Телефон\n',encoding='utf-8')
        (root/'banlist/Компании.csv').write_text('Компания;Срок\n',encoding='utf-8')
        (root/'accepted_rules.md').write_text('Устаревшая политика',encoding='utf-8')
        return root

    def test_explicit_criteria_replace_inheritance_and_resume_by_content(self):
        criteria=self.criteria();previous=common.sha(self.previous/'state.sqlite')
        result=batch.import_archive(self.archive,self.template,self.output,as_of='2026-11-04',criteria_dir=criteria)
        first=common.active()
        self.assertFalse((first/'inputs/overrides.json').exists())
        self.assertEqual(common.sha(first/'inputs/accepted_rules.md'),common.sha(Path(service.__file__).parent.parent/'rules/accepted_rules.md'))
        self.assertEqual(common.sha(first/'inputs/clusters/profiles.xlsx'),common.sha(criteria/'clusters/profiles.xlsx'))
        connection=common.connect(first/'state.sqlite')
        self.assertEqual(connection.execute('SELECT COUNT(*) FROM candidates WHERE result IS NOT NULL').fetchone()[0],0)
        connection.close()
        again=batch.import_archive(self.archive,self.template,self.output,as_of='2026-11-04',criteria_dir=criteria)
        self.assertTrue(again['resumed']);self.assertEqual(result['run'],again['run'])
        (criteria/'banlist/ФИО.csv').write_text('ФИО;Телефон\nТестовый Иван Иванович;\n',encoding='utf-8')
        changed=batch.import_archive(self.archive,self.template,self.output,as_of='2026-11-04',criteria_dir=criteria)
        self.assertNotEqual(result['run'],changed['run'])
        self.assertNotEqual(common.read(first/'portable.json')['rules_hash'],common.read(common.active()/'portable.json')['rules_hash'])
        self.assertEqual(common.sha(self.previous/'state.sqlite'),previous)

    def test_invalid_criteria_do_not_switch_or_publish_run(self):
        criteria=self.criteria();(criteria/'banlist/ФИО.csv').unlink()
        before=common.read(self.root/'state/active.json')
        runs=list((self.root/'state/runs').iterdir())
        with self.assertRaisesRegex(ValueError,'бан-листы'):
            batch.import_archive(self.archive,self.template,criteria_dir=criteria)
        self.assertEqual(before,common.read(self.root/'state/active.json'))
        self.assertEqual(runs,list((self.root/'state/runs').iterdir()))


class TransferBackupTest(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory();self.root=Path(self.temp.name)/'data';self.root.mkdir()

    def tearDown(self):self.temp.cleanup()

    def write(self,name,value=b'content'):
        path=self.root/name;path.parent.mkdir(parents=True,exist_ok=True);path.write_bytes(value);return path

    def test_backup_roles_keep_rule_proofs_and_strip_cv_derivatives(self):
        self.write('cvs/101.pdf',b'%PDF ORIGINAL')
        self.write('state/audit/renders/101/page.png',b'\x89PNG original render')
        self.write('state/inspection/101',b'%PDF no extension')
        self.write('state/copies/101.xlsx',b'questionnaire')
        self.write('state/audit/card.json',b'{"full_text":"stored text"}')
        proof=self.write('state/audit/organization/official.pdf',b'%PDF official source')
        self.write('state/runs/test/inputs/amendments/rule262/official.pdf',b'%PDF frozen proof')
        self.write('sources/clusters/profiles.xlsx')
        self.write('exports/101.xlsx',b'export')
        self.write('backups/old.tar.gz')
        self.write('state/lo_profile/cache/blob.bin')
        common.dump(self.root/'state/backup_evidence.json',{'files':[{
            'path':proof.relative_to(self.root).as_posix(),'sha256':common.sha(proof),'role':'rule_evidence'}]})
        database=self.root/'state/old.sqlite';connection=sqlite3.connect(database)
        connection.executescript('CREATE TABLE source_files(hash TEXT,bytes BLOB); CREATE TABLE results(id TEXT,value TEXT);')
        connection.execute('INSERT INTO source_files VALUES(?,?)',('cv',b'original'))
        connection.execute('INSERT INTO results VALUES(?,?)',('101','saved'));connection.commit();connection.close()
        baseline=transfer_audit.database(database)
        target=Path(self.temp.name)/'backup.tar.gz';archives.archive_tree(self.root,target)
        with tarfile.open(target) as archive:
            names=archive.getnames();manifest=json.load(archive.extractfile('BACKUP_MANIFEST.json'))
        self.assertIn('state/audit/organization/official.pdf',names)
        self.assertIn('state/runs/test/inputs/amendments/rule262/official.pdf',names)
        self.assertIn('exports/101.xlsx',names)
        self.assertIn('state/audit/card.json',names)
        for excluded in ('cvs/101.pdf','state/audit/renders/101/page.png','state/inspection/101',
                         'state/copies/101.xlsx','backups/old.tar.gz','state/lo_profile/cache/blob.bin'):
            self.assertNotIn(excluded,names)
        self.assertGreater(manifest['excluded_counts']['cv_binary_or_render'],0)
        restored=Path(self.temp.name)/'restore';archives.restore(target,restored)
        self.assertEqual(baseline,transfer_audit.database(restored/'state/old.sqlite'))
        self.assertEqual(common.sha(proof),common.sha(restored/proof.relative_to(self.root)))
        with self.assertRaisesRegex(ValueError,'существует'):archives.archive_tree(self.root,target)

    def test_evidence_hash_mismatch_blocks_backup(self):
        self.write('state/audit/proof.pdf',b'%PDF proof')
        common.dump(self.root/'state/backup_evidence.json',{'files':[{
            'path':'state/audit/proof.pdf','sha256':'wrong','role':'rule_evidence'}]})
        with self.assertRaisesRegex(ValueError,'подтверждения'):
            archives.archive_tree(self.root,Path(self.temp.name)/'invalid.tar.gz')

    def test_diagnostic_exclusion_is_hash_bound_and_cannot_hide_working_database(self):
        path=self.write('state/recovery_scratch/state.sqlite',b'broken old scratch')
        record={'path':path.relative_to(self.root).as_posix(),'sha256':common.sha(path),
                'role':'diagnostic_copy','reason':'Старая повреждённая временная копия'}
        inventory=self.root/'state/backup_exclusions.json'
        common.dump(inventory,{'files':[record]})
        archive=Path(self.temp.name)/'backup.tar.gz';archives.archive_tree(self.root,archive)
        with tarfile.open(archive) as contents:
            self.assertNotIn(record['path'],contents.getnames())
            self.assertEqual(json.load(contents.extractfile('BACKUP_MANIFEST.json'))['excluded_diagnostics'],[record])
        restored=Path(self.temp.name)/'restored'
        archives.restore(archive,restored)
        archives.archive_tree(restored,Path(self.temp.name)/'next-backup.tar.gz')
        path.write_bytes(b'changed scratch')
        with self.assertRaisesRegex(ValueError,'диагностическая копия'):
            archives.archive_tree(self.root,Path(self.temp.name)/'changed-backup.tar.gz')
        record['path']='state/runs/active/state.sqlite'
        common.dump(inventory,{'files':[record]})
        with self.assertRaisesRegex(ValueError,'рабочие данные'):archives.diagnostic_inventory(self.root)

    def test_corrupt_archive_is_rejected_before_files_are_restored(self):
        self.write('state/example.json');target=Path(self.temp.name)/'backup.tar.gz'
        archives.archive_tree(self.root,target)
        import io
        with tarfile.open(target) as archive:
            records=[(member,archive.extractfile(member).read()) for member in archive.getmembers()]
        damaged=Path(self.temp.name)/'damaged.tar.gz'
        with tarfile.open(damaged,'w:gz') as archive:
            for member,payload in records:
                if member.name=='state/example.json':payload=b'changed';member.size=len(payload)
                archive.addfile(member,io.BytesIO(payload))
        restored=Path(self.temp.name)/'restore'
        with self.assertRaisesRegex(ValueError,'контрольная сумма'):archives.restore(damaged,restored)
        self.assertEqual(list(restored.iterdir()),[])


class PortableAuditTest(unittest.TestCase):
    setUp=test_archive_batch.ArchiveBatchTest.setUp
    tearDown=test_archive_batch.ArchiveBatchTest.tearDown
    write_archive=test_archive_batch.ArchiveBatchTest.write_archive
    add_candidate=test_archive_batch.ArchiveBatchTest.add_candidate
    append_run=test_archive_batch.ArchiveBatchTest.append_run

    def test_reference_understands_append_layout_and_ignores_unfinished_rows(self):
        run=self.append_run();self.add_candidate(run,'101',3,review='reviewed')
        self.add_candidate(run,'102',4,extraction='error',error='Неполный документ')
        exported=books.export(run)
        rows=audit_repeatability.reference_rows(exported['path'])
        self.assertEqual(set(rows),{'101'});self.assertEqual(rows['101']['department'],'УПИ 88')

    def test_replay_copy_includes_exact_originals_and_global_source_guards(self):
        batch.import_archive(self.archive,self.template,self.output)
        source=common.active();before=common.sha(source/'state.sqlite')
        common.dump(self.root/'state/image_reviews.json',{'images':{}})
        common.dump(self.root/'state/page_reviews.json',{'documents':{}})
        common.dump(self.root/'state/source_blocks/example.json',{'reason':'Неполный источник'})
        common.dump(source/'candidate_source_blocks/103.json',{'reason':'Конфликт комплекта'})
        destination=self.root/'isolated'
        with patch.object(audit_repeatability,'DATA',self.root):
            audit_repeatability.clone_run(source,destination,['101'])
        self.assertEqual(common.sha(source/'state.sqlite'),before)
        for name in ('image_reviews.json','page_reviews.json','source_blocks/example.json'):
            self.assertEqual(common.sha(self.root/'state'/name),common.sha(destination/'state'/name))
        self.assertTrue((destination/'state/runs'/source.name/'candidate_source_blocks/103.json').exists())
        original=self.root/'cvs'/source.name/'folder/101.txt'
        self.assertEqual(common.sha(original),common.sha(destination/'cvs'/source.name/'folder/101.txt'))
        self.assertFalse((destination/'cvs'/source.name/'folder/102.txt').exists())
