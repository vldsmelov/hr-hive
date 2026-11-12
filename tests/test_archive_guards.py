"""Production archive guards reject original/proof drift before assessment."""
import json
import unittest
from app import common, service, documents, batch
import test_archive_batch as fixtures


class ArchiveGuardTest(unittest.TestCase):
    def setUp(self):
        self.fixture=fixtures.ArchiveBatchTest();self.fixture.setUp()
        self.run=self.fixture.imported()
        self.connection=common.connect(self.run/'state.sqlite')
        self.text='Полный исходный текст резюме'
        self.doc={'complete':True,'extraction_version':3,'metadata':{},
                  'pages':[{'page':1,'method':'text','text':self.text}]}
        with self.connection:
            self.connection.execute('UPDATE candidates SET text=?,documents=?,extraction_status=? WHERE id=?',
                                    (self.text,json.dumps([self.doc]),'ready','103'))
        service.claim('reviewer',5);shown=service.show('103')
        self.card={'id':'103','reviewed_text_sha256':shown['text_sha256'],
                   'full_text_review_complete':True,'jobs':[],'comment':'Полный разбор документа'}

    def tearDown(self):
        self.connection.close();self.fixture.tearDown()

    def test_changed_original_is_rejected_without_scan(self):
        row=self.connection.execute('SELECT files FROM candidates WHERE id=?',('103',)).fetchone()
        file=json.loads(row['files'])[0]
        (self.fixture.root/'cvs'/file['path']).write_text('Изменённое резюме',encoding='utf-8')
        with self.assertRaisesRegex(ValueError,'оригинал'):service.preview(self.card,'reviewer')

    def test_revoked_visual_review_blocks_submission(self):
        self.doc['metadata']={'visual_review':[{'ocr':False,'image_sha256':'missing-proof','proof':{}}]}
        with self.connection:
            self.connection.execute('UPDATE candidates SET documents=? WHERE id=?',(json.dumps([self.doc]),'103'))
        with self.assertRaisesRegex(ValueError,'визуальная'):service.submit(self.card,'reviewer')
        self.assertEqual(self.connection.execute('SELECT review_status FROM candidates WHERE id=?',('103',)).fetchone()[0],'pending')

    def incomplete_report(self):
        file=json.loads(self.connection.execute('SELECT files FROM candidates WHERE id=?',('103',)).fetchone()[0])[0]
        report={'id':'103','source_file':file['path'],'sha256':file['sha256'],
                'reviewed_text_sha256':self.card['reviewed_text_sha256'],
                'full_text_review_complete':False,'source_visual_review_complete':True,
                'reason':'Исходник содержит только четвёртую страницу из пяти; имя и начальная часть трудовой истории отсутствуют.',
                'reviewer':'reviewer'}
        path=self.fixture.root/'incomplete_report.json';common.dump(path,report)
        return path,file

    def test_incomplete_original_is_not_reviewed_and_rejection_is_idempotent(self):
        path,file=self.incomplete_report()
        self.assertTrue(service.reject_source('103',path,'reviewer')['changed'])
        self.assertFalse(service.reject_source('103',path,'reviewer')['changed'])
        row=self.connection.execute('SELECT * FROM candidates WHERE id=?',('103',)).fetchone()
        self.assertEqual(row['review_status'],'pending');self.assertEqual(row['extraction_status'],'error')
        self.assertIsNone(row['result'])
        self.assertEqual(self.connection.execute("SELECT count(*) FROM portable_revisions WHERE kind='source_incomplete'").fetchone()[0],1)
        with self.assertRaisesRegex(ValueError,'Текст не готов'):service.submit(self.card,'reviewer')

    def test_source_block_survives_ready_cache_retry_and_blocks_stale_ready_assessment(self):
        path,file=self.incomplete_report();service.reject_source('103',path,'reviewer')
        original=self.fixture.root/'cvs'/file['path']
        # A prior, seemingly complete extraction must not bypass a visual block.
        cached={**self.doc,'sha256':file['sha256']}
        common.dump(self.fixture.root/'state/text_cache'/(file['sha256']+'.json'),cached)
        for use_cache in (True,False):
            with self.assertRaisesRegex(ValueError,'Исходник неполон'):documents.extract_file(original,use_cache)
        service.extract(retry=True,ids=['103'])
        self.assertEqual(self.connection.execute('SELECT extraction_status FROM candidates WHERE id=?',('103',)).fetchone()[0],'error')
        with self.connection:self.connection.execute("UPDATE candidates SET extraction_status='ready' WHERE id='103'")
        with self.assertRaisesRegex(ValueError,'Исходник неполон'):service.preview(self.card,'reviewer')
        changed=self.fixture.root/'other.txt';changed.write_text('Другой полный документ без блокировки',encoding='utf-8')
        self.assertTrue(documents.extract_file(changed)['complete'])

    def test_rejection_requires_correct_original_hash_and_keeps_completed_records(self):
        path,file=self.incomplete_report();report=common.read(path)
        report['sha256']='0'*64;common.dump(path,report)
        with self.assertRaisesRegex(ValueError,'хешу'):service.reject_source('103',path,'reviewer')
        self.assertFalse((self.fixture.root/'state/source_blocks').exists())
        report['sha256']=file['sha256'];common.dump(path,report)
        with self.connection:self.connection.execute("UPDATE candidates SET review_status='reviewed' WHERE id='103'")
        with self.assertRaisesRegex(ValueError,'Завершённый'):service.reject_source('103',path,'reviewer')
        self.assertEqual(self.connection.execute('SELECT extraction_status FROM candidates WHERE id=?',('103',)).fetchone()[0],'ready')

    def test_wrong_document_can_be_blocked_after_full_text_reading_only(self):
        path,file=self.incomplete_report();report=common.read(path)
        report.update(source_visual_review_complete=False,source_document_review_complete=True,
                      reason='Вместо резюме приложен отчёт юридического лица; установить кандидата и его трудовую историю невозможно.',
                      read_files=[file['path']],read_parts=[[0,len(self.text)-1]])
        common.dump(path,report)
        with self.assertRaisesRegex(ValueError,'полное текстовое'):service.reject_source('103',path,'reviewer')
        report['read_parts']=[[0,len(self.text)]];common.dump(path,report)
        self.assertTrue(service.reject_source('103',path,'reviewer')['changed'])
        self.assertEqual(self.connection.execute('SELECT review_status FROM candidates WHERE id=?',('103',)).fetchone()[0],'pending')

    def test_identity_conflict_blocks_candidate_without_poisoning_shared_document(self):
        path,file=self.incomplete_report();report=common.read(path)
        report.update(scope='candidate',reason='В комплекте кандидата находится резюме другого человека; тот же файл корректно относится к отдельному ID.')
        common.dump(path,report);service.reject_source('103',path,'reviewer')
        self.assertFalse((self.fixture.root/'state/source_blocks'/(file['sha256']+'.json')).exists())
        original=self.fixture.root/'cvs'/file['path']
        self.assertTrue(documents.extract_file(original,False)['complete'])
        with self.connection:
            self.connection.execute("UPDATE candidates SET files=? WHERE id='102'",(json.dumps([file]),))
        service.extract(retry=True,ids=['103','102'])
        states=dict(self.connection.execute("SELECT id,extraction_status FROM candidates WHERE id IN ('102','103')"))
        self.assertEqual(states,{'102':'ready','103':'error'})
        with self.connection:self.connection.execute("UPDATE candidates SET extraction_status='ready' WHERE id='103'")
        with self.assertRaisesRegex(ValueError,'Исходник неполон'):service.preview(self.card,'reviewer')
        with self.connection:self.connection.execute("UPDATE candidates SET fingerprint='changed-file-set' WHERE id='103'")
        service.require_complete_candidate_source(self.run,self.connection.execute("SELECT * FROM candidates WHERE id='103'").fetchone())

    def test_parallel_retry_respects_candidate_conflict_and_shared_file(self):
        path,file=self.incomplete_report();report=common.read(path)
        report.update(scope='candidate');common.dump(path,report)
        service.reject_source('103',path,'reviewer')
        with self.connection:self.connection.execute("UPDATE candidates SET files=? WHERE id='102'",(json.dumps([file]),))
        result=batch.extract_parallel(workers=1,retry=True,ids=['103','102'])
        self.assertEqual(result,{'attempted':2,'ready':1,'errors':1})
        states=dict(self.connection.execute("SELECT id,extraction_status FROM candidates WHERE id IN ('102','103')"))
        self.assertEqual(states,{'102':'ready','103':'error'})


if __name__=='__main__':unittest.main()
