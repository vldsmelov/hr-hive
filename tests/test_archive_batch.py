"""Synthetic archive/append-layout regressions; never use production data."""
import copy
import hashlib
import io
import json
import tempfile
import unittest
import zipfile
from datetime import datetime
from pathlib import Path
from unittest.mock import patch
from zoneinfo import ZoneInfo

import fitz
from PIL import Image
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill

from app import archives, batch, books, common, documents, service, parallel_extract
from app.engine.profiles import assess


class ArchiveBatchTest(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.patches = [patch.object(module, 'DATA', self.root)
                        for module in (common, documents, service, books, archives, batch)
                        if hasattr(module, 'DATA')]
        for item in self.patches:
            item.start()
        common.layout()
        book = Workbook()
        sheet = book.active
        sheet.title = 'Прежние результаты'
        sheet.append(['ИИ', 'ID'] + ['Поле'] * 29)
        sheet.append([None, '101'] + ['Прежнее значение'] * 29)
        book.save(self.root / 'sources' / 'Матка №1.xlsx')
        book.close()
        self.previous = service.new_run()
        self.add_candidate(self.previous, '101', 2, 'ready', 'reviewed')
        (self.previous / 'inputs' / 'overrides.json').write_text(
            '{"note":"Дополнение действующей версии"}', encoding='utf-8')
        meta = common.read(self.previous / 'portable.json')
        meta['rules_hash'] = service.rules_hash(self.previous / 'inputs')
        common.dump(self.previous / 'portable.json', meta)
        self.template = self.root / 'exl1.xlsx'
        book = Workbook()
        sheet = book.active
        sheet.title = 'Новая анкета'
        sheet.append(['ID'] + ['Исходное поле ' + str(n) for n in range(2, 39)])
        # Deliberately not sorted by ID: output must follow the source order.
        for ident in ('103', '101', '102', '999'):
            sheet.append([ident] + ['Исходное ' + ident] * 37)
        sheet['C3'] = '=1+2'
        sheet['D3'].hyperlink = 'https://example.test/кандидат-101'
        sheet['D3'].number_format = '@'
        sheet['D3'].fill = PatternFill('solid', fgColor='00FF00')
        sheet['AL3'] = 'Сохранить ручной комментарий'
        sheet['D5'].hyperlink = 'https://example.test/не-входит-в-пакет'
        book.save(self.template)
        book.close()
        self.archive = self.root / 'newCVs.zip'
        self.write_archive({'folder/101.txt': 'Первый документ кандидата 101',
                            'folder/101.md': 'Второй документ кандидата 101',
                            'folder/102.txt': 'Документ кандидата 102',
                            'folder/103.txt': 'Документ кандидата 103'})
        self.output = self.root / 'exports' / 'newCVs.xlsx'

    def tearDown(self):
        for item in reversed(self.patches):
            item.stop()
        self.temp.cleanup()

    def write_archive(self, entries):
        with zipfile.ZipFile(self.archive, 'w', zipfile.ZIP_STORED) as archive:
            for name, value in entries.items():
                archive.writestr(name, value)

    def add_candidate(self, run, ident, row, extraction='ready', review='pending', error=None):
        result = {'id': ident, 'ai': 'да', 'department': 'УПИ 88',
                  'comment': 'Подтверждены продажи и переговоры', 'ban': 'нет',
                  'ban_comment': '', 'clusters': [{'cluster': 'УПИ', 'score': 88}]}
        connection = common.connect(run / 'state.sqlite')
        with connection:
            connection.execute(
                'INSERT INTO candidates(id,source_row,files,fingerprint,extraction_status,'
                'text,error,review_status,result) VALUES(?,?,?,?,?,?,?,?,?)',
                (ident, row, '[]', 'synthetic-hash', extraction,
                 'Полная трудовая история кандидата.\n' * 100, error, review,
                 json.dumps(result, ensure_ascii=False) if review == 'reviewed' else None))
        connection.close()

    def imported(self):
        batch.import_archive(self.archive, self.template, self.output)
        return common.active()

    def append_run(self):
        # Exporter/claims tests intentionally do not depend on archive import.
        run = self.root / 'state' / 'runs' / 'append-test'
        run.mkdir()
        (run / 'inputs').mkdir()
        import shutil
        shutil.copy2(self.template, run / 'inputs' / 'template.xlsx')
        common.dump(run / 'portable.json', {
            'source_book': 'inputs/template.xlsx', 'sheet': 'Новая анкета',
            'output_name': 'append-test.xlsx', 'id_column': 1,
            'source_column_count': 38, 'export_layout': 'append_results',
            'rules_hash': service.rules_hash(run / 'inputs'),
            'implementation_hash': service.implementation_hash()})
        common.dump(run / 'manifest.json', {'as_of': '2026-10-22', 'input_hashes': {}})
        source = common.connect(self.previous / 'state.sqlite')
        target = common.connect(run / 'state.sqlite')
        source.backup(target)
        source.close()
        with target:
            target.execute('DELETE FROM candidates')
        target.close()
        common.dump(self.root / 'state' / 'active.json', {'run': run.name})
        service.prepare(run)
        return run

    def test_import_scopes_ids_groups_files_and_keeps_previous_results(self):
        previous_hash = common.sha(self.previous / 'state.sqlite')
        source_hash = common.sha(self.template)
        # A global scan would accidentally pick up 999; scoped import must not.
        (self.root / 'cvs' / '999.txt').write_text('Постороннее резюме', encoding='utf-8')
        run = self.imported()
        self.assertNotEqual(run, self.previous)
        meta = common.read(run / 'portable.json')
        self.assertEqual((meta['id_column'], meta['source_column_count'], meta['export_layout']),
                         (1, 38, 'append_results'))
        self.assertEqual(common.read(run / 'manifest.json')['as_of'],
                         datetime.now(ZoneInfo('Europe/Moscow')).date().isoformat())
        self.assertTrue((run / 'inputs' / 'overrides.json').exists())
        self.assertTrue(any('accepted_rules' in p.name for p in (run / 'inputs').rglob('*')))
        connection = common.connect(run / 'state.sqlite')
        rows = connection.execute('SELECT * FROM candidates ORDER BY source_row').fetchall()
        self.assertEqual([r['id'] for r in rows], ['103', '101', '102'])
        self.assertEqual(len(json.loads(rows[1]['files'])), 2)
        self.assertTrue(all(r['review_status'] == 'pending' and r['result'] is None for r in rows))
        self.assertEqual(connection.execute('SELECT COUNT(*) FROM fact_cards').fetchone()[0], 0)
        connection.close()
        self.assertEqual(common.sha(self.previous / 'state.sqlite'), previous_hash)
        self.assertEqual(common.sha(self.template), source_hash)
        self.assertEqual(service.source_rows(run)['101'][0], 3)

    def test_export_wording_correction_preserves_assessment_and_rejects_stale_source(self):
        run = self.append_run()
        self.add_candidate(run, '101', 3, review='reviewed')
        connection = common.connect(run / 'state.sqlite')
        original = connection.execute('SELECT result FROM candidates WHERE id=?', ('101',)).fetchone()[0]
        connection.close()
        comment = json.loads(original)['comment']
        common.dump(run / 'export_comments.json', {
            '101': {'source_sha256': hashlib.sha256(comment.encode()).hexdigest(),
                    'text': 'Продажи подтверждены; условия работы требуют уточнения.'}})
        output = books.export(run)
        workbook = load_workbook(output['path'])
        self.assertEqual(workbook.active['AO2'].value,
                         'Продажи подтверждены; условия работы требуют уточнения.')
        workbook.close()
        connection = common.connect(run / 'state.sqlite')
        self.assertEqual(connection.execute('SELECT result FROM candidates WHERE id=?', ('101',)).fetchone()[0], original)
        connection.close()
        common.dump(run / 'export_comments.json', {
            '101': {'source_sha256': 'different source', 'text': 'Не применять'}})
        with self.assertRaisesRegex(ValueError, 'Исходный комментарий изменился'):
            books.export(run)

    def test_import_resumes_same_archive_without_erasing_work(self):
        run = self.imported()
        connection = common.connect(run / 'state.sqlite')
        with connection:
            connection.execute("UPDATE candidates SET extraction_status='error',error='Проверочная ошибка' WHERE id='101'")
        connection.close()
        self.imported()
        self.assertEqual(common.active(), run)
        connection = common.connect(run / 'state.sqlite')
        self.assertEqual(connection.execute("SELECT error FROM candidates WHERE id='101'").fetchone()[0],
                         'Проверочная ошибка')
        self.assertEqual(connection.execute('SELECT COUNT(*) FROM candidates').fetchone()[0], 3)
        connection.close()

    def test_incomplete_import_is_not_resumed_and_changed_original_is_rejected(self):
        incomplete = self.imported()
        meta = common.read(incomplete / 'portable.json')
        meta['import_complete'] = False
        common.dump(incomplete / 'portable.json', meta)
        # Simulate the interruption before schema creation, without touching
        # any real run. The remaining inputs are enough to create a fresh run.
        (incomplete / 'state.sqlite').unlink()
        complete = self.imported()
        self.assertNotEqual(complete, incomplete)
        self.assertTrue(common.read(complete / 'portable.json')['import_complete'])
        inventory = common.read(complete / 'archive_manifest.json')['files']
        original = self.root / 'cvs' / inventory[0]['path']
        original.write_text('Содержимое изменилось после импорта', encoding='utf-8')
        with self.assertRaisesRegex(ValueError, 'оригинал'):
            self.imported()
        self.assertEqual(common.active(), complete)

    def test_archive_validation_rejects_unsafe_paths_crc_and_missing_ids(self):
        for entries in ({'../escape.txt': 'Нельзя распаковывать'},
                        {'folder/888.txt': 'Нет строки в шаблоне'}):
            with self.subTest(entries=entries):
                self.write_archive(entries)
                before = common.read(self.root / 'state' / 'active.json')
                with self.assertRaises((ValueError, zipfile.BadZipFile)):
                    self.imported()
                self.assertEqual(common.read(self.root / 'state' / 'active.json'), before)
                self.assertFalse((self.root / 'escape.txt').exists())
        self.write_archive({'101.txt': b'unique payload for crc validation'})
        raw = self.archive.read_bytes()
        self.archive.write_bytes(raw.replace(b'unique payload', b'broken payload', 1))
        with self.assertRaises((ValueError, zipfile.BadZipFile)):
            self.imported()
        self.assertEqual(common.active(), self.previous)

    def test_duplicate_template_id_is_rejected_before_activation(self):
        workbook = load_workbook(self.template)
        workbook.active['A3'] = '103'
        workbook.save(self.template)
        workbook.close()
        with self.assertRaisesRegex(ValueError, 'ID'):
            self.imported()
        self.assertEqual(common.active(), self.previous)

    def test_canonical_duplicate_zip_paths_are_rejected_before_writing(self):
        self.write_archive({'folder/101.txt': 'Один документ',
                            './folder/101.txt': 'Подмена того же документа'})
        before = set((self.root / 'cvs').rglob('*'))
        with self.assertRaisesRegex(ValueError, 'Повторный путь'):
            self.imported()
        self.assertEqual(common.active(), self.previous)
        self.assertEqual(set((self.root / 'cvs').rglob('*')), before)

    def test_append_export_preserves_38_fields_and_includes_unfinished(self):
        run = self.append_run()
        self.add_candidate(run, '103', 2, 'error', error='OCR не прочитал вторую страницу')
        self.add_candidate(run, '101', 3, 'ready', 'reviewed')
        self.add_candidate(run, '102', 4)
        result = books.export(run)
        self.assertEqual(result['rows'], 3)
        source = load_workbook(self.template)
        exported = load_workbook(result['path'])
        try:
            self.assertEqual(exported.sheetnames, ['Новая анкета'])
            sheet = exported.active
            self.assertEqual((sheet.max_row, sheet.max_column), (4, 44))
            self.assertEqual([sheet.cell(r, 1).value for r in range(2, 5)], ['103', '101', '102'])
            self.assertEqual([sheet.cell(1, c).value for c in range(39, 45)],
                             ['ИИ', 'Подразделение (HR Hive)', 'Комментарий (HR Hive)',
                              'Бан', 'Комментарий к бану', 'Статус обработки'])
            for row in range(1, 5):
                for col in range(1, 39):
                    a, b = source.active.cell(row, col), sheet.cell(row, col)
                    self.assertEqual(a.value, b.value, (row, col))
                    self.assertEqual(a.number_format, b.number_format, (row, col))
                    self.assertEqual(copy.copy(a.fill), copy.copy(b.fill), (row, col))
                    self.assertEqual(a.hyperlink.target if a.hyperlink else None,
                                     b.hyperlink.target if b.hyperlink else None)
            self.assertEqual(sheet['AM3'].value, 'да')
            self.assertEqual(sheet['AN3'].value, 'УПИ 88')
            self.assertEqual(sheet['AP3'].value, 'нет')
            self.assertEqual(sheet['AR3'].value, 'Завершено')
            self.assertIn('вторую страницу', sheet['AR2'].value)
            self.assertEqual(sheet['AR4'].value, 'Ожидает проверки')
            for row in (2, 4):
                self.assertTrue(all(sheet.cell(row, col).value in (None, '') for col in range(39, 44)))
        finally:
            source.close()
            exported.close()
        with zipfile.ZipFile(result['path']) as archive:
            self.assertNotIn('не-входит-в-пакет'.encode(), b''.join(archive.read(n) for n in archive.namelist()))
        connection = common.connect(run / 'state.sqlite')
        rows = connection.execute('SELECT * FROM candidates ORDER BY source_row').fetchall()
        connection.close()
        tampered = load_workbook(result['path'])
        tampered.active['AL3'] = 'Потерян ручной комментарий'
        tampered.save(result['path'])
        tampered.close()
        with self.assertRaises(AssertionError):
            books.verify(result['path'], self.template, 'Новая анкета', rows,
                         common.read(run / 'portable.json'))

    def test_workers_have_disjoint_resumable_claims_and_receipts_require_all_chunks(self):
        run = self.append_run()
        for row, ident in enumerate(('103', '101', '102'), 2):
            self.add_candidate(run, ident, row)
        self.assertEqual(service.claim('worker-1', 1), ['103'])
        self.assertEqual(service.claim('worker-2', 1), ['101'])
        self.assertEqual(service.claim('worker-1', 1), ['103'])
        self.assertEqual(service.claim('worker-3', 5), ['102'])
        shown = service.show('103', 0, 25)
        card = {'id': '103', 'full_text_review_complete': True,
                'reviewed_text_sha256': shown['text_sha256'],
                'comment': 'Нет подтверждённого профильного опыта', 'jobs': [],
                'tracks': {}, 'reviews': []}
        with self.assertRaisesRegex(ValueError, 'весь текст'):
            service.preview(card, 'worker-1')
        offset = shown['next_offset']
        while offset is not None:
            shown = service.show('103', offset, 100)
            offset = shown['next_offset']
        with self.assertRaisesRegex(ValueError, 'worker'):
            service.preview(card, 'worker-2')
        wrong = dict(card, reviewed_text_sha256='outdated')
        with self.assertRaisesRegex(ValueError, 'sha256'):
            service.preview(wrong, 'worker-1')

    def test_three_fixed_card_previews_and_one_immutable_submission(self):
        run = self.append_run()
        self.add_candidate(run, '101', 3)
        service.claim('worker-1', 1)
        shown = service.show('101', 0, 20000)
        self.assertIsNone(shown['next_offset'])
        card = {'id': '101', 'name': 'Синтетический Кандидат', 'phones': [],
                'education': 'Подтверждённая квалификация', 'jobs': [],
                'full_text_review_complete': True,
                'reviewed_text_sha256': shown['text_sha256'],
                'comment': 'Подтверждены инженерная и коммерческая специализации',
                'tracks': {family: ['specialist', 3, 'Трудовая история', 'Подтверждённые задачи']
                           for family in ('engineering', 'sales')},
                'reviews': [['A', [36, 27, 18, 9], 'engineering', 'Подтверждённые задачи', ''],
                            ['B', [32, 24, 16, 8], 'sales', 'Подтверждённые задачи', '']],
                'independent_directions': [['engineering'], ['sales']]}
        profiles = {key: {'cluster': cluster, 'profile': 'Профиль',
                          'education': 'Профильное', 'experience': '1-3 года'}
                    for key, cluster in (('A', 'Инженерный'), ('B', 'Коммерческий'))}
        original = copy.deepcopy(card)
        with patch('app.service.StructuredBans') as bans, patch('app.service.catalog', return_value=profiles):
            bans.return_value.assess.side_effect = lambda *args: {'ban': 'нет', 'findings': [], 'active': []}
            results = [service.preview(card, 'worker-1') for _ in range(3)]
            self.assertEqual(results[0], results[1])
            self.assertEqual(results[1], results[2])
            self.assertEqual(results[0]['clusters'][0], {'cluster': '100 Лидеров', 'score': 92})
            service.submit(card, 'worker-1')
            self.assertTrue(service.submit(card, 'worker-1')['already_completed'])
            changed = dict(card, comment='Новый комментарий к завершённой оценке')
            with self.assertRaisesRegex(ValueError, 'Завершённая проверка'):
                service.submit(changed, 'worker-1')
        self.assertEqual(card, original)
        connection = common.connect(run / 'state.sqlite')
        self.assertEqual(connection.execute("SELECT COUNT(*) FROM portable_revisions WHERE kind='completed'").fetchone()[0], 1)
        self.assertEqual(connection.execute('SELECT COUNT(*) FROM fact_cards').fetchone()[0], 1)
        saved = json.loads(connection.execute("SELECT result FROM candidates WHERE id='101'").fetchone()[0])
        connection.close()
        self.assertEqual(saved, results[0])

    def test_direct_100_leaders_profile_cannot_bypass_independent_professions(self):
        profiles = {'N003': {'cluster': '100 Лидеров', 'profile': 'Все профили',
                             'education': 'Любое', 'experience': 'Любой'}}
        card = {'id': 'synthetic',
                'tracks': {'generic': ['specialist', 1, 'Факт из резюме', 'Одна профессия']},
                'reviews': [['N003', [40, 30, 20, 10], 'generic', 'Подтверждённый факт', '']]}
        for directions in ([], [['generic']], [['generic'], ['generic']]):
            with self.subTest(directions=directions):
                card['independent_directions'] = directions
                assessed, rejected, chosen = assess(card, profiles)
                self.assertEqual(assessed, [])
                self.assertEqual(chosen, [])
                self.assertTrue(any(item['profile'] == 'N003' for item in rejected))

    def test_parallel_file_workers_extract_aliases_and_errors_without_sqlite(self):
        first = self.root / 'first.txt'
        first.write_text('Полная трудовая история кандидата: монтаж и обслуживание.', encoding='utf-8')
        alias = self.root / 'same.md'
        alias.write_bytes(first.read_bytes())
        second = self.root / 'second.txt'
        second.write_text('Другой кандидат выполняет сборку и контроль качества.', encoding='utf-8')
        bad = self.root / 'broken.bin'
        bad.write_bytes(b'unsupported synthetic binary')
        with patch('sqlite3.connect', side_effect=AssertionError('У рабочих извлечения нет доступа к SQLite')):
            results = list(parallel_extract.iter_extracted([first, alias, second, bad], max_workers=2))
        by_path = {Path(path): (doc, error) for path, doc, error in results}
        self.assertEqual(set(by_path), {first, alias, second, bad})
        for path in (first, alias, second):
            document, error = by_path[path]
            self.assertIsNone(error)
            self.assertTrue(document['complete'])
            self.assertEqual(document['original_name'], path.name)
            self.assertEqual(document['sha256'], common.sha(path))
        self.assertEqual(by_path[first][0]['pages'], by_path[alias][0]['pages'])
        self.assertIsNone(by_path[bad][0])
        self.assertIn('Неподдерживаемый', by_path[bad][1])
        self.assertEqual(len(list((self.root / 'state' / 'text_cache').glob('*.json'))), 2)
        with self.assertRaises(ValueError):
            list(parallel_extract.iter_extracted([first], max_workers=3))

    def test_extensionless_pdf_and_docx_include_all_pages_and_tables(self):
        path = self.root / 'extensionless_pdf'
        pdf = fitz.open()
        for number in (1, 2):
            page = pdf.new_page()
            page.insert_text((60, 60), ('Engineering experience page %s: assembly, maintenance, production. ' % number) * 3)
        pdf.save(path)
        pdf.close()
        extracted = documents.extract_file(path)
        self.assertEqual(len(extracted['pages']), 2)
        self.assertIn('page 2', extracted['pages'][1]['text'])
        path = self.root / 'extensionless_docx'
        with zipfile.ZipFile(path, 'w') as archive:
            archive.writestr('word/document.xml',
                '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                '<w:body><w:p><w:r><w:t>Первая работа</w:t></w:r></w:p>'
                '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>Вторая работа в таблице</w:t>'
                '</w:r></w:p></w:tc></w:tr></w:tbl></w:body></w:document>')
        text = '\n'.join(page['text'] for page in documents.extract_file(path)['pages'])
        self.assertIn('Первая работа', text)
        self.assertIn('Вторая работа в таблице', text)

    def test_xlsx_questionnaire_reads_applicant_and_excludes_hidden_reference_lists(self):
        path = self.root / 'questionnaire.xlsx'
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = 'Анкета'
        sheet.append(['Работодатель', 'Обязанности'])
        sheet.append(['Подтверждённый завод', 'Сборка и контроль качества'])
        reference = workbook.create_sheet('Справочник специальностей')
        reference.sheet_state = 'hidden'
        for number in range(100):
            reference.append(['НЕ_ОПЫТ_КАНДИДАТА_' + str(number)])
        workbook.save(path)
        workbook.close()
        extracted = documents.extract_file(path)
        text = '\n'.join(page['text'] for page in extracted['pages'])
        self.assertIn('Подтверждённый завод', text)
        self.assertIn('Сборка и контроль качества', text)
        self.assertNotIn('НЕ_ОПЫТ_КАНДИДАТА_', text)
        self.assertTrue(extracted['complete'])

    def test_stale_or_incomplete_cache_is_not_reused(self):
        path = self.root / 'complete.txt'
        path.write_text('Полное резюме содержит всю трудовую историю кандидата.', encoding='utf-8')
        fingerprint = common.sha(path)
        cache = self.root / 'state' / 'text_cache' / (fingerprint + '.json')
        for invalid in (
            {'sha256': fingerprint, 'pages': [{'page': 1, 'method': 'text', 'text': 'Усечено'}]},
            {'sha256': fingerprint, 'extraction_version': 2, 'complete': True, 'pages': []},
            {'sha256': fingerprint, 'extraction_version': 2, 'complete': True,
             'pages': [{'page': 1, 'method': 'ocr_pending', 'text': 'Текст ещё не готов'}]},
        ):
            with self.subTest(cache=invalid):
                common.dump(cache, invalid)
                extracted = documents.extract_file(path)
                self.assertEqual(extracted['pages'][0]['text'], path.read_text('utf-8'))
                self.assertTrue(extracted['complete'])

    def test_ocr_rotation_rejects_gibberish_and_failure_does_not_cache(self):
        image = Image.new('RGB', (300, 100), 'white')
        with patch('app.documents.shutil.which', return_value='/synthetic/tesseract'), \
                patch('app.documents._ocr_attempt', side_effect=[
                    ('Неверный поворот', 5), ('Подтверждённая трудовая история', 90),
                    ('Неверный поворот', 1), ('Неверный поворот', 2)]) as attempt:
            self.assertEqual(documents.ocr(image), 'Подтверждённая трудовая история')
            self.assertEqual(attempt.call_count, 4)
        path = self.root / 'unreadable.png'
        image.save(path)
        with patch('app.documents.ocr', side_effect=ValueError('Не читается вторая страница')):
            with self.assertRaisesRegex(ValueError, 'вторая страница'):
                documents.extract_file(path)
        self.assertFalse((self.root / 'state' / 'text_cache' / (common.sha(path) + '.json')).exists())


if __name__ == '__main__':
    unittest.main()
