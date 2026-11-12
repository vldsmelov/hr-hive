"""Generated-template projections use synthetic workbooks and no SQLite connections."""
import copy
import json
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest.mock import MagicMock, patch

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection, Side

from app import books


class GeneratedTemplateTest(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.run = self.root / 'run'
        self.run.mkdir()
        self.source = self.run / 'template.xlsx'
        self.output = self.root / 'result.xlsx'
        self.meta = {
            'source_book': 'template.xlsx', 'sheet': 'Из резюме',
            'output_path': str(self.output), 'id_column': 1,
            'source_column_count': 38, 'export_layout': 'append_results',
            'generated_from_resumes': True,
        }
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = self.meta['sheet']
        headers = ['Исходное поле ' + str(col) for col in range(1, 39)]
        for col, title in books.GENERATED_SOURCE_HEADERS.items():
            headers[col - 1] = title
        headers[13] = 'Файлы резюме'
        sheet.append(headers)
        for ident in ('101', '102', '103', '104', '105'):
            sheet.append([ident] + [None] * 37)
            sheet.cell(sheet.max_row, 14, ident + '.docx; ' + ident + '.pdf')
        for col in (2, 11, 13):
            cell = sheet.cell(2, col)
            cell.font = Font(name='Calibri', bold=True, color='224466')
            cell.fill = PatternFill('solid', fgColor='DDEEFF')
            cell.border = Border(bottom=Side(style='thin', color='335577'))
            cell.alignment = Alignment(wrap_text=True, vertical='top')
            cell.protection = Protection(locked=False)
            cell.number_format = '@'
        sheet['D2'] = '=1+2'
        sheet['O2'] = 'Ссылка на оригинал'
        sheet['O2'].hyperlink = 'https://example.test/original'
        sheet['O2'].hyperlink.tooltip = 'Исходная ссылка'
        workbook.save(self.source)
        workbook.close()
        self.rows = []
        self.connection = MagicMock()
        self.connection.execute.return_value.fetchall.side_effect = lambda: self.rows
        self.connection.__enter__.return_value = self.connection
        self.db_patch = patch.object(books, 'connect', return_value=self.connection)
        self.db_patch.start()
        self.sqlite_patch = patch('sqlite3.connect', side_effect=AssertionError('SQLite запрещён в этих тестах'))
        self.sqlite_patch.start()

    def tearDown(self):
        self.sqlite_patch.stop()
        self.db_patch.stop()
        self.temp.cleanup()

    def candidate(self, ident='101', source_row=2, card=None, review='reviewed', extraction='ready'):
        result = {
            'department': 'Подходящих кластеров не выявлено', 'comment': 'Конкретный комментарий',
            'ban': 'нет', 'ban_comment': '', 'clusters': [],
        }
        if card is not None:
            result['fact_card'] = card
        return {'id': ident, 'source_row': source_row, 'review_status': review,
                'extraction_status': extraction, 'error': 'Не читается вторая страница',
                'result': json.dumps(result, ensure_ascii=False)}

    def export(self):
        (self.run / 'portable.json').write_text(json.dumps(self.meta, ensure_ascii=False), encoding='utf-8')
        return books.export(self.run)

    def edit_source(self, updates):
        workbook = load_workbook(self.source)
        for address, value in updates.items():
            workbook.active[address] = value
        workbook.save(self.source)
        workbook.close()

    def test_reviewed_fact_card_projects_only_three_empty_fields_and_preserves_source(self):
        self.rows = [self.candidate(card={'name': 'Синтетический Кандидат Без Разбивки',
                                         'phones': ['+70000000001', '000002'],
                                         'education': 'Подтверждённая квалификация'})]
        original = self.source.read_bytes()
        result = self.export()
        self.assertEqual(self.source.read_bytes(), original)
        self.assertEqual(result['rows'], 1)
        source = load_workbook(self.source)
        exported = load_workbook(self.output)
        try:
            sheet = exported.active
            self.assertEqual((sheet.max_row, sheet.max_column), (2, 44))
            self.assertEqual(sheet['B2'].value, 'Синтетический Кандидат Без Разбивки')
            self.assertEqual(sheet['K2'].value, '+70000000001; 000002')
            self.assertEqual(sheet['M2'].value, 'Подтверждённая квалификация')
            self.assertEqual([sheet.cell(2, col).value for col in (6, 7, 8)], [None] * 3)
            self.assertEqual(sheet['N2'].value, '101.docx; 101.pdf')
            self.assertEqual(sheet['D2'].value, '=1+2')
            self.assertEqual(sheet['D2'].data_type, 'f')
            self.assertEqual(sheet['O2'].hyperlink.target, 'https://example.test/original')
            for col in range(1, 39):
                a, b = source.active.cell(2, col), sheet.cell(2, col)
                for attr in ('font', 'fill', 'border', 'alignment', 'protection'):
                    self.assertEqual(copy.copy(getattr(a, attr)), copy.copy(getattr(b, attr)))
                self.assertEqual(a.number_format, b.number_format)
                if col not in (2, 11, 13):
                    self.assertEqual(a.value, b.value)
                    self.assertEqual(a.data_type, b.data_type)
            self.assertEqual(sheet['AM2'].value, 'да')
            self.assertEqual(sheet['AP2'].value, 'нет')
            self.assertEqual(sheet['AR2'].value, 'Завершено')
        finally:
            source.close()
            exported.close()

    def test_unknown_fields_and_unfinished_rows_remain_blank_with_concrete_status(self):
        stale = {'name': 'Несохранённое имя', 'phones': ['+70000000009'], 'education': 'Не использовать'}
        self.rows = [self.candidate(card={}),
                     self.candidate('103', 4, stale, review='pending', extraction='error'),
                     self.candidate('105', 6, stale, review='pending')]
        self.export()
        workbook = load_workbook(self.output)
        try:
            sheet = workbook.active
            self.assertEqual([sheet.cell(row, 1).value for row in (2, 3, 4)], ['101', '103', '105'])
            for row in (2, 3, 4):
                self.assertEqual([sheet.cell(row, col).value for col in (2, 11, 13)], [None] * 3)
            self.assertEqual(sheet['AM2'].value, 'да')
            self.assertIn(sheet['AM3'].value, (None, ''))
            self.assertIn(sheet['AP3'].value, (None, ''))
            self.assertIn('Не читается вторая страница', sheet['AR3'].value)
            self.assertEqual(sheet['AR4'].value, 'Ожидает проверки')
            self.assertEqual(sheet['N3'].value, '103.docx; 103.pdf')
        finally:
            workbook.close()

    def test_nonempty_manual_values_formulas_zero_and_whitespace_are_preserved(self):
        self.edit_source({'B2': 'Ручная отметка', 'K2': 0, 'M2': '=IF(1=1,"","")',
                          'B3': ' ', 'K3': 'Ручной контакт', 'M3': 'Ручное образование'})
        card = {'name': 'Имя из карточки', 'phones': ['+70000000001'], 'education': 'Из карточки'}
        self.rows = [self.candidate(card=card), self.candidate('102', 3, card)]
        original = self.source.read_bytes()
        self.export()
        self.assertEqual(self.source.read_bytes(), original)
        workbook = load_workbook(self.output)
        try:
            for address, expected in {'B2': 'Ручная отметка', 'K2': 0, 'M2': '=IF(1=1,"","")',
                                      'B3': ' ', 'K3': 'Ручной контакт', 'M3': 'Ручное образование'}.items():
                self.assertEqual(workbook.active[address].value, expected)
            self.assertEqual(workbook.active['M2'].data_type, 'f')
            self.assertEqual(workbook.active['K2'].data_type, 'n')
        finally:
            workbook.close()

    def test_existing_append_layout_without_opt_in_does_not_fill_source_fields(self):
        self.meta.pop('generated_from_resumes')
        self.edit_source({'B1': 'Прежний заголовок', 'K1': 'Контакт CRM', 'M1': 'Прежнее образование'})
        self.rows = [self.candidate(card={'name': 'Не проецировать', 'phones': ['+70000000001'],
                                         'education': 'Не проецировать'})]
        self.export()
        workbook = load_workbook(self.output)
        try:
            self.assertEqual([workbook.active.cell(2, col).value for col in (2, 11, 13)], [None] * 3)
            self.assertEqual(workbook.active['B1'].value, 'Прежний заголовок')
            self.assertEqual(workbook.active['AM2'].value, 'да')
        finally:
            workbook.close()

    def test_generated_mode_rejects_other_layouts_columns_and_headers(self):
        self.rows = [self.candidate(card={'name': 'Имя'})]
        original = dict(self.meta)
        for updates in ({'export_layout': 'legacy'}, {'source_column_count': 37},
                        {'id_column': 2}, {'generated_from_resumes': 'true'}):
            with self.subTest(updates=updates):
                self.meta = dict(original, **updates)
                with self.assertRaisesRegex(ValueError, 'generated_from_resumes'):
                    self.export()
        self.meta = original
        for address in ('A1', 'B1', 'K1', 'M1'):
            with self.subTest(header=address):
                workbook = load_workbook(self.source)
                old = workbook.active[address].value
                workbook.close()
                self.edit_source({address: 'Чужой заголовок'})
                with self.assertRaisesRegex(ValueError, 'схема шаблона'):
                    self.export()
                self.edit_source({address: old})
        self.assertFalse(self.output.exists())

    def test_export_cannot_overwrite_template_or_use_it_as_backup(self):
        self.rows = [self.candidate(card={'name': 'Имя'})]
        original = self.source.read_bytes()
        self.meta['output_path'] = str(self.source)
        with self.assertRaisesRegex(ValueError, 'исходный шаблон'):
            self.export()
        self.assertEqual(self.source.read_bytes(), original)
        # The target's automatic .previous.xlsx backup must not alias the source either.
        target = self.source
        self.source = self.source.with_name('template.previous.xlsx')
        target.rename(self.source)
        self.meta['source_book'] = self.source.name
        self.meta['output_path'] = str(target)
        with self.assertRaisesRegex(ValueError, 'исходный шаблон'):
            self.export()
        self.assertEqual(self.source.read_bytes(), original)
        self.assertFalse(target.exists())

    def test_verify_rejects_changed_projection_results_styles_and_links(self):
        self.rows = [self.candidate(card={'name': 'Подтверждённое имя', 'phones': ['+70000000001'],
                                         'education': 'Подтверждённое образование'})]
        self.export()
        valid = self.output.read_bytes()
        for mutation in ('B2', 'K2', 'M2', 'D2', 'AM2', 'AR2', 'style', 'link'):
            with self.subTest(mutation=mutation):
                self.output.write_bytes(valid)
                workbook = load_workbook(self.output)
                sheet = workbook.active
                if mutation == 'style':
                    sheet['B2'].font = Font(bold=False)
                elif mutation == 'link':
                    sheet['O2'].hyperlink = 'https://example.test/changed'
                else:
                    sheet[mutation] = 'Подмена'
                workbook.save(self.output)
                workbook.close()
                with self.assertRaises(AssertionError):
                    books.verify(self.output, self.source, self.meta['sheet'], self.rows, self.meta)

    def test_projected_text_is_never_an_excel_formula(self):
        self.rows = [self.candidate(card={'name': '=1+2', 'phones': ['+70000000001'],
                                         'education': '=HYPERLINK("https://example.test")'})]
        self.export()
        workbook = load_workbook(self.output)
        try:
            for address in ('B2', 'K2', 'M2'):
                self.assertEqual(workbook.active[address].data_type, 's')
            self.assertEqual(workbook.active['B2'].value, '=1+2')
        finally:
            workbook.close()

    def test_missing_native_cells_keep_implicit_style_when_file_style_zero_is_custom(self):
        # Reproduce the native generator: style 0 is @/wrap/top, while some
        # source blanks have no <c> at all. Those blanks read as General with
        # an implicit zero StyleArray, not the custom file style 0.
        with zipfile.ZipFile(self.source) as archive:
            parts = {name: archive.read(name) for name in archive.namelist()}
        styles = books.E.fromstring(parts['xl/styles.xml'])
        zero = styles.find(books.q('cellXfs'))[0]
        zero.set('numFmtId', '49')
        zero.set('applyAlignment', '1')
        alignment = zero.find(books.q('alignment'))
        if alignment is None:
            alignment = books.E.SubElement(zero, books.q('alignment'))
        alignment.set('wrapText', '1')
        alignment.set('vertical', 'top')
        sheet = books.E.fromstring(parts['xl/worksheets/sheet1.xml'])
        rows = {int(row.get('r')): row for row in sheet.find(books.q('sheetData'))}
        for col in ('B', 'K', 'M'):
            for cell in list(rows[2]):
                if cell.get('r') == col + '2':
                    rows[2].remove(cell)
            # A second candidate has explicit blank cells with an existing
            # custom s-index. Their style must remain that exact index.
            cell = books.E.SubElement(rows[3], books.q('c'), r=col + '3', s='1', t='inlineStr')
            books.E.SubElement(books.E.SubElement(cell, books.q('is')), books.q('t'))
        parts['xl/styles.xml'] = books.E.tostring(styles)
        parts['xl/worksheets/sheet1.xml'] = books.E.tostring(sheet)
        with zipfile.ZipFile(self.source, 'w', zipfile.ZIP_DEFLATED) as archive:
            for name, raw in parts.items():
                archive.writestr(name, raw)
        card = {'name': 'Имя из карточки', 'phones': ['+70000000001'], 'education': 'Образование'}
        self.rows = [self.candidate(card=card), self.candidate('102', 3, card)]
        original = self.source.read_bytes()
        source = load_workbook(self.source)
        try:
            for col in (2, 11, 13):
                self.assertIsNone(source.active.cell(2, col)._style)
                self.assertEqual(source.active.cell(2, col).number_format, 'General')
                self.assertEqual(source.active.cell(3, col).number_format, '@')
        finally:
            source.close()
        self.export()  # Includes the full value/type/style/hyperlink/result verification.
        self.assertEqual(self.source.read_bytes(), original)
        exported = load_workbook(self.output)
        try:
            for col in (2, 11, 13):
                self.assertEqual(exported.active.cell(2, col).number_format, 'General')
                self.assertEqual(copy.copy(exported.active.cell(2, col).alignment), Alignment())
                self.assertEqual(exported.active.cell(3, col).number_format, '@')
            self.assertEqual(exported.active['B2'].value, card['name'])
            self.assertEqual(exported.active['K2'].value, card['phones'][0])
            self.assertEqual(exported.active['M2'].value, card['education'])
        finally:
            exported.close()
        with zipfile.ZipFile(self.output) as archive:
            sheet = books.E.fromstring(archive.read('xl/worksheets/sheet1.xml'))
            cells = {cell.get('r'): cell for cell in sheet.iter(books.q('c'))}
            added = {cells[col + '2'].get('s') for col in ('B', 'K', 'M')}
            self.assertEqual(len(added), 1)
            self.assertNotIn('0', added)
            for col in ('B', 'K', 'M'):
                self.assertEqual(cells[col + '3'].get('s'), '1')


if __name__ == '__main__':
    unittest.main()
