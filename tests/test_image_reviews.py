"""Visual non-text exceptions are exact-hash decisions, never OCR heuristics."""
import hashlib
import io
import json
import tempfile
import subprocess
import unittest
import zipfile
from pathlib import Path
from unittest.mock import patch

from PIL import Image
import fitz
from app import documents


class ImageReviewTest(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.patch = patch.object(documents, 'DATA', self.root)
        self.patch.start()
        image = Image.new('RGB', (800, 800), 'blue')
        stream = io.BytesIO()
        image.save(stream, format='PNG')
        self.content = stream.getvalue()
        self.digest = hashlib.sha256(self.content).hexdigest()

    def tearDown(self):
        self.patch.stop()
        self.temp.cleanup()

    def docx(self, images=None):
        path = self.root / 'synthetic.docx'
        with zipfile.ZipFile(path, 'w') as archive:
            archive.writestr('word/document.xml',
                '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                '<w:body><w:p><w:r><w:t>Полный текст резюме: инженер, проектирование оборудования, '
                'подготовка чертежей, контроль качества.</w:t></w:r></w:p></w:body></w:document>')
            for name, content in (images or {'word/media/picture.png': self.content}).items():
                archive.writestr(name, content)
        return path

    def approve(self, digest=None, **changes):
        proof = {'classification': 'portrait', 'reason': 'Визуально просмотрено: портрет без текста.',
                 'reviewer': 'synthetic-test-reviewer', 'checked_date': '2026-10-22'}
        proof.update(changes)
        path = self.root / 'state' / 'image_reviews.json'
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps({'version': 1, 'images': {digest or self.digest: proof}}),
                        encoding='utf-8')
        return proof

    def test_unreviewed_or_invalid_proof_never_skips_unreadable_image(self):
        path = self.docx()
        cases = [None, {'classification': 'scan'}, {'reason': ''}, {'reviewer': ''},
                 {'checked_date': ''}, {'checked_date': 'неизвестно'},
                 {'digest': '0' * 64}]
        for changes in cases:
            with self.subTest(changes=changes):
                ledger = self.root / 'state' / 'image_reviews.json'
                ledger.unlink(missing_ok=True)
                if changes is not None:
                    self.approve(**changes)
                with patch.object(documents, 'ocr', side_effect=ValueError('Неразборчивый скан')):
                    with self.assertRaisesRegex(ValueError, 'визуальной проверки'):
                        documents.extract_file(path, use_cache=False)
                self.assertFalse((self.root / 'state' / 'text_cache' /
                                  (hashlib.sha256(path.read_bytes()).hexdigest() + '.json')).exists())

    def test_only_exact_approved_nontext_image_is_skipped_with_proof(self):
        path = self.docx()
        proof = self.approve()
        with patch.object(documents, 'ocr', side_effect=AssertionError('Портрет не требует OCR')):
            result = documents.extract_file(path, use_cache=False)
        self.assertTrue(result['complete'])
        self.assertIn('проектирование оборудования', result['pages'][0]['text'])
        evidence = result['metadata']['visual_review'][0]
        self.assertEqual(evidence['image_sha256'], self.digest)
        self.assertEqual(evidence['member'], 'word/media/picture.png')
        self.assertEqual(evidence['proof'], proof)
        self.assertEqual(evidence['status'], 'reviewed_nontext')
        self.assertIs(evidence['ocr'], False)

    def test_approval_does_not_skip_other_embedded_scan_or_standalone_image(self):
        self.approve()
        stream = io.BytesIO()
        Image.new('RGB', (800, 800), 'red').save(stream, format='PNG')
        path = self.docx({'word/media/portrait.png': self.content,
                          'word/media/scan.png': stream.getvalue()})
        with patch.object(documents, 'ocr', side_effect=ValueError('Неразборчивый скан')):
            with self.assertRaisesRegex(ValueError, 'scan.png'):
                documents.extract_file(path, use_cache=False)
            standalone = self.root / 'standalone.png'
            standalone.write_bytes(self.content)
            with self.assertRaisesRegex(ValueError, 'Неразборчивый скан'):
                documents.extract_file(standalone, use_cache=False)

    def test_revoked_proof_invalidates_cached_visual_exception(self):
        path = self.docx()
        self.approve()
        with patch.object(documents, 'ocr', side_effect=AssertionError('Не нужен OCR')):
            documents.extract_file(path)
        (self.root / 'state' / 'image_reviews.json').unlink()
        with patch.object(documents, 'ocr', side_effect=ValueError('Неразборчивый скан')):
            with self.assertRaisesRegex(ValueError, 'визуальной проверки'):
                documents.extract_file(path)


    def pdf(self, first_text=True):
        path = self.root / 'synthetic.pdf'
        pdf = fitz.open()
        first = pdf.new_page()
        if first_text:
            first.insert_text((60, 60), 'Complete engineering experience: assembly, maintenance, production. ' * 3)
        pdf.new_page()
        pdf.save(path)
        pdf.close()
        return path

    def approve_page(self, path, physical_page=2, classification='blank', **changes):
        proof = {'classification': classification, 'reason': 'Страница визуально проверена: текста нет.',
                 'reviewer': 'synthetic-test-reviewer', 'checked_date': '2026-10-22'}
        proof.update(changes)
        ledger = self.root / 'state' / 'page_reviews.json'
        ledger.parent.mkdir(parents=True, exist_ok=True)
        source_hash = hashlib.sha256(path.read_bytes()).hexdigest()
        ledger.write_text(json.dumps({'version': 1, 'documents': {
            source_hash: {'pages': {str(physical_page): proof}}}}), encoding='utf-8')
        return proof

    def test_blank_or_visual_pdf_requires_exact_page_approval_and_preserves_number(self):
        path = self.pdf()
        with patch.object(documents, 'ocr', side_effect=ValueError('Пустая страница')):
            with self.assertRaisesRegex(ValueError, 'Пустая страница'):
                documents.extract_file(path, use_cache=False)
        for classification, method in [('blank', 'reviewed_blank'),
                                       ('portfolio_nontext', 'reviewed_visual'),
                                       ('portfolio_visual', 'reviewed_visual')]:
            with self.subTest(classification=classification):
                proof = self.approve_page(path, classification=classification)
                with patch.object(documents, 'ocr', side_effect=AssertionError('Проверенная страница')):
                    result = documents.extract_file(path, use_cache=False)
                self.assertEqual([page['page'] for page in result['pages']], [1, 2])
                self.assertEqual(result['pages'][0]['method'], 'pdf_text')
                self.assertEqual(result['pages'][1]['method'], method)
                self.assertIn('не дословная цитата', result['pages'][1]['text'])
                self.assertEqual(result['metadata']['page_reviews'][0]['proof'], proof)
                self.assertIs(result['metadata']['page_reviews'][0]['not_cv_quote'], True)

    def test_pdf_approval_cannot_skip_text_or_other_page_or_changed_source(self):
        path = self.pdf()
        self.approve_page(path, physical_page=1)
        with patch.object(documents, 'ocr', side_effect=ValueError('Вторая страница не проверена')):
            with self.assertRaisesRegex(ValueError, 'Вторая страница'):
                documents.extract_file(path, use_cache=False)
        self.approve_page(path)
        # A harmless byte change still invalidates the exact-source decision.
        path.write_bytes(path.read_bytes() + b'\n% changed source\n')
        with patch.object(documents, 'ocr', side_effect=ValueError('Новый исходник не проверен')):
            with self.assertRaisesRegex(ValueError, 'Новый исходник'):
                documents.extract_file(path, use_cache=False)

    def test_full_native_text_is_kept_even_with_incorrect_blank_approval(self):
        path = self.pdf()
        proof = self.approve_page(path)
        ledger = self.root / 'state' / 'page_reviews.json'
        payload = json.loads(ledger.read_text())
        source_hash = hashlib.sha256(path.read_bytes()).hexdigest()
        payload['documents'][source_hash]['pages']['1'] = proof
        ledger.write_text(json.dumps(payload), encoding='utf-8')
        with patch.object(documents, 'ocr', side_effect=AssertionError('OCR не требуется')):
            result = documents.extract_file(path, use_cache=False)
        self.assertEqual(result['pages'][0]['method'], 'pdf_text')
        self.assertIn('engineering experience', result['pages'][0]['text'])
        self.assertEqual([item['physical_page'] for item in result['metadata']['page_reviews']], [2])


    def test_ocr_tsv_quotes_are_literal_and_never_consume_following_rows(self):
        header = 'level\tpage_num\tblock_num\tpar_num\tline_num\tword_num\tleft\ttop\twidth\theight\tconf\ttext\n'
        rows = [
            '5\t1\t1\t1\t1\t1\t10\t10\t20\t10\t96\t"ООО\n',
            '5\t1\t1\t1\t1\t2\t40\t10\t50\t10\t95\tЗавод"\n',
            '5\t1\t1\t1\t2\t1\t10\t30\t70\t10\t97\tинженер-конструктор\n',
            '5\t1\t1\t1\t2\t2\t90\t30\t70\t10\t98\tоборудования\n',
        ]
        completed = subprocess.CompletedProcess([], 0, (header+''.join(rows)).encode(), b'')
        with patch.object(documents.subprocess, 'run', return_value=completed):
            text, score = documents._ocr_attempt(Image.new('RGB', (100, 100), 'white'))
        self.assertIn('"ООО Завод"', text)
        self.assertIn('инженер-конструктор оборудования', text)
        self.assertNotIn('\t', text)
        self.assertNotIn('5\t1', text)
        self.assertGreater(score, 90)

    def test_version_two_ocr_cache_is_rebuilt_but_plain_text_can_be_reused(self):
        path = self.root / 'cache.txt'
        path.write_text('Актуальный полный текст кандидата.', encoding='utf-8')
        digest = hashlib.sha256(path.read_bytes()).hexdigest()
        cache = self.root / 'state' / 'text_cache' / (digest+'.json')
        cache.parent.mkdir(parents=True)
        record = {'sha256': digest, 'extraction_version': 2, 'complete': True,
                  'pages': [{'page': 1, 'method': 'image_ocr', 'text': 'Прежнее распознавание'}]}
        cache.write_text(json.dumps(record), encoding='utf-8')
        result = documents.extract_file(path)
        self.assertEqual(result['extraction_version'], 3)
        self.assertEqual(result['pages'][0]['text'], path.read_text())
        record['pages'][0]['method'] = 'document_xml'
        record['pages'][0]['text'] = path.read_text()
        cache.write_text(json.dumps(record), encoding='utf-8')
        self.assertEqual(documents.extract_file(path), record)

    def test_verified_short_pdf_text_requires_transcription(self):
        path = self.pdf()
        self.approve_page(path, classification='verified_text')
        with patch.object(documents, 'ocr', side_effect=ValueError('Нет транскрипции')):
            with self.assertRaisesRegex(ValueError, 'Страница PDF 2'):
                documents.extract_file(path, use_cache=False)
        self.approve_page(path, classification='verified_text', transcription='КОНЕЦ ДОКУМЕНТА')
        with patch.object(documents, 'ocr', side_effect=AssertionError('Страница прочитана вручную')):
            result = documents.extract_file(path, use_cache=False)
        self.assertEqual(result['pages'][1]['method'], 'reviewed_text')
        self.assertIn('Ручная транскрипция', result['pages'][1]['text'])
        self.assertIn('КОНЕЦ ДОКУМЕНТА', result['pages'][1]['text'])

    def test_verified_transcription_overrides_mixed_pdf_only_for_exact_source(self):
        path = self.pdf()
        pdf = fitz.open(path)
        pdf[0].insert_image(fitz.Rect(20, 100, 580, 800), stream=self.content)
        raw = pdf.tobytes()
        pdf.close()
        path.write_bytes(raw)
        blank_proof = self.approve_page(path)
        with patch.object(documents, 'ocr', return_value='Прежнее распознавание без всех названий программ.'):
            previous = documents.extract_file(path, use_cache=False)
        self.assertEqual(previous['pages'][0]['method'], 'ocr')
        transcription = 'Полная проверенная страница: образование, все программы и проекты.'
        proof = self.approve_page(path, physical_page=1, classification='verified_text',
                                  transcription=transcription)
        ledger = self.root / 'state' / 'page_reviews.json'
        payload = json.loads(ledger.read_text())
        digest = hashlib.sha256(raw).hexdigest()
        payload['documents'][digest]['pages']['2'] = blank_proof
        ledger.write_text(json.dumps(payload), encoding='utf-8')
        self.assertFalse(documents._cached_visual_reviews_valid(previous))
        with patch.object(documents, 'ocr', side_effect=AssertionError('Транскрипция уже проверена')):
            result = documents.extract_file(path)
        self.assertEqual(result['pages'][0]['method'], 'reviewed_text')
        self.assertIn(transcription, result['pages'][0]['text'])
        self.assertEqual(result['metadata']['page_reviews'][0]['proof'], proof)
        self.assertFalse(result['metadata']['page_reviews'][0]['not_cv_quote'])
        self.assertEqual(result['pages'][1]['method'], 'reviewed_blank')
        path.write_bytes(raw + b'\n% different source\n')
        with patch.object(documents, 'ocr', side_effect=ValueError('Нет проверки нового исходника')):
            with self.assertRaisesRegex(ValueError, 'Нет проверки нового исходника'):
                documents.extract_file(path, use_cache=False)

    def test_small_page_keeps_both_page_segmentation_passes(self):
        with patch.object(documents.shutil, 'which', return_value='/synthetic/tesseract'), \
                patch.object(documents, '_ocr_attempt', side_effect=[
                    ('Полная трудовая история кандидата', 95),
                    ('Полная трудовая история кандидата. Специалист по закупкам 2024', 96)]) as attempt:
            text = documents.ocr(Image.new('RGB', (500, 700), 'white'))
        self.assertIn('Дополнительное распознавание всей страницы', text)
        self.assertIn('Специалист по закупкам', text)
        self.assertEqual(attempt.call_args_list[0].args[0].size, (2000, 2800))
        self.assertEqual(attempt.call_args_list[1].kwargs, {'psm': 11})

    def test_large_photo_timeout_uses_bounded_moire_fallback(self):
        with patch.object(documents.shutil, 'which', return_value='/synthetic/tesseract'), \
                patch.object(documents, '_ocr_attempt', side_effect=[
                    subprocess.TimeoutExpired('tesseract', 45),
                    ('Полная трудовая история после подавления муара', 96)]) as attempt:
            text = documents.ocr(Image.new('RGB', (3600, 4000), 'white'))
        self.assertIn('подавления муара', text)
        self.assertEqual(attempt.call_count, 2)
        self.assertEqual(attempt.call_args.args[0].width, 1200)

    def test_low_confidence_is_not_accepted_after_rotations(self):
        with patch.object(documents.shutil, 'which', return_value='/synthetic/tesseract'), \
                patch.object(documents, '_ocr_attempt', return_value=('Сомнительное распознавание', 45.5)):
            with self.assertRaisesRegex(ValueError, 'надёжный текст'):
                documents.ocr(Image.new('RGB', (300, 100), 'white'))

    def test_approved_vector_is_checked_before_pillow_loader(self):
        path = self.docx({'word/media/picture.wmf': self.content})
        self.approve(classification='decoration')
        with patch.object(documents.Image, 'open', side_effect=AssertionError('WMF loader не нужен')):
            result = documents.extract_file(path, use_cache=False)
        self.assertEqual(result['metadata']['visual_review'][0]['member'], 'word/media/picture.wmf')

    def test_standalone_transcription_requires_exact_page_and_current_proof(self):
        path = self.root / 'scan.png'
        path.write_bytes(self.content)
        for changes in [dict(physical_page=2, classification='verified_text', transcription='Полный текст'),
                        dict(physical_page=1, classification='blank')]:
            self.approve_page(path, **changes)
            with patch.object(documents, 'ocr', side_effect=ValueError('Неполный OCR')):
                with self.assertRaisesRegex(ValueError, 'Неполный OCR'):
                    documents.extract_file(path, use_cache=False)
        proof = self.approve_page(path, physical_page=1, classification='verified_text',
                                  transcription='Полная визуальная транскрипция: имя, телефон, все работы.')
        with patch.object(documents, 'ocr', side_effect=AssertionError('Проверено визуально')):
            result = documents.extract_file(path, use_cache=False)
        self.assertIn('все работы', result['pages'][0]['text'])
        self.assertEqual(result['pages'][0]['method'], 'reviewed_text')
        self.assertEqual(result['metadata']['page_reviews'][0]['proof'], proof)
        self.assertFalse(result['metadata']['page_reviews'][0]['not_cv_quote'])
        (self.root / 'state' / 'page_reviews.json').unlink()
        with patch.object(documents, 'ocr', side_effect=ValueError('Доказательство отозвано')):
            with self.assertRaisesRegex(ValueError, 'отозвано'):
                documents.extract_file(path)

    def test_unknown_vector_converts_whole_document_instead_of_skipping(self):
        path = self.docx({'word/media/picture.wmf': b'synthetic vector bytes'})
        def convert(command, **kwargs):
            out = Path(command[command.index('--outdir')+1])
            pdf = fitz.open()
            page = pdf.new_page()
            page.insert_text((60, 60), 'Complete visible document text from vector conversion.')
            pdf.save(out/'source.pdf')
            pdf.close()
            return subprocess.CompletedProcess(command, 0, b'', b'')
        with patch.object(documents.shutil, 'which', return_value='/synthetic/soffice'), \
                patch.object(documents.subprocess, 'run', side_effect=convert):
            result = documents.extract_file(path, use_cache=False)
        self.assertEqual(result['pages'][0]['method'], 'pdf_text')
        self.assertIn('Complete visible document text', result['pages'][0]['text'])
        self.assertEqual(result['metadata']['document_conversion']['members'], ['word/media/picture.wmf'])


if __name__ == '__main__':
    unittest.main()
