"""Synthetic download aliases exercise real extraction without production data."""
import hashlib
import os
import tempfile
import unittest
import zipfile
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from unittest.mock import patch

from app import documents, parallel_extract


class ExtractAliasesTest(unittest.TestCase):
    def setUp(self):
        temporary = tempfile.TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        self.root = Path(temporary.name)
        # Keep the real file extractor, but run workers in-process so patched
        # cache/proof paths stay temporary on both spawn and fork platforms.
        patches = [patch.object(documents, 'DATA', self.root),
                   patch.object(parallel_extract, 'ProcessPoolExecutor', ThreadPoolExecutor),
                   patch.dict(os.environ, {'OMP_THREAD_LIMIT': '1'})]
        for item in patches:
            item.start()
            self.addCleanup(item.stop)

    def docx(self):
        path = self.root / '11989.docx'
        with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as archive:
            archive.writestr('[Content_Types].xml', '''<?xml version="1.0"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
 <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
 <Default Extension="xml" ContentType="application/xml"/>
 <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
 <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
 <Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
</Types>''')
            archive.writestr('_rels/.rels', '''<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
 <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>''')
            archive.writestr('word/document.xml', '''<?xml version="1.0"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
 <w:body>
  <w:p><w:r><w:t>Первая страница: синтетический опыт сборки оборудования.</w:t></w:r></w:p>
  <w:p><w:r><w:br w:type="page"/><w:t>Вторая страница: полная история и техническое образование.</w:t></w:r></w:p>
  <w:sectPr>
   <w:headerReference w:type="default" r:id="rIdHeader"/>
   <w:footerReference w:type="default" r:id="rIdFooter"/>
  </w:sectPr>
 </w:body>
</w:document>''')
            archive.writestr('word/header1.xml', '''<?xml version="1.0"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
 <w:p><w:r><w:t>Синтетический заголовок: дополнительная квалификация.</w:t></w:r></w:p>
</w:hdr>''')
            archive.writestr('word/footer1.xml', '''<?xml version="1.0"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
 <w:p><w:r><w:t>Синтетический нижний колонтитул: конец полного документа.</w:t></w:r></w:p>
</w:ftr>''')
            archive.writestr('word/_rels/document.xml.rels', '''<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
 <Relationship Id="rIdHeader" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
 <Relationship Id="rIdFooter" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
</Relationships>''')
        return path

    def test_download_alias_first_uses_full_docx_and_keeps_all_text_and_identity(self):
        full = self.docx()
        download = full.with_suffix('.crdownload')
        content = full.read_bytes()
        download.write_bytes(content)
        digest = hashlib.sha256(content).hexdigest()
        self.assertFalse((self.root / 'state' / 'text_cache').exists())

        results = list(parallel_extract.iter_extracted(
            [download, full], max_workers=2, use_cache=False))

        self.assertEqual([row[0] for row in results], [str(download), str(full)])
        for path, document, error in results:
            self.assertIsNone(error)
            self.assertIsNotNone(document)
            self.assertTrue(document['complete'])
            self.assertEqual(document['sha256'], digest)
            self.assertEqual(document['original_name'], Path(path).name)
            # Native Word extraction groups the two body pages in one part,
            # with header and footer emitted as additional extracted parts.
            self.assertEqual([page['page'] for page in document['pages']], [1, 2, 3])
            self.assertTrue(all(page['method'] == 'document_xml'
                                for page in document['pages']))
            texts = [page['text'] for page in document['pages']]
            self.assertIn('Первая страница: синтетический опыт сборки оборудования.', texts[0])
            self.assertIn('Вторая страница: полная история и техническое образование.', texts[0])
            self.assertEqual(texts[1], 'Синтетический заголовок: дополнительная квалификация.')
            self.assertEqual(texts[2], 'Синтетический нижний колонтитул: конец полного документа.')
        self.assertEqual(results[0][1]['pages'], results[1][1]['pages'])
        self.assertEqual(full.read_bytes(), content)
        self.assertEqual(download.read_bytes(), content)

    def test_standalone_download_still_fails_without_complete_sibling_or_cache(self):
        full = self.docx()
        content = full.read_bytes()
        full.unlink()
        for suffix in ('.crdownload', '.CRDOWNLOAD'):
            with self.subTest(suffix=suffix):
                download = self.root / ('11989' + suffix)
                download.write_bytes(content)
                results = list(parallel_extract.iter_extracted(
                    [download], max_workers=1, use_cache=False))
                self.assertEqual(len(results), 1)
                path, document, error = results[0]
                self.assertEqual(path, str(download))
                self.assertIsNone(document)
                self.assertIn('Незавершённая загрузка .crdownload', error)
                with self.assertRaisesRegex(ValueError, 'Незавершённая загрузка'):
                    documents.extract_file(download, use_cache=False)
                self.assertFalse((self.root / 'state' / 'text_cache').exists())


if __name__ == '__main__':
    unittest.main()
