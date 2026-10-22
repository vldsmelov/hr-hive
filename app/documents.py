"""Local, full-document extraction with content-addressed cache and OCR."""
import io
import json
import re
import shutil
import subprocess
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

import fitz
from PIL import Image, ImageOps, ImageSequence
from bs4 import BeautifulSoup
from striprtf.striprtf import rtf_to_text
from .common import DATA, dump, read, sha

def decode(raw):
    for enc in ('utf-8-sig', 'utf-16' if raw.startswith((b'\xff\xfe', b'\xfe\xff')) else 'utf-8', 'cp1251'):
        try:
            return raw.decode(enc)
        except UnicodeError:
            pass
    raise ValueError('Не удалось определить кодировку текста')

def broken(text):
    return bool(re.search(r'\?{3,}|\ufffd{2,}', text)) or (len(text) > 40 and sum('\u3400' <= c <= '\u9fff' for c in text) / len(text) > .1)

def ocr(image):
    if not shutil.which('tesseract'):
        raise RuntimeError('OCR недоступен: требуется tesseract rus+eng')
    with tempfile.TemporaryDirectory() as t:
        p = Path(t) / 'page.png'
        ImageOps.exif_transpose(image).convert('RGB').save(p)
        proc = subprocess.run(['tesseract', str(p), 'stdout', '-l', 'rus+eng', '--psm', '3'], capture_output=True, timeout=180)
        if proc.returncode:
            raise ValueError(proc.stderr.decode('utf-8', errors='replace'))
        text = proc.stdout.decode('utf-8').strip()
        if len(re.sub(r'\W', '', text)) < 12 or broken(text):
            raise ValueError('OCR не восстановил читаемый текст страницы')
        return text

def extract_file(path, use_cache=True):
    path = Path(path)
    h = sha(path)
    cache = DATA / 'state' / 'text_cache' / (h + '.json')
    if use_cache and cache.exists():
        return read(cache)
    raw = path.read_bytes()
    if path.suffix.lower() == '.crdownload':
        raise ValueError('Незавершённая загрузка .crdownload')
    pages = []
    def page(text, method):
        if broken(text):
            raise ValueError('Повреждённая кодировка/текстовый слой')
        pages.append({'page': len(pages)+1, 'method': method, 'text': text})
    def pdf(p):
        with fitz.open(p) as doc:
            for pg in doc:
                text = pg.get_text(sort=True).strip()
                images = sum(r.width*r.height for x in pg.get_images() for r in pg.get_image_rects(x[0]))
                if len(re.sub(r'\W', '', text)) < 70 or broken(text) or images > pg.rect.width*pg.rect.height*.5:
                    pix = pg.get_pixmap(matrix=fitz.Matrix(2.5, 2.5), alpha=False)
                    page(ocr(Image.frombytes('RGB', (pix.width, pix.height), pix.samples)), 'ocr')
                else:
                    page(text, 'pdf_text')
    def convert():
        executable = shutil.which('soffice')
        if not executable:
            raise RuntimeError('Требуется LibreOffice для конвертации документа')
        with tempfile.TemporaryDirectory() as t:
            proc = subprocess.run([executable, '-env:UserInstallation='+ (Path(t)/'profile').as_uri(), '--headless', '--convert-to', 'pdf', '--outdir', t, str(path.resolve())], capture_output=True, timeout=120)
            files = list(Path(t).glob('*.pdf'))
            if proc.returncode or not files:
                raise ValueError('LibreOffice не создал PDF: '+proc.stderr.decode('utf-8', errors='replace')[:400])
            pdf(files[0])
    if raw.startswith(b'%PDF'):
        pdf(path)
    elif raw.startswith(b'\xd0\xcf\x11\xe0'):
        convert()
    elif raw.startswith(b'{\\rtf'):
        if not raw.rstrip(b'\r\n\t \x00').endswith(b'}'):
            raise ValueError('Неполный RTF')
        if b'\\pict' in raw:
            convert()
        else:
            page(rtf_to_text(raw.decode('latin1'), errors='strict'), 'rtf')
    elif raw.startswith(b'PK'):
        with zipfile.ZipFile(path) as z:
            names = z.namelist()
            if 'Index/Document.iwa' in names:
                convert()
            else:
                parts = [n for n in names if n == 'word/document.xml' or re.match(r'word/(header|footer|footnotes|endnotes).*\.xml$', n)]
                if not parts:
                    parts = sorted((n for n in names if re.match(r'ppt/slides/slide\d+\.xml$', n)), key=lambda n: int(re.search(r'(\d+)\.xml', n).group(1)))
                if not parts and 'content.xml' in names:
                    parts = ['content.xml']
                if not parts:
                    raise ValueError('Неподдерживаемый ZIP-документ')
                for part in parts:
                    x = ET.fromstring(z.read(part))
                    paragraphs = [n for n in x.iter() if n.tag.split('}')[-1] in ('p', 'h')]
                    text = '\n'.join(''.join(n.itertext()) for n in paragraphs)
                    if text.strip():
                        page(text, 'document_xml')
                for name in names:
                    if '/media/' not in name and not name.startswith('Pictures/'):
                        continue
                    # Do not silently skip a damaged embedded scan.
                    content = z.read(name)
                    try:
                        im = Image.open(io.BytesIO(content))
                    except OSError:
                        if name.lower().endswith(('.png','.jpg','.jpeg','.tif','.tiff','.bmp')):
                            raise ValueError('Повреждено встроенное изображение '+name)
                        continue
                    with im:
                        im.load()
                        if im.width >= 600 and im.height >= 500:
                            page(ocr(im), 'embedded_ocr')
    elif raw.startswith((b'\xff\xd8', b'\x89PNG', b'GIF8', b'BM', b'II*', b'MM\x00')) or path.suffix.lower() in ('.webp', '.jfif'):
        with Image.open(path) as im:
            for frame in ImageSequence.Iterator(im):
                page(ocr(frame.copy()), 'image_ocr')
    elif path.suffix.lower() in ('.txt', '.md'):
        page(decode(raw), 'text')
    elif b'<html' in raw[:1000].lower() or b'<!doctype' in raw[:1000].lower():
        page(BeautifulSoup(raw, 'html.parser').get_text('\n', strip=True), 'html')
    else:
        raise ValueError('Неподдерживаемый или повреждённый формат: '+path.suffix)
    if not pages or not any(p['text'].strip() for p in pages):
        raise ValueError('Пустой документ')
    result = {'sha256': h, 'pages': pages, 'original_name': path.name}
    dump(cache, result)
    return result
