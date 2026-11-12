"""Local, full-document extraction with content-addressed cache and OCR."""
import io
import hashlib
from datetime import date
import json
import os
import csv
import statistics
import re
import shutil
import subprocess
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

import fitz
from PIL import Image, ImageOps, ImageSequence, ImageFilter
from bs4 import BeautifulSoup
from striprtf.striprtf import rtf_to_text
from .common import DATA, read, sha

EXTRACTION_VERSION = 3

def source_block_reason(digest):
    """A human-confirmed incomplete original stays blocked across OCR retries."""
    path = DATA / 'state' / 'source_blocks' / (digest + '.json')
    if not path.exists():
        return None
    proof = read(path)
    if proof.get('sha256') != digest or not str(proof.get('reason', '')).strip():
        raise ValueError('Повреждено подтверждение неполноты исходника')
    return proof['reason']


def require_complete_source(digest):
    reason = source_block_reason(digest)
    if reason:
        raise ValueError('Исходник неполон: ' + reason)


def decode(raw):
    for enc in ('utf-8-sig', 'utf-16' if raw.startswith((b'\xff\xfe', b'\xfe\xff')) else 'utf-8', 'cp1251'):
        try:
            return raw.decode(enc)
        except UnicodeError:
            pass
    raise ValueError('Не удалось определить кодировку текста')

def broken(text):
    return bool(re.search(r'\?{3,}|\ufffd{2,}', text)) or (len(text) > 40 and sum('\u3400' <= c <= '\u9fff' for c in text) / len(text) > .1)

def _ocr_attempt(image, psm=3):
    """One local OCR pass; TSV confidence prevents accepting rotated gibberish."""
    with tempfile.TemporaryDirectory() as t:
        p = Path(t) / 'page.png'
        image.convert('RGB').save(p)
        proc = subprocess.run(
            ['tesseract', str(p), 'stdout', '-l', 'rus+eng', '--psm', str(psm), 'tsv'],
            capture_output=True, timeout=45,
            env={**os.environ, 'OMP_THREAD_LIMIT': '1'})
        if proc.returncode:
            raise ValueError(proc.stderr.decode('utf-8', errors='replace'))
        rows = csv.DictReader(io.StringIO(proc.stdout.decode('utf-8')), delimiter='\t', quoting=csv.QUOTE_NONE)
        lines, words, confidences = {}, [], []
        for row in rows:
            word = (row.get('text') or '').strip()
            if not word:
                continue
            key = tuple(row.get(k, '') for k in ('page_num', 'block_num', 'par_num', 'line_num'))
            lines.setdefault(key, []).append(word)
            if sum(c.isalpha() for c in word) >= 2:
                words.append(word)
                confidences.append(float(row['conf']))
        text = '\n'.join(' '.join(v) for v in lines.values())
        score = statistics.median(confidences) if confidences else 0
        if len(words) < 3 or len(re.sub(r'\W', '', text)) < 12 or broken(text):
            score = 0
        return text, score


def ocr(image):
    if not shutil.which('tesseract'):
        raise RuntimeError('OCR недоступен: требуется tesseract rus+eng')
    image = ImageOps.exif_transpose(image)
    if image.mode in ('RGBA', 'LA') or 'transparency' in image.info:
        rgba = image.convert('RGBA')
        image = Image.alpha_composite(Image.new('RGBA', rgba.size, 'white'), rgba)
    image = image.convert('RGB')
    small_page = 350 <= min(image.size) < 1200 and max(image.size) < 3500
    scale = 4 if max(image.size) < 1200 else 3
    initial = image.resize((image.width*scale, image.height*scale), Image.Resampling.LANCZOS) if small_page else image

    def attempt(im, psm=3):
        try:
            return _ocr_attempt(im) if psm == 3 else _ocr_attempt(im, psm=psm)
        except subprocess.TimeoutExpired:
            return '', 0

    text, score = attempt(initial)
    attempts = [(text, score)]
    selected = initial
    if score < 65 and max(image.size) > 3500:
        # Suppress photographic moire before reducing a large camera image.
        filtered = ImageOps.grayscale(image).filter(ImageFilter.GaussianBlur(1.5))
        filtered = filtered.resize((1200, max(1, round(filtered.height*1200/filtered.width))),
                                   Image.Resampling.LANCZOS)
        filtered = ImageOps.autocontrast(filtered)
        result = attempt(filtered)
        attempts.append(result)
        if result[1] > score:
            text, score = result
            selected = filtered
    if score < 65:
        for angle in (90, 180, 270):
            attempts.append(attempt(selected.rotate(angle, expand=True)))
        text, score = max(attempts, key=lambda result: result[1])
    if score < 65:
        raise ValueError('OCR не восстановил надёжный текст страницы после проверки поворота и масштаба')
    if small_page:
        # Sparse-layout pass often recovers small date/job labels omitted by
        # automatic page segmentation. Preserve it explicitly for full review.
        extra, extra_score = attempt(initial, psm=11)
        if extra_score >= 65 and extra.strip() != text.strip():
            text += '\n\n[Дополнительное распознавание всей страницы: разреженный текст]\n'+extra
    return text


def _recalculate_xlsx(path):
    """Recalculate a disposable local copy; never update the applicant source."""
    executable = shutil.which('soffice')
    if not executable:
        raise RuntimeError('Требуется LibreOffice Calc для проверки формул XLSX')
    with tempfile.TemporaryDirectory() as t:
        t = Path(t)
        source = t / 'input' / 'questionnaire.xlsx'
        source.parent.mkdir()
        source.write_bytes(Path(path).read_bytes())
        out = t / 'out'
        out.mkdir()
        proc = subprocess.run(
            [executable, '-env:UserInstallation='+(t/'profile').as_uri(),
             '--headless', '--convert-to', 'xlsx', '--outdir', str(out), str(source)],
            capture_output=True, timeout=180,
            env={**os.environ, 'OMP_THREAD_LIMIT': '1'})
        result = out / source.name
        if proc.returncode or not result.exists():
            raise ValueError('LibreOffice Calc не пересчитал копию XLSX: '+
                             proc.stderr.decode('utf-8', errors='replace')[:400])
        return result.read_bytes()


def _nontext_review(image_sha256):
    """Only explicit, content-addressed visual decisions may bypass image OCR."""
    ledger = DATA / 'state' / 'image_reviews.json'
    if not ledger.exists():
        return None
    payload = read(ledger)
    proof = payload.get('images', {}).get(image_sha256)
    if not isinstance(proof, dict) or proof.get('classification') not in ('portrait', 'logo', 'decoration'):
        return None
    if any(not isinstance(proof.get(k), str) or not proof[k].strip()
           for k in ('reason', 'reviewer', 'checked_date')):
        return None
    try:
        date.fromisoformat(proof['checked_date'])
    except ValueError:
        return None
    return {k: proof[k] for k in ('classification', 'reason', 'reviewer', 'checked_date')}


def _page_review(source_sha256, physical_page):
    ledger = DATA / 'state' / 'page_reviews.json'
    if not ledger.exists():
        return None
    payload = read(ledger)
    proof = payload.get('documents', {}).get(source_sha256, {}).get('pages', {}).get(str(physical_page))
    if not isinstance(proof, dict) or proof.get('classification') not in ('blank', 'portfolio_nontext', 'portfolio_visual', 'verified_text'):
        return None
    if any(not isinstance(proof.get(k), str) or not proof[k].strip()
           for k in ('reason', 'reviewer', 'checked_date')):
        return None
    try:
        date.fromisoformat(proof['checked_date'])
    except ValueError:
        return None
    fields = ('classification', 'reason', 'reviewer', 'checked_date')
    if proof['classification'] == 'verified_text':
        if not isinstance(proof.get('transcription'), str) or not proof['transcription'].strip():
            return None
        fields += ('transcription',)
    if proof['classification'] == 'portfolio_visual':
        if 'visual_description' in proof:
            if not isinstance(proof['visual_description'], str) or not proof['visual_description'].strip():
                return None
            fields += ('visual_description',)
        if 'transcription' in proof:
            if not isinstance(proof['transcription'], str):
                return None
            fields += ('transcription',)
    return {k: proof[k] for k in fields}


def _cached_visual_reviews_valid(cached):
    recorded_pages = {}
    for item in cached.get('metadata', {}).get('page_reviews', []):
        proof = _page_review(item.get('source_sha256', ''), item.get('physical_page'))
        if not proof or proof != item.get('proof'):
            return False
        recorded_pages[item.get('physical_page')] = proof
    # A newly verified transcription must also invalidate an older OCR/native
    # result that predates the proof. Preview uses this check before saving.
    for page in cached.get('pages', []):
        proof = _page_review(cached.get('sha256', ''), page.get('page'))
        if (proof and proof['classification'] == 'verified_text' and
                recorded_pages.get(page.get('page')) != proof):
            return False
    for item in cached.get('metadata', {}).get('visual_review', []):
        if item.get('ocr') is False:
            proof = _nontext_review(item.get('image_sha256', ''))
            if not proof or proof != item.get('proof'):
                return False
    return True


def _atomic_cache(path, obj):
    path.parent.mkdir(parents=True, exist_ok=True)
    # Workers may encounter identical files: unique temporary names and atomic
    # replace avoid shared .tmp races without touching SQLite.
    with tempfile.NamedTemporaryFile('w', dir=path.parent, encoding='utf-8',
                                     prefix=path.name+'.', suffix='.tmp', delete=False) as f:
        temporary = Path(f.name)
        json.dump(obj, f, ensure_ascii=False, indent=2, default=str)
    try:
        os.replace(temporary, path)
    finally:
        temporary.unlink(missing_ok=True)


def _auxiliary_mpo(raw, image, index):
    if index == 0 or image.format != 'MPO':
        return False
    # A secondary MPO frame is ignored ONLY with explicit HDR/gain-map
    # metadata. Ordinary multi-picture files still require every frame.
    markers = (b'hdrgm:', b'urn:com:apple:photo:2020:aux:hdrgainmap',
               b'urn:com:apple:photo:2020:aux:hdr-gain-map')
    # A two-frame photo with a smaller gain map is the supported HDR case.
    # Multiple full-size pictures remain pages even when HDR tags are present.
    if getattr(image, 'n_frames', 1) != 2 or not any(marker in raw for marker in markers):
        return False
    current = image.tell()
    secondary_size = image.size
    image.seek(0)
    primary_size = image.size
    image.seek(current)
    return secondary_size[0]*secondary_size[1] < primary_size[0]*primary_size[1]


def extract_file(path, use_cache=True):
    path = Path(path)
    h = sha(path)
    require_complete_source(h)
    cache = DATA / 'state' / 'text_cache' / (h + '.json')
    if use_cache and cache.exists():
        cached = read(cache)
        version_compatible = (cached.get('extraction_version') == EXTRACTION_VERSION or
                              (cached.get('extraction_version') == 2 and
                               not any('ocr' in pg.get('method', '').lower()
                                       for pg in cached.get('pages', []))))
        if (version_compatible and
                cached.get('complete') is True and cached.get('sha256') == h and
                cached.get('pages') and _cached_visual_reviews_valid(cached) and
                all(pg.get('page') == n and pg.get('text', '').strip() and
                    not broken(pg['text']) and pg.get('method') != 'ocr_pending'
                    for n, pg in enumerate(cached['pages'], 1))):
            return cached
    raw = path.read_bytes()
    if path.suffix.lower() == '.crdownload':
        raise ValueError('Незавершённая загрузка .crdownload')
    pages = []
    metadata = {}
    def reviewed_image(content, member):
        digest = hashlib.sha256(content).hexdigest()
        proof = _nontext_review(digest)
        if proof:
            metadata.setdefault('visual_review', []).append({
                'image_sha256': digest, 'member': member, 'ocr': False,
                'proof': proof, 'status': 'reviewed_nontext'})
            return True
        return False
    def page(text, method):
        if not text.strip() or broken(text):
            raise ValueError('Пустая страница или повреждённая кодировка/текстовый слой')
        pages.append({'page': len(pages)+1, 'method': method, 'text': text})
    def pdf(p):
        with fitz.open(p) as doc:
            for pg in doc:
                text = pg.get_text(sort=True).strip()
                physical_page = pg.number + 1
                proof = _page_review(h, physical_page)
                # A reviewed blank/photo page never replaces readable native
                # text. Exact source hash and physical page prevent drift.
                if proof and (proof['classification'] == 'verified_text' or
                              len(re.sub(r'\W', '', text)) < 12):
                    if proof['classification'] == 'verified_text':
                        method = 'reviewed_text'
                        page('[Ручная транскрипция после визуальной проверки страницы '
                             f'{physical_page}]\n'+proof['transcription'], method)
                    else:
                        method = 'reviewed_blank' if proof['classification'] == 'blank' else 'reviewed_visual'
                        description = proof.get('visual_description', proof['reason'])
                        body = (f'[Визуальная проверка страницы {physical_page}: {description} '
                                'Описание проверяющего, не дословная цитата резюме.]')
                        if proof.get('transcription'):
                            body += ('\n[Основные надписи портфолио: ручная транскрипция]\n'+
                                     proof['transcription'])
                        page(body, method)
                    metadata.setdefault('page_reviews', []).append({
                        'source_sha256': h, 'physical_page': physical_page,
                        'proof': proof,
                        'not_cv_quote': proof['classification'] != 'verified_text'})
                    continue
                images = sum(r.width*r.height for x in pg.get_images() for r in pg.get_image_rects(x[0]))
                if len(re.sub(r'\W', '', text)) < 12 or broken(text) or images > pg.rect.width*pg.rect.height*.5:
                    pix = pg.get_pixmap(matrix=fitz.Matrix(2.5, 2.5), alpha=False)
                    try:
                        page(ocr(Image.frombytes('RGB', (pix.width, pix.height), pix.samples)), 'ocr')
                    except (ValueError, subprocess.TimeoutExpired) as error:
                        raise ValueError(f'Страница PDF {physical_page}: {error}') from error
                else:
                    page(text, 'pdf_text')
    def convert(suffix=None):
        executable = shutil.which('soffice')
        if not executable:
            raise RuntimeError('Требуется LibreOffice для конвертации документа')
        with tempfile.TemporaryDirectory() as t:
            converted_source = Path(t)/('source'+(suffix or path.suffix or '.doc'))
            converted_source.write_bytes(raw)
            proc = subprocess.run([executable, '-env:UserInstallation='+ (Path(t)/'profile').as_uri(), '--headless', '--convert-to', 'pdf', '--outdir', t, str(converted_source)], capture_output=True, timeout=120, env={**os.environ, 'OMP_THREAD_LIMIT': '1'})
            files = list(Path(t).glob('*.pdf'))
            if proc.returncode or not files:
                raise ValueError('LibreOffice не создал PDF: '+proc.stderr.decode('utf-8', errors='replace')[:400])
            pdf(files[0])
    def spreadsheet():
        from openpyxl import load_workbook
        formulas = load_workbook(io.BytesIO(raw), data_only=False, keep_links=False)
        has_formulas = any(c.data_type == 'f' for ws in formulas if ws.sheet_state == 'visible'
                           for row in ws.iter_rows() for c in row)
        values_raw = _recalculate_xlsx(path) if has_formulas else raw
        values = load_workbook(io.BytesIO(values_raw), data_only=True, keep_links=False)
        metadata['hidden_sheets'] = [ws.title for ws in formulas if ws.sheet_state != 'visible']
        metadata['formulas_recalculated'] = has_formulas
        metadata['sheets'] = []
        metadata['visual_review'] = []
        for ws in formulas:
            if ws.sheet_state != 'visible':
                continue
            lines = []
            formula_cells = []
            for row in ws.iter_rows():
                cells = []
                for cell in row:
                    if cell.value is None:
                        continue
                    evaluated = values[ws.title][cell.coordinate]
                    value = evaluated.value
                    if evaluated.data_type == 'e':
                        raise ValueError(f'Ошибка формулы XLSX {ws.title}!{cell.coordinate}: {value}')
                    if cell.data_type == 'f':
                        formula_cells.append({'cell': cell.coordinate, 'formula': cell.value,
                                              'value': value.isoformat() if hasattr(value, 'isoformat') else value})
                        # Empty IFERROR results are genuinely empty; the formula
                        # remains in metadata, never presented as applicant data.
                        if value is None or value == '':
                            continue
                    else:
                        value = cell.value
                    cells.append(f'{cell.coordinate}: {value}')
                if cells:
                    lines.append(' | '.join(cells))
            metadata['sheets'].append({'name': ws.title, 'print_area': str(ws.print_area),
                                       'formula_cells': formula_cells})
            if lines:
                page('Лист: '+ws.title+'\n'+'\n'.join(lines), 'xlsx_cells')
            for number, picture in enumerate(ws._images, 1):
                content = picture._data()
                member = f'{ws.title}:image{number}'
                if reviewed_image(content, member):
                    continue
                attachment = {'sheet': ws.title, 'image': number, 'width': picture.width,
                              'height': picture.height,
                              'image_sha256': hashlib.sha256(content).hexdigest()}
                try:
                    with Image.open(io.BytesIO(content)) as im:
                        im.load()
                        text = ocr(im)
                    page(text, 'xlsx_image_ocr')
                    attachment['ocr'] = True
                except ValueError as error:
                    raise ValueError(f'Изображение XLSX {member} требует визуальной проверки '
                                     f'(sha256={attachment["image_sha256"]}): {error}') from error
                metadata['visual_review'].append(attachment)
        formulas.close()
        values.close()

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
            bad_member = z.testzip()
            if bad_member:
                raise ValueError('Повреждён ZIP-документ: CRC '+bad_member)
            render_members = [n for n in names if n.lower().endswith(('.wmf', '.emf', '.wdp', '.hdp', '.jxr'))
                              and not _nontext_review(hashlib.sha256(z.read(n)).hexdigest())]
            if render_members and 'word/document.xml' in names:
                metadata['document_conversion'] = {'method': 'libreoffice_pdf',
                                                     'reason': 'Встроенная векторная/HD-графика',
                                                     'members': render_members}
                convert('.docx')
            elif 'xl/workbook.xml' in names:
                spreadsheet()
            elif 'Index/Document.iwa' in names:
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
                    if reviewed_image(content, name):
                        continue
                    try:
                        im = Image.open(io.BytesIO(content))
                    except OSError:
                        if name.lower().endswith(('.png','.jpg','.jpeg','.tif','.tiff','.bmp')):
                            raise ValueError('Повреждено встроенное изображение '+name)
                        continue
                    with im:
                        im.load()
                        if im.width >= 600 and im.height >= 500:
                            try:
                                page(ocr(im), 'embedded_ocr')
                            except ValueError as error:
                                digest = hashlib.sha256(content).hexdigest()
                                raise ValueError(f'Встроенное изображение {name} требует визуальной '
                                                 f'проверки (sha256={digest}): {error}') from error
    elif raw.startswith((b'\xff\xd8', b'\x89PNG', b'GIF8', b'BM', b'II*', b'MM\x00')) or path.suffix.lower() in ('.webp', '.jfif'):
        with Image.open(path) as im:
            for index, frame in enumerate(ImageSequence.Iterator(im)):
                if _auxiliary_mpo(raw, im, index):
                    metadata.setdefault('auxiliary_frames_ignored', []).append(index+1)
                    continue
                proof = _page_review(h, index+1)
                if proof and proof['classification'] == 'verified_text':
                    page('[Ручная транскрипция после визуальной проверки изображения '
                         f'{index+1}]\n'+proof['transcription'], 'reviewed_text')
                    metadata.setdefault('page_reviews', []).append({
                        'source_sha256': h, 'physical_page': index+1,
                        'proof': proof, 'not_cv_quote': False})
                else:
                    page(ocr(frame.copy()), 'image_ocr')
    elif path.suffix.lower() in ('.txt', '.md'):
        page(decode(raw), 'text')
    elif b'<html' in raw[:1000].lower() or b'<!doctype' in raw[:1000].lower():
        page(BeautifulSoup(raw, 'html.parser').get_text('\n', strip=True), 'html')
    else:
        raise ValueError('Неподдерживаемый или повреждённый формат: '+path.suffix)
    if not pages or not any(p['text'].strip() for p in pages):
        raise ValueError('Пустой документ')
    result = {'sha256': h, 'pages': pages, 'original_name': path.name,
              'extraction_version': EXTRACTION_VERSION, 'complete': True,
              'metadata': metadata}
    _atomic_cache(cache, result)
    return result
