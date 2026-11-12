"""Rebuild one worksheet by patching XML; preserve source cells and styles."""
import copy
import hashlib
import json
import os
import posixpath
import re
import shutil
import zipfile
from datetime import datetime
from pathlib import Path
from lxml import etree as E

from openpyxl import load_workbook
from openpyxl.styles.cell_style import CellStyle, StyleArray
from openpyxl.utils import column_index_from_string, get_column_letter, range_boundaries
from openpyxl.xml.functions import tostring as spreadsheet_xml
from .common import DATA, connect, dump, read, sha, stamp

S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
P = 'http://schemas.openxmlformats.org/package/2006/relationships'
C = 'http://schemas.openxmlformats.org/package/2006/content-types'
APPEND_HEADERS = ('ИИ', 'Подразделение (HR Hive)', 'Комментарий (HR Hive)',
                  'Бан', 'Комментарий к бану', 'Статус обработки')
GENERATED_SOURCE_HEADERS = {1: 'ID', 2: 'ФИО из резюме',
                            11: 'Телефон из резюме', 13: 'Образование из резюме'}
_SOURCE_VALUE = object()


def _generated_layout(meta):
    enabled = meta.get('generated_from_resumes', False)
    if enabled is False:
        return False
    if (enabled is not True or meta.get('export_layout') != 'append_results' or
            type(meta.get('id_column')) is not int or meta['id_column'] != 1 or
            type(meta.get('source_column_count')) is not int or meta['source_column_count'] != 38):
        raise ValueError('generated_from_resumes требует append_results: ID в A и 38 исходных колонок')
    return True


def _generated_source_projection(source, sheet_name, rows, meta, projected_styles=None):
    """Project saved CV facts into empty generated-template cells, never the source."""
    if not _generated_layout(meta):
        return {}
    workbook = load_workbook(source, data_only=False)
    try:
        sheet = workbook[sheet_name]
        if (sheet.max_column != 38 or
                any(sheet.cell(1, col).value != title for col, title in GENERATED_SOURCE_HEADERS.items())):
            raise ValueError('Чужая схема шаблона generated_from_resumes: проверьте точные заголовки')
        projection = {}
        style_indexes = {tuple(style): index for index, style in enumerate(workbook._cell_styles)}
        for row in rows:
            if row['review_status'] != 'reviewed':
                continue
            card = json.loads(row['result']).get('fact_card') or {}
            if not isinstance(card, dict):
                raise ValueError('Некорректная fact_card сохранённого результата')
            name, education = card.get('name') or '', card.get('education') or ''
            phones = card.get('phones') or []
            if (not isinstance(name, str) or not isinstance(education, str) or
                    not isinstance(phones, list) or
                    any(not isinstance(phone, (str, int)) or isinstance(phone, bool) for phone in phones)):
                raise ValueError('Некорректные ФИО, телефоны или образование в fact_card')
            values = {2: name, 11: '; '.join(str(phone) for phone in phones if str(phone)),
                      13: education}
            for col, value in values.items():
                original = sheet.cell(row['source_row'], col)
                if original.value in (None, '') and value:
                    key = (row['source_row'], col)
                    projection[key] = value
                    if projected_styles is not None:
                        # A missing XML cell has an implicit zero StyleArray even
                        # when the file's cellXfs[0] uses a custom number format.
                        source_style = original._style or StyleArray()
                        style_key = tuple(source_style)
                        index = style_indexes.get(style_key)
                        xml = None
                        if index is None:
                            xf = CellStyle.from_array(source_style)
                            if source_style.alignmentId:
                                xf.alignment = workbook._alignments[source_style.alignmentId]
                            if source_style.protectionId:
                                xf.protection = workbook._protections[source_style.protectionId]
                            node = E.fromstring(spreadsheet_xml(xf.to_tree()))
                            for child in node.iter():
                                if not child.tag.startswith('{'):
                                    child.tag = q(child.tag)
                            xml = E.tostring(node)
                        projected_styles[key] = (style_key, index, xml)
        return projection
    finally:
        workbook.close()


def export_comments(run, rows):
    """Apply reviewed wording corrections to the export, preserving saved assessments."""
    path = run / 'export_comments.json'
    if not path.exists():
        return rows
    corrections = read(path)
    indexed = {row['id']: row for row in rows}
    for ident, correction in corrections.items():
        if ident not in indexed or indexed[ident]['review_status'] != 'reviewed':
            raise ValueError('Правка комментария относится к несохранённой карточке: ' + ident)
        row = dict(indexed[ident])
        result = json.loads(row['result'])
        original = result.get('comment', '')
        if hashlib.sha256(original.encode('utf-8')).hexdigest() != correction['source_sha256']:
            raise ValueError('Исходный комментарий изменился: ' + ident)
        result['comment'] = correction['text']
        row['result'] = json.dumps(result, ensure_ascii=False)
        indexed[ident] = row
    return [indexed[row['id']] for row in rows]


def append_values(row):
    """Do not represent an unread/unsubmitted candidate as a checked non-ban."""
    if row['review_status'] == 'reviewed':
        result = json.loads(row['result'])
        return ('да', result['department'], result.get('comment', ''),
                result['ban'], result.get('ban_comment', ''), 'Завершено')
    if row['extraction_status'] == 'error':
        status = 'Ошибка извлечения: ' + (row['error'] or 'Причина не сохранена')
    elif row['extraction_status'] == 'ready':
        status = 'Ожидает проверки'
    else:
        status = 'Ожидает извлечения'
    return ('', '', '', '', '', status)

def q(name): return '{'+S+'}'+name

def template(run):
    meta = read(run / 'portable.json')
    return run / meta['source_book']

def source_sheet(run):
    return read(run / 'portable.json')['sheet']

def write_cell(row, col, value, index, style=None):
    ref = get_column_letter(col)+str(index)
    cell = next((c for c in row if c.get('r') == ref), None)
    if cell is None:
        cell = E.SubElement(row, q('c'), {'r':ref})
        if style is not None: cell.set('s', style)
    for child in list(cell): cell.remove(child)
    cell.set('t', 'inlineStr')
    inline=E.SubElement(cell,q('is'));text=E.SubElement(inline,q('t'));text.text=str(value or '')
    text.set('{http://www.w3.org/XML/1998/namespace}space','preserve')

def remap_ranges(value, mapping, max_column=33):
    result=[]
    for token in value.split():
        try: c1,r1,c2,r2=range_boundaries(token)
        except ValueError: continue
        if not c1 or not c2: continue
        destinations=sorted(d for s,d in mapping.items() if (r1 or 1)<=s<=(r2 or 1048576))
        if not destinations: continue
        groups=[]
        for n in destinations:
            if groups and n==groups[-1][-1]+1: groups[-1].append(n)
            else: groups.append([n])
        if c1 > max_column: continue
        for group in groups: result.append(f'{get_column_letter(c1)}{group[0]}:{get_column_letter(min(c2,max_column))}{group[-1]}')
    return ' '.join(result)

def export(run, target=None):
    meta=read(run/'portable.json')
    _generated_layout(meta)
    append = meta.get('export_layout') == 'append_results'
    source_columns = int(meta.get('source_column_count', 38)) if append else 29
    max_column = source_columns + len(APPEND_HEADERS) if append else 33
    end_column = get_column_letter(max_column)
    c=connect(run/'state.sqlite')
    rows=c.execute("SELECT * FROM candidates ORDER BY source_row" if append else
                   "SELECT * FROM candidates WHERE review_status='reviewed' ORDER BY source_row").fetchall()
    c.close()
    if append:
        rows = export_comments(run, rows)
    path=template(run)
    projected_styles = {}
    projection = _generated_source_projection(path, meta['sheet'], rows, meta, projected_styles)
    target=Path(target) if target else Path(meta['output_path']) if meta.get('output_path') else DATA/'exports'/meta['output_name']
    if _generated_layout(meta) and path.resolve() in {
            target.resolve(), target.with_name(target.stem+'.tmp.xlsx').resolve(),
            target.with_name(target.stem+'.previous.xlsx').resolve()}:
        raise ValueError('Экспорт generated_from_resumes не может изменять исходный шаблон')
    target.parent.mkdir(parents=True,exist_ok=True)
    with zipfile.ZipFile(path) as z:
        items={n:z.read(n) for n in z.namelist()}
    wrapped_styles = {}
    styles = E.fromstring(items['xl/styles.xml']) if append else None
    cell_styles = styles.find(q('cellXfs')) if append else None
    added_source_styles = {}
    def projected_style(key):
        style_key, index, xml = projected_styles[key]
        if index is not None:
            return str(index)
        if style_key not in added_source_styles:
            added_source_styles[style_key] = str(len(cell_styles))
            cell_styles.append(E.fromstring(xml))
        return added_source_styles[style_key]
    def wrapped_style(index):
        key = int(index or 0)
        if key not in wrapped_styles:
            xf = copy.deepcopy(cell_styles[key])
            alignment = xf.find(q('alignment'))
            if alignment is None:
                alignment = E.Element(q('alignment'))
                xf.insert(0, alignment)
            alignment.set('wrapText', '1')
            xf.set('applyAlignment', '1')
            wrapped_styles[key] = str(len(cell_styles))
            cell_styles.append(xf)
        return wrapped_styles[key]
    workbook=E.fromstring(items['xl/workbook.xml']);rels=E.fromstring(items['xl/_rels/workbook.xml.rels'])
    sheets=workbook.find(q('sheets'));selected=next(s for s in sheets if s.get('name')==meta['sheet'])
    rid=selected.get('{'+R+'}id');sheetrel=next(x for x in rels if x.get('Id')==rid)
    part=sheetrel.get('Target');part=part.lstrip('/') if part.startswith('/') else posixpath.normpath(posixpath.join('xl',part))
    for s in list(sheets):
        if s is not selected:sheets.remove(s)
    for rel in list(rels):
        if rel.get('Type','').endswith('/worksheet') and rel.get('Id')!=rid:
            t=rel.get('Target');t=t.lstrip('/') if t.startswith('/') else posixpath.normpath(posixpath.join('xl',t))
            items.pop(t,None);items.pop(posixpath.dirname(t)+'/_rels/'+posixpath.basename(t)+'.rels',None);rels.remove(rel)
        if rel.get('Type','').endswith('/calcChain'):
            items.pop('xl/calcChain.xml',None);rels.remove(rel)
    names=workbook.find(q('definedNames'))
    if names is not None: workbook.remove(names)
    for view in workbook.findall('.//'+q('workbookView')): view.set('activeTab','0');view.set('firstSheet','0')
    sheet=E.fromstring(items[part]);data=sheet.find(q('sheetData'))
    old={int(row.get('r')):row for row in data};mapping={1:1}
    for row in list(data):data.remove(row)
    header=copy.deepcopy(old[1]);data.append(header)
    style=next((x.get('s') for x in header if x.get('r')=='AE1'),None)
    if append:
        style = next((x.get('s') for x in header if x.get('r') == get_column_letter(source_columns)+'1'), style)
        style = wrapped_style(style)
        for col, title in enumerate(APPEND_HEADERS, source_columns + 1):
            write_cell(header, col, title, 1, style)
    else:
        write_cell(header,32,'Бан',1,style);write_cell(header,33,'Комментарий к бану',1,style)
    for dest,r in enumerate(rows,2):
        mapping[r['source_row']]=dest
        row=copy.deepcopy(old[r['source_row']]);row.set('r',str(dest));row.set('spans',f'1:{max_column}')
        for cell in list(row):
            col=re.match(r'[A-Z]+',cell.get('r')).group()
            if column_index_from_string(col)>max_column:row.remove(cell);continue
            cell.set('r',col+str(dest))
        rowstyle=next((x.get('s') for x in row if x.get('r')=='AE'+str(dest)),None)
        if append:
            rowstyle = wrapped_style(rowstyle)
            values = list(enumerate(append_values(r), source_columns + 1))
            for col in (2, 11, 13):
                key = (r['source_row'], col)
                if key in projection:
                    write_cell(row, col, projection[key], dest, projected_style(key))
        else:
            result=json.loads(r['result'])
            values = [(1,'да'),(30,result['department']),(31,result.get('comment','')),(32,result['ban']),(33,result.get('ban_comment',''))]
        for col,val in values:
            write_cell(row,col,val,dest,rowstyle)
        row[:]=sorted(row,key=lambda x:column_index_from_string(re.match(r'[A-Z]+',x.get('r')).group()))
        data.append(row)
    if append:
        # The source widths stay intact; new result columns must fit Russian text.
        columns = sheet.find(q('cols'))
        if columns is None:
            columns = E.Element(q('cols'))
            sheet.insert(list(sheet).index(data), columns)
        for column in list(columns):
            if int(column.get('min')) > source_columns:
                columns.remove(column)
            elif int(column.get('max')) > source_columns:
                column.set('max', str(source_columns))
        for col, width in enumerate((8, 46, 80, 10, 90, 48), source_columns + 1):
            E.SubElement(columns, q('col'), min=str(col), max=str(col),
                         width=str(width), customWidth='1')
        cell_styles.set('count', str(len(cell_styles)))
        items['xl/styles.xml'] = E.tostring(styles, encoding='utf-8', xml_declaration=True)
    for elem in list(sheet):
        tag=elem.tag.split('}')[-1]
        if tag=='dimension':elem.set('ref',f'A1:{end_column}{len(rows)+1}')
        elif tag=='autoFilter':elem.set('ref',f'A1:{end_column}{len(rows)+1}')
        elif tag in ('tableParts','drawing','legacyDrawing','legacyDrawingHF'):
            sheet.remove(elem)
        elif tag=='hyperlinks':
            for link in list(elem):
                ref=link.get('ref','');m=re.fullmatch(r'([A-Z]+)(\d+)',ref)
                if not m or int(m.group(2)) not in mapping:elem.remove(link)
                else:link.set('ref',m.group(1)+str(mapping[int(m.group(2))]))
        elif tag=='mergeCells':
            for merge in list(elem):
                a,b,c1,d=range_boundaries(merge.get('ref'))
                if b in mapping and d in mapping and mapping[d]-mapping[b]==d-b:
                    merge.set('ref',f'{get_column_letter(a)}{mapping[b]}:{get_column_letter(c1)}{mapping[d]}')
                else:elem.remove(merge)
        elif tag=='conditionalFormatting':
            ranges=remap_ranges(elem.get('sqref',''),mapping,max_column)
            if ranges:elem.set('sqref',ranges)
            else:sheet.remove(elem)
    # Both standard and x14 extension validation ranges are retained and remapped.
    for parent in sheet.iter():
        for elem in list(parent):
            if elem.tag.split('}')[-1]=='dataValidation':
                if 'sqref' in elem.attrib:
                    value=remap_ranges(elem.get('sqref'),mapping,max_column)
                    if value:elem.set('sqref',value)
                    else:parent.remove(elem)
                else:
                    refs=[x for x in elem.iter() if x.tag.split('}')[-1]=='sqref']
                    for ref in refs:ref.text=remap_ranges(ref.text or '',mapping,max_column)
                    if refs and not any(x.text for x in refs):parent.remove(elem)
        if parent.tag.split('}')[-1]=='dataValidations':
            parent.set('count',str(sum(x.tag.split('}')[-1]=='dataValidation' for x in parent)))
    # Do not carry unused candidate strings or unreferenced hyperlink addresses.
    shared_part='xl/sharedStrings.xml'
    if shared_part in items:
        shared=E.fromstring(items[shared_part]);entries=list(shared)
        cells=[x for x in sheet.iter(q('c')) if x.get('t')=='s']
        used=sorted({int(x.find(q('v')).text) for x in cells})
        indexes={old:new for new,old in enumerate(used)}
        for child in list(shared):shared.remove(child)
        for index in used:shared.append(entries[index])
        for cell in cells:cell.find(q('v')).text=str(indexes[int(cell.find(q('v')).text)])
        shared.set('count',str(len(cells)));shared.set('uniqueCount',str(len(used)))
        items[shared_part]=E.tostring(shared,encoding='utf-8',xml_declaration=True)
    sheet_rels=posixpath.dirname(part)+'/_rels/'+posixpath.basename(part)+'.rels'
    if sheet_rels in items:
        sr=E.fromstring(items[sheet_rels]);used_rids={x.get('{'+R+'}id') for x in sheet.iter() if x.get('{'+R+'}id')}
        for rel in list(sr):
            if rel.get('Id') not in used_rids:sr.remove(rel)
        items[sheet_rels]=E.tostring(sr,encoding='utf-8',xml_declaration=True)
    items[part]=E.tostring(sheet,encoding='utf-8',xml_declaration=True)
    items['xl/workbook.xml']=E.tostring(workbook,encoding='utf-8',xml_declaration=True)
    items['xl/_rels/workbook.xml.rels']=E.tostring(rels,encoding='utf-8',xml_declaration=True)
    ct=E.fromstring(items['[Content_Types].xml'])
    for x in list(ct):
        if x.get('PartName','').lstrip('/') not in items and x.tag.endswith('Override'):ct.remove(x)
    items['[Content_Types].xml']=E.tostring(ct,encoding='utf-8',xml_declaration=True)
    tmp=target.with_name(target.stem+'.tmp.xlsx')
    with zipfile.ZipFile(tmp,'w',zipfile.ZIP_DEFLATED) as z:
        for n,raw in items.items():z.writestr(n,raw)
    verify(tmp,path,meta['sheet'],rows,meta)
    try:
        if target.exists():shutil.copy2(target,target.with_name(target.stem+'.previous.xlsx'))
        os.replace(tmp,target)
    except PermissionError:
        target=target.with_name(target.stem+'.'+datetime.now().strftime('%Y%m%d_%H%M%S')+'.xlsx');os.replace(tmp,target)
    c=connect(run/'state.sqlite')
    c.execute('CREATE TABLE IF NOT EXISTS portable_exports(created TEXT,path TEXT,sha256 TEXT,count INTEGER)')
    with c:
        c.execute('INSERT INTO portable_exports VALUES(?,?,?,?)',(stamp(),str(target.name),sha(target),len(rows)))
    c.close()
    return {'path':str(target),'rows':len(rows),'sha256':sha(target)}

def verify(path,source,sheet_name,rows,meta=None):
    meta = meta or {}
    projection = _generated_source_projection(source, sheet_name, rows, meta)
    append = meta.get('export_layout') == 'append_results'
    source_columns = int(meta.get('source_column_count', 38)) if append else 29
    id_column = int(meta.get('id_column', 1)) if append else 2
    max_column = source_columns + len(APPEND_HEADERS) if append else 33
    a=load_workbook(source,data_only=False); b=load_workbook(path,data_only=False)
    try:
        assert len(b.worksheets)==1
        s=a[sheet_name];t=b.worksheets[0]
        assert t.max_row==len(rows)+1 and t.max_column==max_column
        if append:
            assert len({r['source_row'] for r in rows}) == len(rows), 'Повтор исходной строки'
            assert [r['source_row'] for r in rows] == sorted(r['source_row'] for r in rows)
            for col in range(1, source_columns + 1):
                _verify_source_cell(s.cell(1, col), t.cell(1, col), 'header', col)
            assert tuple(t.cell(1, col).value for col in range(source_columns + 1, max_column + 1)) == APPEND_HEADERS
        seen=set()
        for dest,r in enumerate(rows,2):
            ident=str(t.cell(dest,id_column).value);assert ident==r['id'] and ident not in seen;seen.add(ident)
            for col in range(1 if append else 2,source_columns + 1):
                x=s.cell(r['source_row'],col);y=t.cell(dest,col)
                _verify_source_cell(x, y, ident, col,
                                    projection.get((r['source_row'], col), _SOURCE_VALUE))
            if append:
                for col, value in enumerate(append_values(r), source_columns + 1):
                    assert (t.cell(dest,col).value or '') == (value or ''), (ident,col,'result/status')
                if r['review_status'] != 'reviewed':
                    continue
            result=json.loads(r['result'])
            if not append:
                for col,value in [(1,'да'),(30,result['department']),(31,result.get('comment','')),(32,result['ban']),(33,result.get('ban_comment',''))]:
                    assert (t.cell(dest,col).value or '')==(value or ''),(ident,col,'result')
            if result['ban']=='да':assert result['department']=='Бан' and not result.get('clusters')
            clusters=result.get('clusters',[])
            assert len(clusters)<=3 and len({x['cluster'] for x in clusters})==len(clusters)
            assert all(60<=x['score']<=100 for x in clusters)
            assert [x['score'] for x in clusters]==sorted((x['score'] for x in clusters),reverse=True)
        assert seen == {r['id'] for r in rows}
    finally:a.close();b.close()


def _verify_source_cell(source, exported, ident, col, expected_value=_SOURCE_VALUE):
    expected = source.value if expected_value is _SOURCE_VALUE else expected_value
    assert expected==exported.value,(ident,col,'value')
    if expected_value is not _SOURCE_VALUE:
        assert exported.data_type=='s',(ident,col,'projected type')
    assert source.data_type==exported.data_type or source.value is None,(ident,col,'type')
    assert source.number_format==exported.number_format,(ident,col,'format')
    for attr in ('font','fill','border','alignment','protection'):
        assert copy.copy(getattr(source,attr))==copy.copy(getattr(exported,attr)),(ident,col,attr)
    for attr in ('target','location','tooltip','display'):
        assert (getattr(source.hyperlink,attr) if source.hyperlink else None)==(getattr(exported.hyperlink,attr) if exported.hyperlink else None),(ident,col,'hyperlink')
