"""Rebuild one worksheet by patching XML; preserve source cells and styles."""
import copy
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
from openpyxl.utils import column_index_from_string, get_column_letter, range_boundaries
from .common import DATA, connect, dump, read, sha, stamp

S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
P = 'http://schemas.openxmlformats.org/package/2006/relationships'
C = 'http://schemas.openxmlformats.org/package/2006/content-types'

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

def remap_ranges(value, mapping):
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
        for group in groups: result.append(f'{get_column_letter(c1)}{group[0]}:{get_column_letter(min(c2,33))}{group[-1]}')
    return ' '.join(result)

def export(run, target=None):
    meta=read(run/'portable.json')
    c=connect(run/'state.sqlite')
    rows=c.execute("SELECT * FROM candidates WHERE review_status='reviewed' ORDER BY source_row").fetchall()
    c.close()
    path=template(run)
    target=Path(target) if target else DATA/'exports'/meta['output_name']
    target.parent.mkdir(parents=True,exist_ok=True)
    with zipfile.ZipFile(path) as z:
        items={n:z.read(n) for n in z.namelist()}
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
    write_cell(header,32,'Бан',1,style);write_cell(header,33,'Комментарий к бану',1,style)
    for dest,r in enumerate(rows,2):
        mapping[r['source_row']]=dest
        row=copy.deepcopy(old[r['source_row']]);row.set('r',str(dest));row.set('spans','1:33')
        for cell in list(row):
            col=re.match(r'[A-Z]+',cell.get('r')).group()
            if column_index_from_string(col)>33:row.remove(cell);continue
            cell.set('r',col+str(dest))
        result=json.loads(r['result'])
        rowstyle=next((x.get('s') for x in row if x.get('r')=='AE'+str(dest)),None)
        for col,val in [(1,'да'),(30,result['department']),(31,result.get('comment','')),(32,result['ban']),(33,result.get('ban_comment',''))]:
            write_cell(row,col,val,dest,rowstyle)
        row[:]=sorted(row,key=lambda x:column_index_from_string(re.match(r'[A-Z]+',x.get('r')).group()))
        data.append(row)
    for elem in list(sheet):
        tag=elem.tag.split('}')[-1]
        if tag=='dimension':elem.set('ref',f'A1:AG{len(rows)+1}')
        elif tag=='autoFilter':elem.set('ref',f'A1:AG{len(rows)+1}')
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
            ranges=remap_ranges(elem.get('sqref',''),mapping)
            if ranges:elem.set('sqref',ranges)
            else:sheet.remove(elem)
    # Both standard and x14 extension validation ranges are retained and remapped.
    for parent in sheet.iter():
        for elem in list(parent):
            if elem.tag.split('}')[-1]=='dataValidation':
                if 'sqref' in elem.attrib:
                    value=remap_ranges(elem.get('sqref'),mapping)
                    if value:elem.set('sqref',value)
                    else:parent.remove(elem)
                else:
                    refs=[x for x in elem.iter() if x.tag.split('}')[-1]=='sqref']
                    for ref in refs:ref.text=remap_ranges(ref.text or '',mapping)
                    if refs and not any(x.text for x in refs):parent.remove(elem)
        if parent.tag.split('}')[-1]=='dataValidations':
            parent.set('count',str(sum(x.tag.split('}')[-1]=='dataValidation' for x in parent)))
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
    verify(tmp,path,meta['sheet'],rows)
    if target.exists():shutil.copy2(target,target.with_name(target.stem+'.previous.xlsx'))
    try:os.replace(tmp,target)
    except PermissionError:
        target=target.with_name(target.stem+'.'+datetime.now().strftime('%Y%m%d_%H%M%S')+'.xlsx');os.replace(tmp,target)
    c=connect(run/'state.sqlite')
    c.execute('CREATE TABLE IF NOT EXISTS portable_exports(created TEXT,path TEXT,sha256 TEXT,count INTEGER)')
    with c:
        c.execute('INSERT INTO portable_exports VALUES(?,?,?,?)',(stamp(),str(target.name),sha(target),len(rows)))
    c.close()
    return {'path':str(target),'rows':len(rows),'sha256':sha(target)}

def verify(path,source,sheet_name,rows):
    a=load_workbook(source,data_only=False); b=load_workbook(path,data_only=False)
    try:
        assert len(b.worksheets)==1
        s=a[sheet_name];t=b.worksheets[0]
        assert t.max_row==len(rows)+1 and t.max_column==33
        seen=set()
        for dest,r in enumerate(rows,2):
            ident=str(t.cell(dest,2).value);assert ident==r['id'] and ident not in seen;seen.add(ident)
            for col in range(2,30):
                x=s.cell(r['source_row'],col);y=t.cell(dest,col)
                assert x.value==y.value,(ident,col,'value')
                assert x.data_type==y.data_type or x.value is None,(ident,col,'type')
                assert x.number_format==y.number_format,(ident,col,'format')
                for attr in ('font','fill','border','alignment','protection'):
                    assert copy.copy(getattr(x,attr))==copy.copy(getattr(y,attr)),(ident,col,attr)
                for attr in ('target','location','tooltip','display'):
                    assert (getattr(x.hyperlink,attr) if x.hyperlink else None)==(getattr(y.hyperlink,attr) if y.hyperlink else None),(ident,col,'hyperlink')
            result=json.loads(r['result'])
            for col,value in [(1,'да'),(30,result['department']),(31,result.get('comment','')),(32,result['ban']),(33,result.get('ban_comment',''))]:
                assert (t.cell(dest,col).value or '')==(value or ''),(ident,col,'result')
            if result['ban']=='да':assert result['department']=='Бан' and not result.get('clusters')
            clusters=result.get('clusters',[])
            assert len(clusters)<=3 and len({x['cluster'] for x in clusters})==len(clusters)
            assert all(60<=x['score']<=100 for x in clusters)
            assert [x['score'] for x in clusters]==sorted((x['score'] for x in clusters),reverse=True)
    finally:a.close();b.close()
