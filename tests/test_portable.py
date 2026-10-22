import copy
import io
import json
import os
import sqlite3
import subprocess
import tarfile
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest.mock import patch

import fitz
from PIL import Image,ImageDraw,ImageFont
from lxml import etree as E
from openpyxl import Workbook,load_workbook
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.styles import PatternFill

from app import common,documents,service,books,archives
from app.engine.profiles import assess
from app.engine.pilot150_bans import end_date,StructuredBans

class PortableTest(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory();self.root=Path(self.temp.name)
        self.patches=[patch.object(m,'DATA',self.root) for m in (common,documents,service,books,archives)]
        for p in self.patches:p.start()
        common.layout()
        wb=Workbook();s=wb.active;s.title='Матка №1'
        s.append(['ИИ','ID']+['Поле '+str(i) for i in range(3,30)]+['Подразделение','Комментарий'])
        for ident in ('101','102','103'):
            s.append([None,ident]+['значение']*27+['старый результат','старый комментарий'])
        s['D2'].hyperlink='https://example.test/unprocessed-101'
        s['C3']='=1+2';s['D3'].hyperlink='https://example.test/путь';s['D3'].number_format='@';s['D3'].fill=PatternFill('solid',fgColor='00FF00')
        dv=DataValidation(type='list',formula1='"да,нет"');dv.add('E2:E4');s.add_data_validation(dv)
        wb.create_sheet('Другой лист');wb.save(self.root/'sources'/'Матка №1.xlsx')
        self.run=service.new_run();self.source=service.source_rows(self.run)

    def tearDown(self):
        for p in reversed(self.patches):p.stop()
        self.temp.cleanup()

    def candidate(self,ident='102',done=False):
        c=common.connect(self.run/'state.sqlite')
        result={'department':'УПИ 88','comment':'Уточнить опыт переговоров','ban':'нет','ban_comment':'','clusters':[{'cluster':'УПИ','score':88}]}
        with c:c.execute('INSERT INTO candidates(id,source_row,files,fingerprint,extraction_status,text,review_status,result) VALUES(?,?,?,?,?,?,?,?)',(ident,self.source[ident][0],'[]','hash','ready','Полный текст кандидата\n' * 20,'reviewed' if done else 'pending',json.dumps(result,ensure_ascii=False) if done else None))
        c.close()

    def test_export_source_fields_links_validation_and_recovery(self):
        self.candidate(done=True)
        result=books.export(self.run)
        p=Path(result['path']);w=load_workbook(p);s=w.worksheets[0]
        self.assertEqual(w.sheetnames,['Матка №1']);self.assertEqual(s.max_row,2)
        self.assertEqual(s['B2'].value,'102');self.assertEqual(s['C2'].value,'=1+2')
        self.assertEqual(s['D2'].hyperlink.target,'https://example.test/путь')
        self.assertEqual(str(s.data_validations.dataValidation[0].sqref),'E2')
        self.assertEqual(s['AE2'].value,'Уточнить опыт переговоров');w.close()
        with zipfile.ZipFile(p) as z:
            self.assertNotIn(b'unprocessed-101',b''.join(z.read(n) for n in z.namelist()))
        # A crash after transaction and before workbook export loses no result.
        p.unlink();self.assertEqual(books.export(self.run)['rows'],1)
        real_replace=os.replace
        calls=[]
        def locked_replace(a,b):
            calls.append(b)
            if len(calls)==1:raise PermissionError()
            return real_replace(a,b)
        with patch('app.books.os.replace',side_effect=locked_replace):
            other=books.export(self.run)
            self.assertNotEqual(Path(other['path']).name,p.name)
        with patch('app.books.shutil.copy2',side_effect=PermissionError()):
            other=books.export(self.run)
            self.assertNotEqual(Path(other['path']).name,p.name)

    def test_scan_missing_preserves_completed_unicode_and_change(self):
        self.candidate(done=True)
        first=service.scan();self.assertEqual(first['new'],[])
        c=common.connect(self.run/'state.sqlite');self.assertEqual(c.execute("SELECT review_status FROM candidates WHERE id='102'").fetchone()[0],'reviewed');c.close()
        sub=self.root/'cvs'/'Папка с пробелами';sub.mkdir();p=sub/'101.txt';p.write_text('Пример полного текста резюме',encoding='utf-8')
        self.assertEqual(service.scan()['new'],['101'])
        service.extract();claims=service.claim('codex',5);self.assertIn('101',claims)
        service.show('101',0,5)
        card={'id':'101','full_text_review_complete':True,'reviewed_text_sha256':'wrong'}
        with self.assertRaisesRegex(ValueError,'sha256'):service.preview(card)
        full=service.show('101',5,20000)
        p.rename(sub/'101.md');self.assertEqual(service.scan()['changed'],[])
        (sub/'101.md').write_text('Полностью изменённое резюме',encoding='utf-8')
        self.assertEqual(service.scan()['changed'],['101'])

    def test_partial_read_and_english_rejected(self):
        self.candidate();service.claim('codex',5);part=service.show('102',0,10)
        card={'id':'102','full_text_review_complete':True,'reviewed_text_sha256':part['text_sha256'],'jobs':[],'tracks':{},'reviews':[],'comment':'Short experience'}
        with self.assertRaisesRegex(ValueError,'весь текст'):service.preview(card)
        service.show('102',10,20000)
        with self.assertRaisesRegex(ValueError,'русский'):service.preview(card)

    def test_legacy_complete_text_is_reused_by_content_hash(self):
        p=self.root/'cvs'/'101.txt';p.write_text('Полный текст документа для синтетической проверки кэша.',encoding='utf-8')
        service.scan();h=common.sha(p);pages=[{'page':1,'method':'text','text':p.read_text('utf-8')}]
        c=common.connect(self.run/'state.sqlite')
        with c:c.execute('INSERT INTO documents VALUES(?,?,?,?,?,?,?)',(h,'old/windows/path','text','ready',json.dumps(pages,ensure_ascii=False),None,common.stamp()))
        c.close()
        with patch('app.service.extract_file',side_effect=AssertionError('Повторное извлечение не требуется')):
            service.extract()
        c=common.connect(self.run/'state.sqlite');row=c.execute("SELECT extraction_status,text FROM candidates WHERE id='101'").fetchone();c.close()
        self.assertEqual(row[0],'ready');self.assertIn(pages[0]['text'],row[1])

    def test_new_version_rebinds_rows_and_duplicate_ids_rejected(self):
        self.candidate(done=True)
        p=self.root/'sources'/'Матка №1.xlsx';wb=load_workbook(p);ws=wb['Матка №1']
        ws['B2']='102';ws['B3']='101';wb.save(p);wb.close()
        service.scan(new_version=True);run=common.active();c=common.connect(run/'state.sqlite')
        row=c.execute("SELECT source_row,review_status,text,result FROM candidates WHERE id='102'").fetchone();c.close()
        self.assertEqual(row[0],2);self.assertEqual(row[1],'pending');self.assertTrue(row[2]);self.assertIsNone(row[3])
        p=books.template(run);wb=load_workbook(p);ws=wb['Матка №1'];ws['B3']='102';wb.save(p);wb.close()
        with self.assertRaisesRegex(ValueError,'Повторный ID'):service.source_rows(run)

    def test_backup_strips_cv_blobs_and_restores(self):
        self.candidate(done=True)
        c=sqlite3.connect(self.run/'state.sqlite');c.execute('CREATE TABLE source_files(hash TEXT,bytes BLOB)');c.execute('INSERT INTO source_files VALUES(?,?)',('h',b'%PDF ORIGINAL CV'));c.commit();c.close()
        (self.root/'cvs'/'101.pdf').write_bytes(b'%PDF ORIGINAL CV')
        archive=self.root/'backups'/'test.tar.gz';archives.archive_tree(self.root,archive)
        target=self.root/'empty_restore';archives.restore(archive,target)
        self.assertFalse((target/'cvs'/'101.pdf').exists())
        c=sqlite3.connect(target/'state/runs'/self.run.name/'state.sqlite');self.assertEqual(c.execute('SELECT COUNT(*) FROM source_files').fetchone()[0],0);self.assertEqual(c.execute("SELECT COUNT(*) FROM candidates WHERE review_status='reviewed'").fetchone()[0],1);c.close()
        with self.assertRaises(ValueError):archives.restore(archive,target)

    def test_scores_unrestricted_maximum_no_age_and_fpga(self):
        profiles={'X':{'cluster':'А','profile':'профиль','education':'Любое','experience':'Любой'},'Y':{'cluster':'А','profile':'другой','education':'Любое','experience':'Любой'},'Z':{'cluster':'КБ','profile':'ПЛИС','education':'Любое','experience':'Любой'}}
        card={'id':'synthetic','tracks':{'fpga':['junior',1,'Цитата','задачи']},'reviews':[['X',[35,25,0,0],'fpga','факты',''],['Y',[30,20,0,0],'fpga','факты',''],['Z',[35,25,0,0],'fpga','факты','']]}
        _,rejected,result=assess(card,profiles);self.assertEqual(len(result),2);self.assertFalse(rejected);self.assertEqual(result[0]['score'],86)
        changed=copy.deepcopy(card);changed.update(age=90,sex='female',birth_date='1930-01-01')
        self.assertEqual(assess(card,profiles),assess(changed,profiles))

    def test_termination_upper_boundary(self):
        self.assertEqual(str(end_date('2024')),'2024-12-31')
        self.assertEqual(str(end_date('2024-02')),'2024-02-29')
        self.assertIsNone(end_date('present'))

    def ban_sources(self):
        folder=self.run/'inputs'/'banlist';folder.mkdir()
        (folder/'ФИО.csv').write_text('ФИО;Телефон\nНесуществующий Тестовый Персонаж;\n',encoding='utf-8')
        rows=['Компания;Срок']
        for n in range(2,350):rows.append(('Тестовый запрет;1 год' if n==10 else 'Компания'+str(n)+';Свободно'))
        (folder/'Компании.csv').write_text('\n'.join(rows),encoding='utf-8')
        meta=common.read(self.run/'portable.json');meta['rules_hash']=service.rules_hash(self.run/'inputs');common.dump(self.run/'portable.json',meta)
        meta=common.read(self.run/'manifest.json');meta['as_of']='2026-10-22';meta['input_hashes']={p.relative_to(self.run/'inputs').as_posix():common.sha(p) for p in (self.run/'inputs').rglob('*') if p.is_file()};common.dump(self.run/'manifest.json',meta)

    def test_bans_employment_only_term_and_transneft(self):
        self.ban_sources();engine=StructuredBans(self.run)
        card={'name':'Синтетический Кандидат','phones':[],'jobs':[],'education':'Транснефть'}
        source=[None]*33
        self.assertEqual(engine.assess(card,source)['ban'],'нет')
        card['jobs']=[['Транснефть-Приволга','2020','2021','Инженер','Самара','private']]
        self.assertEqual(engine.assess(card,source)['ban'],'да')
        card['jobs']=[['Тестовый запрет','2020','2024-01','Инженер','Самара','private']]
        self.assertEqual(engine.assess(card,source)['ban'],'нет')
        card['jobs'][0][2]='2026-01';self.assertEqual(engine.assess(card,source)['ban'],'да')

    def test_submit_transaction_idempotent_and_export(self):
        self.ban_sources();self.candidate();service.claim('codex',1);shown=service.show('102')
        card={'id':'102','name':'Синтетический Кандидат','phones':[],'education':'Подтверждённое образование','jobs':[],'tracks':{},'reviews':[],'comment':'Нет подтверждённого профильного опыта','reviewed_text_sha256':shown['text_sha256'],'full_text_review_complete':True}
        with patch('app.service.catalog',return_value={}):
            result=service.submit(card,'codex');self.assertIn('export',result)
            self.assertTrue(service.submit(card,'codex')['already_completed'])
        c=common.connect(self.run/'state.sqlite');self.assertEqual(c.execute("SELECT review_status FROM candidates WHERE id='102'").fetchone()[0],'reviewed');c.close()

    def test_full_document_formats_and_ocr(self):
        sentence='Синтетическое резюме. Инженер выполняет монтаж оборудования и контроль качества.'
        p=self.root/'101.txt';p.write_text(sentence,encoding='utf-8');self.assertEqual(documents.extract_file(p)['pages'][0]['text'],sentence)
        rtf=self.root/'102.rtf';rtf.write_text('{\\rtf1\\ansi Synthetic resume with engineering experience.}',encoding='ascii');self.assertIn('engineering',documents.extract_file(rtf)['pages'][0]['text'])
        docx=self.root/'103.docx'
        with zipfile.ZipFile(docx,'w') as z:z.writestr('word/document.xml','<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:tbl><w:tr><w:tc><w:p><w:r><w:t>'+sentence+'</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:body></w:document>')
        self.assertIn(sentence,documents.extract_file(docx)['pages'][0]['text'])
        p=self.root/'104.pdf';pdf=fitz.open()
        for n in range(2):
            pg=pdf.new_page();pg.insert_text((70,70),'Synthetic resume page '+str(n+1)+'. Engineering experience in production and control. '*3)
        pdf.save(p);pdf.close();self.assertEqual(len(documents.extract_file(p)['pages']),2)
        if not shutil_which('tesseract'):return
        font=ImageFont.truetype('/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf',36)
        image=Image.new('RGB',(1700,600),'white');draw=ImageDraw.Draw(image)
        draw.text((40,80),'Синтетическое резюме инженера\nМонтаж оборудования и контроль качества\nОпыт работы на производстве',font=font,fill='black')
        ip=self.root/'105.png';image.save(ip);self.assertIn('инженера',documents.extract_file(ip)['pages'][0]['text'].lower())
        pdf=fitz.open()
        for n in range(2):
            pg=pdf.new_page(width=850,height=300);pg.insert_image(pg.rect,filename=ip)
        p=self.root/'106.pdf';pdf.save(p);pdf.close();res=documents.extract_file(p);self.assertEqual(len(res['pages']),2);self.assertTrue(all(x['method']=='ocr' for x in res['pages']))
        p=self.root/'bad.crdownload';p.write_bytes(b'incomplete');self.assertRaises(ValueError,documents.extract_file,p)

    def test_office_conversion_odt_pptx_and_unicode_folder(self):
        sentence='Synthetic engineering resume with practical production and equipment maintenance experience.'
        odt=self.root/'111.odt'
        with zipfile.ZipFile(odt,'w') as z:
            z.writestr('content.xml','<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><office:body><office:text><text:p>'+sentence+'</text:p></office:text></office:body></office:document-content>')
        self.assertIn(sentence,documents.extract_file(odt)['pages'][0]['text'])
        pptx=self.root/'112.pptx'
        with zipfile.ZipFile(pptx,'w') as z:
            z.writestr('ppt/slides/slide1.xml','<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:p><a:r><a:t>'+sentence+'</a:t></a:r></a:p></p:sld>')
        self.assertIn(sentence,documents.extract_file(pptx)['pages'][0]['text'])
        if shutil_which('soffice'):
            rtf=self.root/'113.rtf';rtf.write_text('{\\rtf1\\ansi '+sentence+'}',encoding='ascii')
            out=self.root/'doc_output';out.mkdir()
            subprocess.run(['soffice','-env:UserInstallation='+(self.root/'lo').as_uri(),'--headless','--convert-to','doc','--outdir',str(out),str(rtf)],check=True,capture_output=True,timeout=60)
            converted=out/'113.doc';self.assertTrue(converted.exists())
            self.assertIn('Synthetic',documents.extract_file(converted)['pages'][0]['text'])
        import unicodedata
        folder=self.root/'cvs'/unicodedata.normalize('NFD','Папка й');folder.mkdir()
        (folder/'101.txt').write_text(sentence,encoding='utf-8');self.assertEqual(service.scan()['new'],['101'])
        folder.rename(folder.with_name(unicodedata.normalize('NFC',folder.name)))
        self.assertEqual(service.scan()['changed'],[])

def shutil_which(name):
    import shutil
    return shutil.which(name)

if __name__=='__main__':unittest.main()
