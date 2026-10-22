import collections
import hashlib
import json
import re
import shutil
import sqlite3
import unicodedata
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

from openpyxl import load_workbook
from .common import DATA, active, connect, dump, layout, read, sha, stamp
from .documents import extract_file, broken
from .books import template, source_sheet, export
from .engine.bans import StructuredBans
from .engine.pilot150_bans import brief
from .engine.pilot150_evidence import attach
from .engine.profiles import assess, catalog

def portable_meta(run):
    return read(run/'portable.json')

def implementation_hash():
    folder=Path(__file__).parent/'engine'
    pairs=[(p.name,sha(p)) for p in sorted(folder.glob('*.py'))]
    return hashlib.sha256(json.dumps(pairs).encode()).hexdigest()

def rules_hash(inputs):
    pairs=[(unicodedata.normalize('NFC',str(p.relative_to(inputs)).replace('\\','/')),sha(p)) for p in inputs.rglob('*') if p.is_file() and p.suffix.lower() in ('.csv','.xlsx','.json')]
    return hashlib.sha256(json.dumps(sorted(pairs),ensure_ascii=False).encode()).hexdigest()

def prepare(run):
    meta=portable_meta(run)
    if 'implementation_hash' not in meta:
        meta['implementation_hash']=implementation_hash();dump(run/'portable.json',meta)
    c=connect(run/'state.sqlite')
    c.executescript('''
    CREATE TABLE IF NOT EXISTS portable_receipts(id TEXT, text_sha256 TEXT, start INTEGER, end INTEGER, PRIMARY KEY(id,text_sha256,start));
    CREATE TABLE IF NOT EXISTS portable_claims(id TEXT PRIMARY KEY,worker TEXT,created TEXT);
    CREATE TABLE IF NOT EXISTS portable_revisions(n INTEGER PRIMARY KEY,id TEXT,kind TEXT,payload TEXT,created TEXT);
    ''')
    c.close()

def source_rows(run):
    wb=load_workbook(template(run),read_only=True,data_only=False)
    result={str(v[1]):(n,v) for n,v in enumerate(wb[source_sheet(run)].iter_rows(min_row=2,values_only=True),2) if len(v)>1 and v[1] is not None}
    wb.close();return result

def new_run(previous=None):
    layout()
    ident='review_'+datetime.now(ZoneInfo('Europe/Moscow')).strftime('%Y%m%d_%H%M%S_%f')
    run=DATA/'state'/'runs'/ident;run.mkdir()
    if previous:
        shutil.copytree(previous/'inputs',run/'inputs')
        shutil.copytree(DATA/'sources',run/'inputs',dirs_exist_ok=True)
        pm=portable_meta(previous)
        meta=read(previous/'manifest.json')
        src=run/pm['source_book']
    else:
        shutil.copytree(DATA/'sources',run/'inputs')
        candidates=[p for p in (run/'inputs').glob('*.xlsx') if 'образец' not in p.name.lower()]
        if len(candidates)!=1:raise ValueError('В data/sources нужна одна исходная книга XLSX и каталоги clusters/banlist')
        src=candidates[0];meta={}
    wb=load_workbook(src,read_only=True)
    sheet=next((n for n in wb.sheetnames if 'матка' in n.lower()),wb.sheetnames[0]);wb.close()
    meta.update(as_of=datetime.now(ZoneInfo('Europe/Moscow')).date().isoformat(),sheet=sheet,version='portable-1')
    meta['input_hashes']={p.relative_to(run/'inputs').as_posix():sha(p) for p in (run/'inputs').rglob('*') if p.is_file()}
    dump(run/'manifest.json',meta)
    dump(run/'portable.json',{'source_book':src.relative_to(run).as_posix(),'sheet':sheet,'output_name':'Результаты проверки.xlsx','rules_hash':rules_hash(run/'inputs'),'implementation_hash':implementation_hash(),'created':stamp(),'previous':previous.name if previous else None})
    c=connect(run/'state.sqlite')
    c.executescript('''CREATE TABLE candidates(id TEXT PRIMARY KEY,source_row INTEGER NOT NULL,files TEXT NOT NULL,fingerprint TEXT NOT NULL,extraction_status TEXT NOT NULL DEFAULT 'pending',text TEXT,documents TEXT,error TEXT,review_status TEXT NOT NULL DEFAULT 'pending',result TEXT,reviewed_utc TEXT,exported_revision INTEGER);
    CREATE TABLE fact_cards(id TEXT PRIMARY KEY,card TEXT,text_sha256 TEXT,updated_utc TEXT);
    CREATE TABLE documents(hash TEXT PRIMARY KEY,path TEXT,kind TEXT,status TEXT,pages TEXT,error TEXT,updated_utc TEXT);''')
    if previous:
        old=sqlite3.connect((previous/'state.sqlite').resolve().as_uri()+'?mode=ro',uri=True);old.row_factory=sqlite3.Row
        for r in old.execute('SELECT * FROM candidates'):
            # Reuse full text only. Scores/bans must be reconsidered for new criteria.
            c.execute('INSERT INTO candidates(id,source_row,files,fingerprint,extraction_status,text,documents,error) VALUES(?,?,?,?,?,?,?,?)',tuple(r[k] for k in ('id','source_row','files','fingerprint','extraction_status','text','documents','error')))
        old.close()
    c.commit();c.close();prepare(run)
    dump(DATA/'state'/'active.json',{'run':ident})
    return run

def scan(new_version=False):
    layout()
    run=active() if (DATA/'state'/'active.json').exists() else new_run()
    prepare(run)
    meta=portable_meta(run)
    # Compare only sources provided by the operator. Missing files are not deletions.
    changed=[]
    for p in (DATA/'sources').rglob('*'):
        if p.is_file():
            target=run/'inputs'/p.relative_to(DATA/'sources')
            if not target.exists() or sha(p)!=sha(target):changed.append(str(p.relative_to(DATA/'sources')))
    if changed:
        if not new_version:raise ValueError('Изменены источники: '+', '.join(changed)+'. Запустите scan --new-version для новой проверки.')
        run=new_run(run)
    elif new_version:
        run=new_run(run)
    if rules_hash(run/'inputs')!=portable_meta(run)['rules_hash']:
        raise ValueError('Изменён снимок источников активного запуска; восстановите его и создайте новую версию через data/sources')
    if implementation_hash()!=portable_meta(run)['implementation_hash']:
        raise ValueError('Изменён движок правил: требуется scan --new-version')
    prepare(run);source=source_rows(run)
    groups=collections.defaultdict(list);ignored=[]
    for p in (DATA/'cvs').rglob('*'):
        if not p.is_file():continue
        ident=unicodedata.normalize('NFC',p.stem)
        if not ident.isdigit():ignored.append(p.name);continue
        groups[ident].append({'path':p.relative_to(DATA/'cvs').as_posix(),'sha256':sha(p),'size':p.stat().st_size})
    c=connect(run/'state.sqlite');new=[];changed_ids=[]
    for ident,files in groups.items():
        if ident not in source:continue
        files.sort(key=lambda f:unicodedata.normalize('NFC',f['path']))
        # Content identity is independent of absolute path or Unicode normalization.
        hashes=sorted(set(f['sha256'] for f in files));fp=hashlib.sha256(json.dumps(hashes).encode()).hexdigest()
        old=c.execute('SELECT * FROM candidates WHERE id=?',(ident,)).fetchone()
        old_hashes=sorted(set(f['sha256'] for f in json.loads(old['files']))) if old else []
        with c:
            if old and hashes==old_hashes:
                c.execute('UPDATE candidates SET files=?,fingerprint=? WHERE id=?',(json.dumps(files,ensure_ascii=False),fp,ident))
            elif old:
                c.execute('INSERT INTO portable_revisions(id,kind,payload,created) VALUES(?,?,?,?)',(ident,'cv_changed',json.dumps(dict(old),ensure_ascii=False),stamp()))
                c.execute("UPDATE candidates SET files=?,fingerprint=?,extraction_status='pending',review_status='pending',text=NULL,documents=NULL,error=NULL,result=NULL,exported_revision=NULL WHERE id=?",(json.dumps(files,ensure_ascii=False),fp,ident))
                c.execute('DELETE FROM portable_claims WHERE id=?',(ident,));changed_ids.append(ident)
            else:
                c.execute('INSERT INTO candidates(id,source_row,files,fingerprint) VALUES(?,?,?,?)',(ident,source[ident][0],json.dumps(files,ensure_ascii=False),fp));new.append(ident)
    c.close()
    result={'run':run.name,'new':new,'changed':changed_ids,'ignored':ignored,'ids_not_in_table':sorted(set(groups)-set(source))}
    dump(run/'last_scan.json',result);return result

def status(all_runs=False):
    layout()
    if not (DATA/'state'/'active.json').exists():return {'message':'Нет активного запуска. Восстановите архив или добавьте источники и выполните scan.'}
    runs=[p.parent for p in (DATA/'state'/'runs').glob('*/portable.json')] if all_runs else [active()]
    result=[]
    for run in runs:
        c=connect(run/'state.sqlite')
        counts=[dict(x) for x in c.execute('SELECT extraction_status,review_status,COUNT(*) count FROM candidates GROUP BY 1,2')]
        results=[json.loads(x[0]) for x in c.execute("SELECT result FROM candidates WHERE review_status='reviewed' AND result IS NOT NULL")]
        c.close()
        result.append({'run':run.name,'counts':counts,'processed':len(results),'banned':sum(x.get('ban')=='да' for x in results),'without_clusters':sum(x.get('ban')=='нет' and not x.get('clusters') for x in results)})
    return result

def locate(file):
    p=DATA/'cvs'/file['path']
    if p.is_file() and sha(p)==file['sha256']:return p
    # Legacy Windows paths and reorganized incoming folders are rebound by hash.
    for p in (DATA/'cvs').rglob('*'):
        if p.is_file() and p.stat().st_size==file.get('size',p.stat().st_size) and sha(p)==file['sha256']:return p
    raise FileNotFoundError('Оригинал CV отсутствует; добавьте его в data/cvs: '+file['sha256'])

def extract(retry=False,ids=None):
    run=active();prepare(run);c=connect(run/'state.sqlite');done=[]
    rows=c.execute("SELECT * FROM candidates WHERE review_status!='reviewed' ORDER BY source_row").fetchall()
    for row in rows:
        if ids and row['id'] not in ids:continue
        if row['extraction_status']=='ready' and not retry:continue
        if row['extraction_status']=='error' and not retry:continue
        try:
            docs=[];used=set()
            for f in json.loads(row['files']):
                if f['sha256'] in used:continue
                used.add(f['sha256']);p=locate(f);doc=extract_file(p,use_cache=not retry)
                doc['file']=p.relative_to(DATA/'cvs').as_posix();docs.append(doc)
            text='\n\n'.join('FILE '+d['file']+'\n'+'\n'.join('PAGE '+str(p['page'])+'\n'+p['text'] for p in d['pages']) for d in docs)
            with c:c.execute("UPDATE candidates SET text=?,documents=?,extraction_status='ready',error=NULL WHERE id=?",(text,json.dumps(docs,ensure_ascii=False),row['id']))
            done.append({'id':row['id'],'status':'ready'})
        except Exception as e:
            with c:c.execute("UPDATE candidates SET extraction_status='error',error=? WHERE id=?",(str(e),row['id']))
            done.append({'id':row['id'],'status':'error','error':str(e)})
        print(json.dumps(done[-1],ensure_ascii=False),flush=True)
    c.close();return {'attempted':len(done)}

def claim(worker,count):
    run=active();prepare(run);c=connect(run/'state.sqlite')
    with c:
        c.execute('BEGIN IMMEDIATE')
        owned=[r[0] for r in c.execute("SELECT c.id FROM candidates c JOIN portable_claims p ON p.id=c.id WHERE p.worker=? AND c.extraction_status='ready' AND c.review_status!='reviewed' ORDER BY c.source_row LIMIT ?",(worker,count))]
        ids=owned or [r[0] for r in c.execute("SELECT c.id FROM candidates c LEFT JOIN portable_claims p ON p.id=c.id WHERE p.id IS NULL AND c.extraction_status='ready' AND c.review_status!='reviewed' ORDER BY c.source_row LIMIT ?",(count,))]
        for ident in ids:c.execute('INSERT OR IGNORE INTO portable_claims VALUES(?,?,?)',(ident,worker,stamp()))
    c.close();return ids

def show(ident,offset=0,limit=10000):
    if offset<0 or not 1<=limit<=20000:raise ValueError('offset >=0; limit от 1 до 20000')
    run=active();prepare(run);c=connect(run/'state.sqlite');r=c.execute('SELECT text FROM candidates WHERE id=?',(ident,)).fetchone()
    if not r:raise ValueError('Неизвестный ID')
    text=r['text'] or '';h=hashlib.sha256(text.encode()).hexdigest();end=min(len(text),offset+limit)
    if offset>len(text):raise ValueError('offset за пределами текста')
    with c:c.execute('INSERT OR REPLACE INTO portable_receipts VALUES(?,?,?,?)',(ident,h,offset,end))
    c.close();return {'id':ident,'text_sha256':h,'offset':offset,'end':end,'total_characters':len(text),'next_offset':end if end<len(text) else None,'text':text[offset:end]}

def russian(value,label):
    if not value:return
    if broken(value) or re.search(r'\?{2,}',value):raise ValueError(label+': повреждённые символы')
    if not re.search('[А-Яа-яЁё]',value):raise ValueError(label+': требуется русский текст')

def preview(card,worker=None):
    run=active();prepare(run);c=connect(run/'state.sqlite');ident=str(card['id'])
    r=c.execute('SELECT * FROM candidates WHERE id=?',(ident,)).fetchone()
    intervals=c.execute('SELECT start,end FROM portable_receipts WHERE id=? AND text_sha256=? ORDER BY start',(ident,card.get('reviewed_text_sha256',''))).fetchall()
    owner=c.execute('SELECT worker FROM portable_claims WHERE id=?',(ident,)).fetchone()
    c.close()
    if not r or r['extraction_status']!='ready':raise ValueError('Текст не готов')
    if not card.get('full_text_review_complete'):raise ValueError('Не подтверждено полное чтение')
    if rules_hash(run/'inputs')!=portable_meta(run)['rules_hash']:raise ValueError('Источники версии изменены; требуется новая версия проверки')
    if implementation_hash()!=portable_meta(run)['implementation_hash']:raise ValueError('Движок правил изменён; требуется scan --new-version')
    h=hashlib.sha256(r['text'].encode()).hexdigest()
    if card.get('reviewed_text_sha256')!=h:raise ValueError('Укажите актуальный reviewed_text_sha256 из show')
    end=0
    for a,b in intervals:
        if a>end:break
        end=max(end,b)
    if end<len(r['text']):raise ValueError('Прочитан не весь текст через show; следующий offset='+str(end))
    if worker:
        if not owner or owner[0]!=worker:raise ValueError('ID не закреплен за worker')
    if any(len(j)<6 for j in card.get('jobs',[])):raise ValueError('Работа: работодатель, начало, конец, должность/задачи, регион, тип организации')
    russian(card.get('comment',''),'Комментарий')
    for review in card.get('reviews',[]):
        russian(review[3],'Основание оценки')
        if len(review)>4:russian(review[4],'Пробелы')
    ban=attach(StructuredBans(run).assess(card,source_rows(run)[ident][1]),card,r['text'])
    if any(not f.get('source_excerpt') for f in ban['active']):raise ValueError('Для бана требуется цитата; добавьте ban_quote_anchors')
    assessed,rejected,chosen=assess(card,catalog(run)) if ban['ban']=='нет' else ([],[],[])
    clusters=[{'cluster':x['cluster'],'score':x['score']} for x in chosen]
    comment=card.get('comment','') if ban['ban']=='нет' else ''
    if not comment and chosen:comment='; '.join(dict.fromkeys(x['missing'] for x in chosen if x['missing']))
    result={'id':ident,'ai':'да','ban':ban['ban'],'ban_comment':brief(ban),'department':'Бан' if ban['ban']=='да' else '; '.join(x['cluster']+' '+str(x['score']) for x in clusters) if clusters else 'Подходящих кластеров не выявлено','comment':comment,'clusters':clusters,'ban_assessment':ban,'profile_assessments':assessed,'excluded_profiles':rejected,'fact_card':card,'text_sha256':h,'rules_hash':portable_meta(run)['rules_hash'],'review_method':'portable-semantic-v1'}
    return result

def submit(card,worker):
    run=active();result=preview(card,worker);c=connect(run/'state.sqlite');ident=str(card['id'])
    prior=c.execute('SELECT review_status,result FROM candidates WHERE id=?',(ident,)).fetchone()
    if prior['review_status']=='reviewed':
        identical=json.loads(prior['result']).get('fact_card')==card
        c.close()
        if identical:return {'id':ident,'already_completed':True}
        raise ValueError('Завершённая проверка: создайте новую версию для пересмотра')
    with c:
        c.execute('BEGIN IMMEDIATE')
        old=c.execute('SELECT * FROM candidates WHERE id=?',(ident,)).fetchone()
        if old['extraction_status']!='ready' or hashlib.sha256((old['text'] or '').encode()).hexdigest()!=result['text_sha256']:
            raise ValueError('Документ изменился между preview и сохранением; перечитайте новую версию')
        if old['review_status']=='reviewed':
            raise ValueError('Завершённая проверка: создайте новую версию для пересмотра')
        c.execute('INSERT INTO portable_revisions(id,kind,payload,created) VALUES(?,?,?,?)',(ident,'completed',json.dumps(result,ensure_ascii=False),stamp()))
        c.execute('INSERT OR REPLACE INTO fact_cards VALUES(?,?,?,?)',(ident,json.dumps(card,ensure_ascii=False),result['text_sha256'],stamp()))
        c.execute("UPDATE candidates SET review_status='reviewed',result=?,reviewed_utc=?,exported_revision=NULL WHERE id=?",(json.dumps(result,ensure_ascii=False),stamp(),ident))
    count=c.execute("SELECT COUNT(*) FROM candidates WHERE review_status='reviewed' AND exported_revision IS NULL").fetchone()[0]
    c.execute('CREATE TABLE IF NOT EXISTS portable_exports(created TEXT,path TEXT,sha256 TEXT,count INTEGER)')
    latest=c.execute('SELECT created FROM portable_exports ORDER BY rowid DESC LIMIT 1').fetchone();c.close()
    elapsed=(datetime.now().astimezone()-datetime.fromisoformat(latest[0])).total_seconds() if latest else 601
    if count>=25 or elapsed>=600:
        result['export']=export(run)
        c=connect(run/'state.sqlite')
        with c:c.execute("UPDATE candidates SET exported_revision=1 WHERE review_status='reviewed'")
        c.close()
    return result
