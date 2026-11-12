"""Archive-scoped fresh reviews. Only the parent process writes SQLite."""
import hashlib
import json
import shutil
import stat
import tempfile
import zipfile
from datetime import date, datetime
from pathlib import Path, PurePosixPath
from zoneinfo import ZoneInfo

from openpyxl import load_workbook
from . import service
from .common import DATA, active, connect, dump, layout, read, sha, stamp


def _registry_files(path, as_of):
    """Copy only validated package members, never neighbouring local files."""
    from .amendments import validate_registry, snapshot_path
    path=Path(path)
    registry=validate_registry(path,as_of)
    members={'registry.json':path}
    proofs=list(registry.get('supplementary_sources',[]))
    for entity in registry['entities']:proofs.extend(entity['evidence'])
    for proof in proofs:
        for key in ('snapshot','raw_snapshot'):
            if proof.get(key):members[proof[key]]=snapshot_path(path.parent,proof[key])
    return members


def _criteria_key(run, registry_files=None, criteria_dir=None):
    # Source workbooks are retained as provenance, but do not define professions.
    root=Path(criteria_dir) if criteria_dir is not None else run/'inputs'
    pairs={p.relative_to(root).as_posix():sha(p)
           for p in root.rglob('*') if p.is_file()
           and not (p.parent==root and p.suffix.lower()=='.xlsx')
           and p.name!='accepted_rules.md'}
    if registry_files is not None:
        prefix='amendments/rule262/'
        pairs={name:value for name,value in pairs.items() if not name.startswith(prefix)}
        pairs.update({prefix+name:sha(path) for name,path in registry_files.items()})
    policy=Path(__file__).parent.parent/'rules/accepted_rules.md'
    pairs['accepted_rules.md']=sha(policy)
    return hashlib.sha256(json.dumps(sorted(pairs.items()),ensure_ascii=False).encode()).hexdigest()


def import_archive(file, template, output=None, as_of=None, partner_registry=None,
                   generated_from_resumes=False, criteria_dir=None):
    """Validate all inputs first, then atomically publish a new independent run."""
    archive=Path(file);book=Path(template)
    if not archive.is_file() or not book.is_file():raise ValueError('Архив и шаблон должны существовать')
    if as_of is None:as_of=datetime.now(ZoneInfo('Europe/Moscow')).date().isoformat()
    try:
        if not isinstance(as_of,str) or date.fromisoformat(as_of).isoformat()!=as_of:raise ValueError()
    except (TypeError,ValueError) as error:
        raise ValueError('Дата проверки должна иметь формат YYYY-MM-DD') from error
    layout();previous=active()
    criteria_root=Path(criteria_dir) if criteria_dir is not None else previous/'inputs'
    if criteria_dir is not None:
        if not criteria_root.is_dir():raise ValueError('Каталог критериев не существует')
        if any(p.is_symlink() for p in criteria_root.rglob('*')):
            raise ValueError('Ссылки в каталоге критериев недопустимы')
        if len(list((criteria_root/'clusters').glob('*.xlsx')))!=1:
            raise ValueError('Нужна одна книга критериев в clusters/*.xlsx')
        if not list((criteria_root/'banlist').glob('*.csv')):
            raise ValueError('Нужны полные бан-листы в banlist/*.csv')
        # Validate the same catalogue and CSV reader used for actual reviews.
        from .engine.profiles import catalog
        from .engine.bans import StructuredBans
        with tempfile.TemporaryDirectory() as temporary:
            check=Path(temporary);shutil.copytree(criteria_root,check/'inputs')
            dump(check/'manifest.json',{'as_of':as_of})
            catalog(check);StructuredBans(check)
    registry_files=_registry_files(partner_registry,as_of) if partner_registry else None
    if registry_files is None:
        from .amendments import REGISTRY_PATH, validate_registry
        existing=criteria_root/REGISTRY_PATH
        if existing.exists():validate_registry(existing,as_of)
    criteria=_criteria_key(previous,registry_files,criteria_root)
    key_data={'archive_sha256':sha(archive),'template_sha256':sha(book),
              'criteria_sha256':criteria,'implementation_hash':service.implementation_hash(),
              'assessment_date':as_of}
    if generated_from_resumes:key_data['generated_from_resumes']=True
    key=hashlib.sha256(json.dumps(key_data,sort_keys=True).encode()).hexdigest()
    legacy_key=hashlib.sha256(json.dumps({k:v for k,v in key_data.items() if k!='assessment_date'},sort_keys=True).encode()).hexdigest()
    for p in sorted((DATA/'state/runs').glob('*/portable.json')):
        meta=read(p)
        if meta.get('batch_key') in (key,legacy_key):
            run=p.parent
            if not meta.get('import_complete'):continue
            if read(run/'manifest.json')['as_of']!=as_of:continue
            if service.rules_hash(run/'inputs')!=meta['rules_hash']:raise ValueError('Изменён снимок импортированного запуска')
            # Verify the immutable archive inventory, even when resuming.
            manifest=read(run/'archive_manifest.json')
            for rec in manifest['files']:
                if sha(DATA/'cvs'/rec['path'])!=rec['sha256']:raise ValueError('Изменён оригинал: '+rec['path'])
            c=connect(run/'state.sqlite')
            try:
                if {r[0] for r in c.execute('SELECT id FROM candidates')}!=set(manifest['ids']):
                    raise ValueError('Неполный импорт: состав базы не совпадает с манифестом')
            finally:c.close()
            dump(DATA/'state/active.json',{'run':run.name})
            return {'run':run.name,'candidates':len(manifest['ids']),'files':len(manifest['files']),'resumed':True}
    wb=load_workbook(book,read_only=True,data_only=False)
    try:
        ws=wb.worksheets[0];sheet=ws.title
        if ws.max_column!=38 or str(ws.cell(1,1).value).strip()!='ID':raise ValueError('Ожидается шаблон с ID в A и 38 столбцами')
        source={}
        for n,values in enumerate(ws.iter_rows(min_row=2,values_only=True),2):
            if values[0] is None or not str(values[0]).strip():continue
            ident=str(values[0]).strip()
            if ident in source:raise ValueError('Повторный ID в исходной таблице: '+ident)
            source[ident]=n
    finally:wb.close()
    groups={};entries=[];names=set()
    with zipfile.ZipFile(archive) as z:
        bad=z.testzip()
        if bad:raise ValueError('Ошибка CRC: '+bad)
        for info in z.infolist():
            rel=PurePosixPath(info.filename)
            if (rel.is_absolute() or '..' in rel.parts or '\\' in info.filename
                or info.flag_bits&1 or stat.S_ISLNK(info.external_attr>>16)):
                raise ValueError('Недопустимый путь архива: '+info.filename)
            if info.is_dir():continue
            canonical=rel.as_posix()
            if canonical in names:raise ValueError('Повторный путь архива: '+info.filename)
            names.add(canonical)
            ident=Path(rel.name).stem
            if not ident.isdigit():raise ValueError('Имя файла должно содержать числовой ID: '+info.filename)
            if ident not in source:raise ValueError('ID отсутствует в шаблоне: '+ident)
            entries.append((info.filename,ident));groups.setdefault(ident,[])
        if not entries:raise ValueError('Пустой архив резюме')
        run_name='archive_'+datetime.now(ZoneInfo('Europe/Moscow')).strftime('%Y%m%d_%H%M%S_%f')
        run=DATA/'state/runs'/run_name
        originals=DATA/'cvs'/run_name
        target=Path(output) if output else DATA/'exports'/('newCVs_'+as_of+'.xlsx')
        if target.resolve() in (book.resolve(),archive.resolve()):raise ValueError('Нельзя перезаписывать исходник')
        if target.exists():target=target.with_name(target.stem+'_'+run_name+target.suffix)
        run.mkdir();originals.mkdir()
        try:
            shutil.copytree(criteria_root,run/'inputs')
            # Row templates are provenance, not part of an explicit criteria set.
            if criteria_dir is not None:
                for old_book in (run/'inputs').glob('*.xlsx'):old_book.unlink()
            if registry_files is not None:
                destination=run/'inputs/amendments/rule262'
                if destination.exists():shutil.rmtree(destination)
                for name,registry_source in registry_files.items():
                    dest=destination/name;dest.parent.mkdir(parents=True,exist_ok=True)
                    shutil.copy2(registry_source,dest)
            source_book=run/'inputs/batch_template.xlsx';shutil.copy2(book,source_book)
            shutil.copy2(Path(__file__).parent.parent/'rules/accepted_rules.md',run/'inputs/accepted_rules.md')
            inventory=[]
            for name,ident in entries:
                dest=originals.joinpath(*PurePosixPath(name).parts);dest.parent.mkdir(parents=True,exist_ok=True)
                with z.open(name) as src,dest.open('wb') as dst:shutil.copyfileobj(src,dst)
                rec={'id':ident,'path':dest.relative_to(DATA/'cvs').as_posix(),'sha256':sha(dest),'size':dest.stat().st_size}
                groups[ident].append(rec);inventory.append(rec)
            meta={'source_book':'inputs/batch_template.xlsx','sheet':sheet,'id_column':1,
                  'source_column_count':38,'export_layout':'append_results','output_name':target.name,
                  'output_path':str(target),'rules_hash':service.rules_hash(run/'inputs'),
                  'implementation_hash':service.implementation_hash(),'created':stamp(),
                  'previous':previous.name,'batch_key':key,**key_data}
            dump(run/'portable.json',meta)
            dump(run/'manifest.json',{'as_of':as_of,'sheet':sheet,'version':'archive-1',
                  'input_hashes':{p.relative_to(run/'inputs').as_posix():sha(p) for p in (run/'inputs').rglob('*') if p.is_file()}})
            dump(run/'archive_manifest.json',{'archive':archive.name,**key_data,'files':inventory,
                 'ids':sorted(groups,key=lambda i:source[i]),'fresh_review':True})
            c=connect(run/'state.sqlite')
            try:
                c.executescript('''CREATE TABLE candidates(id TEXT PRIMARY KEY,source_row INTEGER NOT NULL,files TEXT NOT NULL,fingerprint TEXT NOT NULL,extraction_status TEXT NOT NULL DEFAULT 'pending',text TEXT,documents TEXT,error TEXT,review_status TEXT NOT NULL DEFAULT 'pending',result TEXT,reviewed_utc TEXT,exported_revision INTEGER);
                CREATE TABLE fact_cards(id TEXT PRIMARY KEY,card TEXT,text_sha256 TEXT,updated_utc TEXT);
                CREATE TABLE documents(hash TEXT PRIMARY KEY,path TEXT,kind TEXT,status TEXT,pages TEXT,error TEXT,updated_utc TEXT);''')
                with c:
                    for ident,files in groups.items():
                        files.sort(key=lambda f:f['path'])
                        fp=hashlib.sha256(json.dumps(sorted(set(f['sha256'] for f in files))).encode()).hexdigest()
                        c.execute('INSERT INTO candidates(id,source_row,files,fingerprint) VALUES(?,?,?,?)',
                                  (ident,source[ident],json.dumps(files,ensure_ascii=False),fp))
            finally:c.close()
            service.prepare(run)
            meta['import_complete']=True;dump(run/'portable.json',meta)
        except BaseException:
            shutil.rmtree(run,ignore_errors=True);shutil.rmtree(originals,ignore_errors=True);raise
    dump(DATA/'state/active.json',{'run':run_name})
    return {'run':run_name,'candidates':len(groups),'files':len(entries),'resumed':False,'output':str(target)}


def extract_parallel(workers=2,retry=False,ids=None):
    if not 1<=workers<=2:raise ValueError('Допустимо не более двух процессов извлечения')
    run=active();service.prepare(run);c=connect(run/'state.sqlite')
    try:
        rows=[dict(r) for r in c.execute("SELECT * FROM candidates WHERE review_status!='reviewed' ORDER BY source_row")
              if (not ids or r['id'] in ids) and (retry or r['extraction_status']=='pending')]
    finally:c.close()
    unique={f['sha256']:f for row in rows for f in json.loads(row['files'])}
    from .parallel_extract import iter_extracted
    results={};remaining={r['id']:{f['sha256'] for f in json.loads(r['files'])} for r in rows};done=[]
    paths={str(DATA/'cvs'/f['path']):h for h,f in unique.items()}
    for path,h in paths.items():
        if sha(path)!=h:raise ValueError('Изменён хеш оригинала: '+path)
    # No SQLite connection is inherited or held open during worker startup.
    iterator=iter_extracted(paths,max_workers=workers,use_cache=not retry)
    for path,doc,error in iterator:
        if doc and (doc.get('sha256')!=paths[path] or doc.get('complete') is not True
                    or not doc.get('pages') or sha(path)!=paths[path]):
            doc,error=None,'Изменён оригинал или не подтверждена полнота извлечения'
        h=paths[path];results[h]=(doc,error)
        ready=[row for row in rows if row['id'] in remaining and remaining[row['id']]<=results.keys()]
        if not ready:continue
        c=connect(run/'state.sqlite')
        try:
            for row in ready:
                docs=[];errors=[];used=set()
                try:service.require_complete_candidate_source(run,row)
                except ValueError as exc:errors.append(str(exc))
                for f in json.loads(row['files']):
                    if f['sha256'] in used:continue
                    used.add(f['sha256']);doc,error=results[f['sha256']]
                    if error:errors.append(f['path']+': '+error)
                    else:
                        doc=dict(doc);doc['file']=f['path'];docs.append(doc)
                if errors:
                    with c:c.execute("UPDATE candidates SET extraction_status='error',error=? WHERE id=?",('; '.join(errors),row['id']))
                    item={'id':row['id'],'status':'error','error':'; '.join(errors)}
                else:
                    text='\n\n'.join('FILE '+d['file']+'\n'+'\n'.join('PAGE '+str(p['page'])+'\n'+p['text'] for p in d['pages']) for d in docs)
                    with c:c.execute("UPDATE candidates SET text=?,documents=?,extraction_status='ready',error=NULL WHERE id=?",(text,json.dumps(docs,ensure_ascii=False),row['id']))
                    item={'id':row['id'],'status':'ready'}
                del remaining[row['id']];done.append(item);print(json.dumps(item,ensure_ascii=False),flush=True)
        finally:c.close()
    return {'attempted':len(done),'ready':sum(x['status']=='ready' for x in done),'errors':sum(x['status']=='error' for x in done)}


def amend_archive(parent_name, registry_file, engine_snapshot, output=None):
    """Fork an exact archive scope. Never alter or relabel the parent results.

    Run only inside Docker, with the coordinator's processing work stopped.
    Every completed parent assessment becomes pending in the child; its raw card
    and genuine reading receipts survive with explicit provenance for preview.
    """
    import sqlite3
    from .amendments import REGISTRY_PATH, canonical_hash, validate_registry, snapshot_path
    if Path(parent_name).name != parent_name:
        raise ValueError('Нужно имя родительского запуска')
    parent = DATA/'state/runs'/parent_name
    pm = read(parent/'portable.json'); manifest = read(parent/'manifest.json')
    inventory = read(parent/'archive_manifest.json')
    if not pm.get('batch_key') or not pm.get('import_complete'):
        raise ValueError('Нужен завершенный импорт архивного запуска')
    if service.rules_hash(parent/'inputs') != pm['rules_hash']:
        raise ValueError('Изменён снимок родительских правил')
    engine_snapshot = Path(engine_snapshot)
    if not engine_snapshot.is_dir() or not list(engine_snapshot.glob('*.py')):
        raise ValueError('Нужен полный снимок родительского движка')
    old_hashes = [service.implementation_hash_from(engine_snapshot, ending) for ending in ('lf', 'crlf')]
    if pm['implementation_hash'] not in old_hashes:
        raise ValueError('Снимок родительского движка не соответствует сохраненному хешу')
    registry_file = Path(registry_file)
    registry = validate_registry(registry_file, manifest['as_of'])
    registry_sha = sha(registry_file)
    current_engine = service.implementation_hash()
    # Include the entire logical parent DB, not wall-clock time or file layout,
    # so the same request resumes after a successful atomic publication.
    source = sqlite3.connect((parent/'state.sqlite').resolve().as_uri()+'?mode=ro', uri=True)
    source.row_factory = sqlite3.Row
    try:
        source.execute('BEGIN')
        logical = hashlib.sha256('\n'.join(source.iterdump()).encode()).hexdigest()
        key = canonical_hash({'parent':parent_name,'parent_database':logical,
                              'registry':registry_sha,'engine':current_engine,
                              'rules':pm['rules_hash']})
        for candidate in sorted((DATA/'state/runs').glob('*/portable.json')):
            cm = read(candidate)
            if cm.get('amendment_key') == key and cm.get('import_complete'):
                child = candidate.parent
                if (service.rules_hash(child/'inputs') != cm['rules_hash'] or
                        not service.implementation_matches(cm['implementation_hash'])):
                    raise ValueError('Изменена созданная amendment-версия')
                from .amendments import verify_amendment_run
                verify_amendment_run(child)
                verify_connection=connect(child/'state.sqlite')
                try:
                    if {r[0] for r in verify_connection.execute('SELECT id FROM candidates')}!=set(inventory['ids']):
                        raise ValueError('Изменён состав amendment-версии')
                finally:verify_connection.close()
                validate_registry(child/'inputs'/REGISTRY_PATH, manifest['as_of'])
                for rec in inventory['files']:
                    if sha(DATA/'cvs'/rec['path']) != rec['sha256']:
                        raise ValueError('Изменён оригинал amendment')
                dump(DATA/'state/active.json', {'run':child.name})
                return {'run':child.name,'parent':parent_name,'resumed':True,
                        'output':cm['output_path']}
        rows = [dict(r) for r in source.execute('SELECT * FROM candidates ORDER BY source_row')]
        expected = set(inventory['ids'])
        if {r['id'] for r in rows} != expected or len(expected) != len(inventory['ids']):
            raise ValueError('Состав родителя не совпадает с архивным манифестом')
        expected_files = {}
        for rec in inventory['files']:
            expected_files.setdefault(rec['id'], []).append(rec)
            if sha(DATA/'cvs'/rec['path']) != rec['sha256']:
                raise ValueError('Изменён оригинал: '+rec['path'])
        for row in rows:
            if sorted(json.loads(row['files']),key=lambda f:f['path']) != sorted(expected_files[row['id']],key=lambda f:f['path']):
                raise ValueError('Набор файлов отличается от архивного манифеста')
        name = 'amendment_'+datetime.now(ZoneInfo('Europe/Moscow')).strftime('%Y%m%d_%H%M%S_%f')
        child = DATA/'state/runs'/name
        target = Path(output) if output else DATA/'exports'/(Path(pm['output_name']).stem+'_'+name+'.xlsx')
        if target.suffix.lower() != '.xlsx':raise ValueError('Для amendment требуется отдельный Excel .xlsx')
        if target.exists():target=target.with_name(target.stem+'_'+name+target.suffix)
        protected = [parent/pm['source_book'],DATA/'cvs',parent]
        if any(target.resolve() == x.resolve() or (x.is_dir() and target.resolve().is_relative_to(x.resolve())) for x in protected):
            raise ValueError('Нельзя перезаписывать исходники/родительский запуск')
        child.mkdir()
        try:
            shutil.copytree(parent/'inputs',child/'inputs')
            dest = child/'inputs'/REGISTRY_PATH
            dest.parent.mkdir(parents=True,exist_ok=True)
            shutil.copy2(registry_file,dest)
            for entity in registry['entities']:
                for proof in entity['evidence']:
                    rel = Path(proof['snapshot']); out=dest.parent/rel
                    out.parent.mkdir(parents=True,exist_ok=True)
                    shutil.copy2(snapshot_path(registry_file.parent,rel),out)
                    if proof.get('raw_snapshot'):
                        raw_rel=Path(proof['raw_snapshot']);raw_out=dest.parent/raw_rel
                        raw_out.parent.mkdir(parents=True,exist_ok=True)
                        shutil.copy2(snapshot_path(registry_file.parent,raw_rel),raw_out)
            for proof in registry.get('supplementary_sources', []):
                rel=Path(proof['snapshot']);out=dest.parent/rel
                out.parent.mkdir(parents=True,exist_ok=True)
                shutil.copy2(snapshot_path(registry_file.parent,rel),out)
            validate_registry(dest,manifest['as_of'])
            shutil.copy2(parent/'archive_manifest.json',child/'archive_manifest.json')
            for sub in ('candidate_source_blocks',):
                if (parent/sub).exists():shutil.copytree(parent/sub,child/sub)
            audit=child/'audit';audit.mkdir()
            shutil.copy2(parent/'portable.json',audit/'parent_portable.json')
            shutil.copy2(parent/'manifest.json',audit/'parent_manifest.json')
            shutil.copytree(engine_snapshot,audit/'parent_engine')
            # BEGIN already pins the parent snapshot. Backup via an independent
            # read connection: sqlite backup must not run inside its write txn.
            backup_source=sqlite3.connect((parent/'state.sqlite').resolve().as_uri()+'?mode=ro',uri=True)
            snapshot=sqlite3.connect(audit/'parent_state.sqlite')
            try:
                backup_source.backup(snapshot)
                if hashlib.sha256('\n'.join(snapshot.iterdump()).encode()).hexdigest()!=logical:
                    raise ValueError('Родитель изменился во время создания amendment; остановите обработку')
            finally:snapshot.close();backup_source.close()
            shutil.copy2(audit/'parent_state.sqlite',child/'state.sqlite')
            cm=dict(pm)
            cm.update(previous=parent_name,parent_run=parent_name,created=stamp(),
                      rules_hash=service.rules_hash(child/'inputs'),implementation_hash=current_engine,
                      amendment_key=key,batch_key=canonical_hash({'parent_batch':pm['batch_key'],'amendment':key}),
                      output_name=target.name,output_path=str(target),import_complete=False)
            dump(child/'portable.json',cm)
            nm=dict(manifest);nm.update(version='archive-amendment-1',parent_run=parent_name,
                input_hashes={p.relative_to(child/'inputs').as_posix():sha(p) for p in (child/'inputs').rglob('*') if p.is_file()})
            dump(child/'manifest.json',nm)
            service.prepare(child)
            c=connect(child/'state.sqlite')
            provenance=[];cards=0
            try:
                with c:
                    c.execute('DELETE FROM portable_revisions')
                    c.execute('DELETE FROM portable_receipts')
                    c.execute('DELETE FROM fact_cards')
                    if c.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name='portable_exports'").fetchone():c.execute('DELETE FROM portable_exports')
                    c.execute("UPDATE candidates SET review_status='pending',result=NULL,reviewed_utc=NULL,exported_revision=NULL")
                    for row in rows:
                        ident=row['id'];text=row['text'] or '';th=hashlib.sha256(text.encode()).hexdigest()
                        intervals=[list(r) for r in source.execute('SELECT start,end FROM portable_receipts WHERE id=? AND text_sha256=? ORDER BY start',(ident,th))]
                        valid=[];end=0
                        for start,stop in intervals:
                            if start<0 or stop<start or stop>len(text):continue
                            valid.append([start,stop])
                            if start<=end:end=max(end,stop)
                        result=json.loads(row['result']) if row['result'] else None
                        fact=result.get('fact_card') if result else None
                        # Preserve actual reads for the exact text; no synthetic
                        # 0..N interval is created, even for incomplete coverage.
                        for start,stop in valid:c.execute('INSERT INTO portable_receipts VALUES(?,?,?,?)',(ident,th,start,stop))
                        record={'id':ident,'fingerprint':row['fingerprint'],'text_sha256':th,
                                'parent_review_status':row['review_status'],'read_receipts':valid,
                                'full_receipt_coverage':bool(text) and end==len(text),
                                'parent_result_sha256':canonical_hash(result) if result else None}
                        if row['review_status']=='reviewed':
                            if (not fact or fact.get('reviewed_text_sha256')!=th or not fact.get('full_text_review_complete') or not record['full_receipt_coverage']):
                                raise ValueError('Не подтверждены исходные факты/полное чтение: '+ident)
                            record['parent_card_sha256']=canonical_hash(fact)
                            original_folder=audit/'parent_fact_cards';original_folder.mkdir(exist_ok=True)
                            dump(original_folder/(ident+'.json'),fact)
                            from .amendments import prepare_relation_evidence
                            migrated, relation_matches = prepare_relation_evidence(fact, text, registry)
                            record['relation_candidates'] = relation_matches
                            migrated['amendment_provenance']={'parent_run':parent_name,'parent_card_sha256':record['parent_card_sha256'],
                                'parent_result_sha256':record['parent_result_sha256'],'amendment_key':key}
                            folder=child/'migrated_cards';folder.mkdir(exist_ok=True)
                            (folder/(ident+'.jsonl')).write_text(json.dumps(migrated,ensure_ascii=False)+'\n',encoding='utf-8')
                            c.execute('INSERT OR REPLACE INTO portable_claims VALUES(?,?,?)',(ident,'amendment',stamp()))
                            cards+=1
                        provenance.append(record)
                    receipt={'schema_version':1,'parent_run':parent_name,'amendment_key':key,'created':stamp(),
                        'parent_database_logical_sha256':logical,'parent_database_snapshot_sha256':sha(audit/'parent_state.sqlite'),
                        'parent_portable_sha256':sha(audit/'parent_portable.json'),'parent_manifest_sha256':sha(audit/'parent_manifest.json'),
                        'old_rules_hash':pm['rules_hash'],'new_rules_hash':cm['rules_hash'],
                        'old_implementation_hash':pm['implementation_hash'],'new_implementation_hash':current_engine,
                        'parent_engine_files':{p.name:sha(p) for p in sorted((audit/'parent_engine').glob('*.py'))},
                        'new_engine_files':{p.name:sha(p) for p in sorted((Path(service.__file__).parent/'engine').glob('*.py'))},
                        'registry_sha256':registry_sha,'amendment':registry['amendment'],
                        'archive_manifest_sha256':sha(child/'archive_manifest.json'),
                        'input_hashes':nm['input_hashes'],'candidates':provenance,
                        'scores_reused_as_results':False,'raw_fact_cards_retained':cards,
                        'parent_export_checksums':[{ 'path':str(p),'sha256':sha(p)} for p in [Path(pm.get('output_path',''))] if p.is_file()]}
                    dump(child/'amendment_receipt.json',receipt)
                    c.execute('INSERT INTO portable_revisions(id,kind,payload,created) VALUES(?,?,?,?)',('', 'amendment_created',json.dumps({'key':key,'parent':parent_name},ensure_ascii=False),stamp()))
            finally:c.close()
            cm['import_complete']=True;cm['amendment_receipt_sha256']=sha(child/'amendment_receipt.json')
            dump(child/'portable.json',cm)
        except BaseException:
            shutil.rmtree(child,ignore_errors=True)
            raise
    finally:source.close()
    dump(DATA/'state/active.json',{'run':name})
    return {'run':name,'parent':parent_name,'candidates':len(rows),'migrated_cards':cards,
            'resumed':False,'output':str(target),'worker':'amendment',
            'note':'Перенесенные факты требуют preview/submit новой версии; прежние результаты не изменены'}
