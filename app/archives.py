"""Verified archives, including consistent SQLite snapshots; no CV binaries."""
import hashlib
import json
import os
import shutil
import sqlite3
import tarfile
import tempfile
from collections import Counter
from pathlib import Path
from .common import DATA, dump, layout, read, sha, stamp

EXCLUDED_DIRS={'cvs','converted','ocr_images','audio_env','audio_models','.venv','venv','__pycache__','.git','node_modules','hr-hive','transfer','migration_staging'}
EXCLUDED_SUFFIXES={'.png','.jpg','.jpeg','.gif','.webp','.bmp','.tif','.tiff','.pdf','.doc','.docx','.rtf','.pages','.odt','.aac','.pptx','.dotx','.jfif','.crdownload','.zip','.pyc'}
CV_BINARY_SUFFIXES=EXCLUDED_SUFFIXES|{'.heic','.emf','.wmf','.svg','.ico','.bin','.xls','.xlsm','.ods'}


def evidence_inventory(source):
    """Explicit, hash-bound public proofs outside frozen inputs; never whole dirs."""
    inventory=source/'state/backup_evidence.json'
    result={}
    if not inventory.exists():return result
    for record in read(inventory)['files']:
        rel=Path(record['path'])
        if (rel.is_absolute() or '..' in rel.parts or '\\' in record['path']
                or rel.parts[0] in ('cvs','backups') or record.get('role')!='rule_evidence'):
            raise ValueError('Недопустимый путь подтверждения правил')
        path=source/rel
        if path.is_symlink() or not path.is_file() or sha(path)!=record['sha256']:
            raise ValueError('Изменён файл подтверждения правил: '+rel.as_posix())
        result[rel.as_posix()]=record['sha256']
    return result


def diagnostic_inventory(source, require_files=True):
    inventory=source/'state/backup_exclusions.json';result={}
    if not inventory.exists():return result
    for record in read(inventory)['files']:
        rel=Path(record['path'])
        if (rel.is_absolute() or '..' in rel.parts or '\\' in record['path']
                or len(rel.parts)<3 or rel.parts[0]!='state'
                or rel.parts[1] in ('runs','history','source_blocks','cards')
                or record.get('role')!='diagnostic_copy' or not record.get('reason')):
            raise ValueError('Нельзя исключить рабочие данные из бэкапа')
        path=source/rel
        if ((require_files or path.exists()) and
                (not path.is_file() or path.is_symlink() or sha(path)!=record['sha256'])):
            raise ValueError('Изменена исключаемая диагностическая копия: '+str(rel))
        result[rel.as_posix()]=record
    return result


def exclusion_reason(path, source, evidence, skip_cvs=True, diagnostics=None):
    rel=path.relative_to(source);parts=rel.parts
    if path.is_symlink():return 'symlink'
    if parts[0]=='backups':return 'previous_backup'
    if skip_cvs and parts[0]=='cvs':return 'cv_original'
    if any(x.lower() in ('__pycache__','.git','audio_models','audio_env','.venv','venv','node_modules','lo_profile') for x in parts):
        return 'runtime_cache'
    if path.name.endswith(('-wal','-shm','.tmp','.crdownload')) or path.name.startswith('~$'):
        return 'temporary'
    if rel.as_posix() in (diagnostics or {}):return 'diagnostic_copy'
    if path.suffix.lower() in ('.tar','.gz','.zip','.7z'):return 'nested_archive'
    if rel.as_posix() in evidence:return None
    # Criteria may legitimately contain PDF/image evidence. Frozen input paths
    # and the operator's full source set have an explicit non-CV purpose.
    if parts[0]=='sources' or 'inputs' in parts:return None
    if not skip_cvs:return None
    if path.suffix.lower() in CV_BINARY_SUFFIXES:return 'cv_binary_or_render'
    if path.suffix.lower()=='.xlsx' and path.stem.isdigit() and parts[0]!='exports':
        return 'cv_questionnaire'
    if not path.suffix:
        with path.open('rb') as stream:head=stream.read(16)
        if head.startswith((b'%PDF',b'PK\x03\x04',b'\xd0\xcf\x11\xe0',b'\x89PNG',b'\xff\xd8\xff',b'{\\rtf')):
            return 'cv_binary_without_extension'
    return None

def snapshot(source,dest,strip_originals=True):
    dest.parent.mkdir(parents=True,exist_ok=True)
    src=sqlite3.connect(source.resolve().as_uri()+'?mode=ro',uri=True)
    dst=sqlite3.connect(dest)
    src.backup(dst);src.close()
    tables={x[0] for x in dst.execute("SELECT name FROM sqlite_master WHERE type='table'")}
    omitted=0
    if strip_originals and 'source_files' in tables:
        omitted=dst.execute('SELECT COUNT(*) FROM source_files').fetchone()[0]
        dst.execute('DELETE FROM source_files');dst.commit();dst.execute('VACUUM')
    assert dst.execute('PRAGMA integrity_check').fetchone()[0]=='ok'
    dst.close();return omitted

def archive_tree(source,target,skip_cvs=True):
    """Snapshot databases into staging before compressing; manifest covers every file."""
    source=Path(source);target=Path(target);target.parent.mkdir(parents=True,exist_ok=True)
    if target.exists():raise ValueError('Архив уже существует; выберите новое имя')
    # Excluded scratch files are absent after restore. Preserve the record;
    # still enforce its hash if a file with that path exists again.
    evidence=evidence_inventory(source);diagnostics=diagnostic_inventory(source,require_files=False);excluded=Counter()
    manifest={'format':'hr-hive-backup-v1','created_utc':stamp(),'files':[],
              'excluded_originals':skip_cvs,'excluded_files':[],'rule_evidence':evidence,
              'excluded_diagnostics':list(diagnostics.values())}
    temporary=target.with_name(target.name+'.tmp')
    with tempfile.TemporaryDirectory() as t:
        stage=Path(t)
        with tarfile.open(temporary,'w:gz',compresslevel=5) as tar:
            for p in sorted(source.rglob('*')):
                if not p.is_file():continue
                if p.resolve() in (target.resolve(),temporary.resolve(),target.with_name(target.name+'.sha256').resolve()):continue
                rel=p.relative_to(source)
                reason=exclusion_reason(p,source,evidence,skip_cvs,diagnostics)
                if reason:
                    excluded[reason]+=1
                    # Original hashes/paths already live in candidate manifests.
                    if reason not in ('cv_original','previous_backup','runtime_cache'):
                        manifest['excluded_files'].append({'path':rel.as_posix(),'reason':reason})
                    continue
                use=p;stripped=0
                if p.suffix=='.sqlite':
                    use=stage/rel;stripped=snapshot(p,use)
                rec={'path':rel.as_posix(),'size':use.stat().st_size,'sha256':sha(use)}
                if stripped:rec['excluded_cv_blobs']=stripped
                manifest['files'].append(rec)
                tar.add(use,arcname=rel.as_posix(),recursive=False)
            manifest['excluded_counts']=dict(excluded)
            payload=json.dumps(manifest,ensure_ascii=False,indent=2).encode()
            import io
            info=tarfile.TarInfo('BACKUP_MANIFEST.json');info.size=len(payload);tar.addfile(info,io.BytesIO(payload))
    temporary.replace(target)
    checksum=sha(target);target.with_name(target.name+'.sha256').write_text(checksum+'  '+target.name+'\n',encoding='ascii')
    return {'path':str(target),'sha256':checksum,'files':len(manifest['files']),'bytes':target.stat().st_size}

def restore(path,destination=None):
    dest=Path(destination) if destination else DATA
    dest.mkdir(parents=True,exist_ok=True)
    if any(p.name not in ('backups','cvs') for p in dest.iterdir()):
        raise ValueError('Восстановление разрешено только в пустой каталог данных (cvs и backups допустимы)')
    with tarfile.open(path,'r:gz') as tar:
        manifest=json.load(tar.extractfile('BACKUP_MANIFEST.json'))
        records={r['path']:r for r in manifest['files']}
        if len(records)!=len(manifest['files']):raise ValueError('Повторные пути в манифесте')
        # Verify every member before writing anything.
        seen=set()
        for member in tar.getmembers():
            if member.name=='BACKUP_MANIFEST.json':continue
            rel=Path(member.name)
            if not member.isfile() or rel.is_absolute() or '..' in rel.parts or '\\' in member.name:
                raise ValueError('Недопустимый путь архива: '+member.name)
            if member.name in seen or member.name not in records:raise ValueError('Неожиданный файл архива')
            seen.add(member.name);rec=records[member.name]
            h=hashlib.sha256();size=0
            with tar.extractfile(member) as f:
                for block in iter(lambda:f.read(1024*1024),b''):h.update(block);size+=len(block)
            if h.hexdigest()!=rec['sha256'] or size!=rec['size']:raise ValueError('Не совпадает контрольная сумма: '+member.name)
        if seen!=set(records):raise ValueError('Отсутствуют файлы архива')
        for member in tar.getmembers():
            if member.name=='BACKUP_MANIFEST.json':continue
            out=dest/member.name;out.parent.mkdir(parents=True,exist_ok=True)
            with tar.extractfile(member) as src,out.open('wb') as dst:shutil.copyfileobj(src,dst)
    return {'restored':len(records),'destination':str(dest)}

def migrate(workspace,project):
    """One-off migration. Original workspace is read-only throughout."""
    workspace=Path(workspace).resolve();project=Path(project).resolve();data=project/'data'
    if (data/'state'/'active.json').exists():raise ValueError('Migration already exists')
    data.mkdir(exist_ok=True)
    for name in ('cvs','sources','state/runs','exports','backups','state/history'):(data/name).mkdir(parents=True,exist_ok=True)
    omissions=[];copied=0
    for p in workspace.rglob('*'):
        if not p.is_file():continue
        rel=p.relative_to(workspace)
        if any(x.lower() in EXCLUDED_DIRS for x in rel.parts):continue
        if p.suffix.lower() in EXCLUDED_SUFFIXES or p.name.endswith(('-wal','-shm','.tmp')):continue
        out=data/'state/history'/rel;out.parent.mkdir(parents=True,exist_ok=True)
        if p.suffix=='.sqlite':
            omitted=snapshot(p,out);omissions.append({'path':rel.as_posix(),'original_cv_blobs_removed':omitted})
        else:shutil.copy2(p,out)
        copied+=1
    for name in ('clusters','banlist'):
        shutil.copytree(workspace/name,data/'sources'/name,dirs_exist_ok=True)
    for p in workspace.glob('*.xlsx'):
        if p.name=='Матка №1.xlsx' or p.name=='образец.xlsx':shutil.copy2(p,data/'sources'/p.name)
        shutil.copy2(p,data/'exports'/p.name)
    audit=workspace/'resume_audit'
    for src in (audit/'runs').iterdir():
        if not (src/'state.sqlite').exists():continue
        # All historical bytes are retained above; operational runs are independent.
        dest=data/'state/runs'/src.name;dest.mkdir()
        for name in ('manifest.json','profile_catalog.json','resume_matching_algorithm.json','semantic_reviews.jsonl'):
            if (src/name).exists():shutil.copy2(src/name,dest/name)
        snapshot(src/'state.sqlite',dest/'state.sqlite')
        if (src/'semantic_queue.sqlite').exists():snapshot(src/'semantic_queue.sqlite',dest/'semantic_queue.sqlite')
        if (src/'inputs').exists():shutil.copytree(src/'inputs',dest/'inputs')
        if (src/'engine/geography').exists():shutil.copytree(src/'engine/geography',dest/'engine/geography')
        meta=read(src/'manifest.json')
        books=list((dest/'inputs').glob('*.xlsx'))
        source=next((p for p in books if p.name=='Матка №1.xlsx'),books[0] if books else None)
        if not source:continue
        from .service import rules_hash,prepare
        output=Path(str(meta.get('output','Результаты проверки.xlsx')).replace('\\','/')).name
        dump(dest/'portable.json',{'source_book':source.relative_to(dest).as_posix(),'sheet':meta['sheet'],'output_name':output,'rules_hash':rules_hash(dest/'inputs'),'migrated_from':src.name,'created':stamp()})
        prepare(dest)
    current=read(audit/'active_full_review.json')['run'].replace('\\','/').rsplit('/',1)[-1]
    dump(data/'state/active.json',{'run':current})
    dump(data/'state/migration.json',{'created':stamp(),'history_files':copied,'sqlite_snapshots':omissions,'raw_cv_files_excluded':True,'notes':'Оригинальные пути сохранены как история. Рабочие пути берутся из portable.json и data/cvs; отсутствие оригиналов не сбрасывает проверки.'})
    return {'history_files':copied,'active':current,'databases':len(omissions)}
