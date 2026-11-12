"""Capture/verify a portable data passport. Run only inside Docker, without writers."""
import argparse
import hashlib
import json
import sqlite3
from pathlib import Path

from app import archives, service
from app.common import DATA, dump, read, sha, stamp


def database(path):
    connection=sqlite3.connect(path.resolve().as_uri()+'?mode=ro',uri=True)
    try:
        connection.execute('PRAGMA query_only=ON')
        connection.execute('BEGIN')
        integrity=[row[0] for row in connection.execute('PRAGMA integrity_check')]
        if integrity!=['ok']:raise ValueError('Повреждена база: '+str(path))
        tables={}
        for name,definition in connection.execute("SELECT name,sql FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%' ORDER BY name"):
            quoted='"'+name.replace('"','""')+'"'
            rows=[]
            if name!='source_files':
                for row in connection.execute('SELECT * FROM '+quoted):
                    values=[{'blob_sha256':hashlib.sha256(v).hexdigest()} if isinstance(v,bytes) else v for v in row]
                    rows.append(json.dumps(values,ensure_ascii=False,separators=(',',':')))
            digest=hashlib.sha256('\n'.join(sorted(rows)).encode()).hexdigest()
            tables[name]={'rows':len(rows),'sha256':digest,'schema':definition}
        return {'integrity':'ok','tables':tables}
    finally:connection.close()


def inspect(root):
    root=Path(root);evidence=archives.evidence_inventory(root)
    diagnostics=archives.diagnostic_inventory(root,require_files=False)
    result={'active':read(root/'state/active.json'),'databases':{},'files':{},'runs':{},
            'excluded_diagnostics':diagnostics,'runtime':{'implementation_hash':service.implementation_hash()}}
    for path in sorted(root.rglob('*')):
        if not path.is_file() or archives.exclusion_reason(path,root,evidence,diagnostics=diagnostics):continue
        rel=path.relative_to(root).as_posix()
        if rel in ('state/transfer_passport.json','state/transfer_verification.json'):continue
        if path.suffix=='.sqlite':result['databases'][rel]=database(path)
        else:result['files'][rel]=sha(path)
    for path in sorted((root/'state/runs').glob('*/portable.json')):
        run=path.parent;meta=read(path);manifest=read(run/'manifest.json')
        connection=sqlite3.connect((run/'state.sqlite').resolve().as_uri()+'?mode=ro',uri=True)
        try:
            counts=[list(row) for row in connection.execute('SELECT extraction_status,review_status,COUNT(*) FROM candidates GROUP BY extraction_status,review_status ORDER BY 1,2')]
        finally:connection.close()
        result['runs'][run.name]={'as_of':manifest.get('as_of'),'counts':counts,
            'rules_hash':meta['rules_hash'],'actual_rules_hash':service.rules_hash(run/'inputs'),
            'implementation_hash':meta.get('implementation_hash')}
        if result['runs'][run.name]['rules_hash']!=result['runs'][run.name]['actual_rules_hash']:
            raise ValueError('Изменён снимок правил: '+run.name)
        if meta.get('amendment_key'):
            from app.amendments import verify_amendment_run
            verify_amendment_run(run)
    return result


def main():
    parser=argparse.ArgumentParser(__doc__)
    parser.add_argument('command',choices=['capture','verify'])
    parser.add_argument('--commit');parser.add_argument('--passport',default=str(DATA/'state/transfer_passport.json'))
    parser.add_argument('--report',default=str(DATA/'state/transfer_verification.json'))
    parser.add_argument('--source-bytes',action='store_true',help='Также проверить неизменность исходных файлов SQLite; только на исходном компьютере')
    args=parser.parse_args()
    observed=inspect(DATA)
    if args.command=='capture':
        if not args.commit or len(args.commit)!=40:parser.error('--commit: полный SHA коммита')
        passport={'format':'hr-hive-transfer-v1','created_utc':stamp(),'code_commit':args.commit,
                  'originals_included':False,'baseline':observed,
                  'source_database_sha256':{name:sha(DATA/name) for name in observed['databases']}}
        dump(Path(args.passport),passport)
        print(json.dumps({'passport':args.passport,'databases':len(observed['databases']),'files':len(observed['files']),'active':observed['active']},ensure_ascii=False))
    else:
        passport=read(args.passport);expected=passport['baseline']
        differences=[]
        for section in ('active','databases','files','runs','excluded_diagnostics','runtime'):
            if expected[section]!=observed[section]:
                if section=='active':differences.append('active')
                else:
                    differences.extend(section+':'+key for key in sorted(set(expected[section])|set(observed[section]))
                                       if expected[section].get(key)!=observed[section].get(key))
        if args.source_bytes:
            differences.extend('source_bytes:'+name for name,value in passport['source_database_sha256'].items()
                               if sha(DATA/name)!=value)
        report={'verified_utc':stamp(),'code_commit':passport['code_commit'],'ok':not differences,
                'differences':differences,'databases':len(observed['databases']),
                'files':len(observed['files']),'active':observed['active'],'runs':observed['runs']}
        dump(Path(args.report),report);print(json.dumps(report,ensure_ascii=False,indent=2))
        if differences:raise SystemExit(1)


if __name__=='__main__':main()
