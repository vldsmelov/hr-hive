"""Compare migrated operational records to original databases without altering originals."""
import json
import sqlite3
from pathlib import Path
from app.common import dump,stamp,DATA
from app.books import export

original=Path('..')/'resume_audit/runs'
report=[]
for run in (DATA/'state/runs').iterdir():
    if not (run/'portable.json').exists():continue
    source=original/run.name/'state.sqlite'
    old=sqlite3.connect(source.resolve().as_uri()+'?mode=ro',uri=True);old.row_factory=sqlite3.Row
    new=sqlite3.connect(run/'state.sqlite');new.row_factory=sqlite3.Row
    before={r['id']:dict(r) for r in old.execute('SELECT * FROM candidates')}
    after={r['id']:dict(r) for r in new.execute('SELECT * FROM candidates')}
    assert before.keys()==after.keys()
    corrected=0
    for ident,r in before.items():
        n=after[ident]
        for field in ('id','source_row','files','fingerprint','extraction_status','review_status','text','documents','error','reviewed_utc'):
            assert r[field]==n[field],(run.name,ident,field)
        if r['result']:
            a=json.loads(r['result']);b=json.loads(n['result'])
            for field in ('ban','department','clusters','ban_assessment','profile_assessments','excluded_profiles'):
                assert a.get(field)==b.get(field),(run.name,ident,field)
            corrected+=a.get('comment')!=b.get('comment')
    tables={r[0] for r in new.execute("SELECT name FROM sqlite_master WHERE type='table'")}
    if 'source_files' in tables:assert new.execute('SELECT COUNT(*) FROM source_files').fetchone()[0]==0
    assert new.execute('PRAGMA integrity_check').fetchone()[0]=='ok'
    old.close();new.close()
    report.append({'run':run.name,'candidates':len(before),'comments_restored':corrected,'all_other_assessment_fields_equal':True})
dump(DATA/'state/migration_verification.json',{'created':stamp(),'runs':report})
print(json.dumps(report,ensure_ascii=False,indent=2))
for name in ('incoming_1194_8848_20261020_130442','full_20261016_132123'):
    print(json.dumps(export(DATA/'state/runs'/name),ensure_ascii=False))
