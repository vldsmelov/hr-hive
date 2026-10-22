import hashlib
import json
import os
import sqlite3
from datetime import datetime, timezone
from pathlib import Path

DATA = Path(os.environ.get('HR_DATA', '/data'))

def stamp():
    return datetime.now(timezone.utc).isoformat()

def sha(path):
    h = hashlib.sha256()
    with Path(path).open('rb') as f:
        for block in iter(lambda: f.read(1024 * 1024), b''):
            h.update(block)
    return h.hexdigest()

def read(path):
    return json.loads(Path(path).read_text('utf-8-sig'))

def dump(path, obj):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_name(path.name + '.tmp')
    tmp.write_text(json.dumps(obj, ensure_ascii=False, indent=2), encoding='utf-8')
    os.replace(tmp, path)

def connect(path):
    c = sqlite3.connect(path, timeout=60)
    c.row_factory = sqlite3.Row
    c.execute('PRAGMA journal_mode=WAL')
    c.execute('PRAGMA synchronous=FULL')
    return c

def active():
    return DATA / 'state' / 'runs' / read(DATA / 'state' / 'active.json')['run']

def layout():
    for name in ('cvs', 'sources', 'state/runs', 'exports', 'backups'):
        (DATA / name).mkdir(parents=True, exist_ok=True)
