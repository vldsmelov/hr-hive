import argparse
import importlib.metadata
import json
import platform
import shutil
import subprocess
import sys
from pathlib import Path
from .common import DATA,active,layout,read

def main():
    p=argparse.ArgumentParser(description='HR Hive: полные документы, семантическая оценка Codex и устойчивые результаты')
    p.add_argument('command',choices=['doctor','status','use','scan','import-archive','amend-archive','extract','retry','claim','show','preview','submit','reject-source','export','backup','restore','profiles'])
    p.add_argument('--worker',default='codex');p.add_argument('--count',type=int,default=5)
    p.add_argument('--id');p.add_argument('--ids',nargs='+');p.add_argument('--offset',type=int,default=0);p.add_argument('--limit',type=int,default=10000)
    p.add_argument('--file');p.add_argument('--output');p.add_argument('--destination');p.add_argument('--all',action='store_true');p.add_argument('--new-version',action='store_true')
    p.add_argument('--run')
    p.add_argument('--engine-snapshot');p.add_argument('--template');p.add_argument('--workers',type=int,default=1)
    p.add_argument('--as-of',help='Дата проверки архива, YYYY-MM-DD (по умолчанию сегодня в Москве)')
    p.add_argument('--partner-registry',help='Проверенный снимок реестра 262 для даты нового архива')
    p.add_argument('--criteria-dir',help='Полный набор clusters/, banlist/ и дополнений для новой версии архива')
    p.add_argument('--generated-from-resumes',action='store_true',help='Новая таблица из подтверждённых сведений CV')
    a=p.parse_args()
    from . import service
    if a.command=='doctor':
        langs=subprocess.run(['tesseract','--list-langs'],capture_output=True,text=True) if shutil.which('tesseract') else None
        result={'python':sys.version,'architecture':platform.machine(),'data_directory':str(DATA),'libreoffice':shutil.which('soffice'),'tesseract':shutil.which('tesseract'),'ocr_languages':langs.stdout if langs else None,'dependencies':{x:importlib.metadata.version(x) for x in ('openpyxl','PyMuPDF','Pillow','lxml','rapidfuzz')}}
        if not result['libreoffice'] or not langs or not all(x in langs.stdout.split() for x in ('rus','eng')):raise RuntimeError('Неполное окружение OCR/LibreOffice')
    elif a.command=='status':result=service.status(a.all)
    elif a.command=='use':
        if not a.run or Path(a.run).name!=a.run or not (DATA/'state/runs'/a.run/'portable.json').is_file():p.error('--run должен быть именем существующего запуска из status --all')
        from .common import dump
        dump(DATA/'state/active.json',{'run':a.run});result={'active_run':a.run}
    elif a.command=='scan':result=service.scan(a.new_version)
    elif a.command=='import-archive':
        if not a.file or not a.template:p.error('--file и --template обязательны')
        from .batch import import_archive
        result=import_archive(a.file,a.template,a.output,a.as_of,a.partner_registry,
                              generated_from_resumes=a.generated_from_resumes,criteria_dir=a.criteria_dir)
    elif a.command=='amend-archive':
        if not a.run or not a.file or not a.engine_snapshot:p.error('--run, --file и --engine-snapshot обязательны')
        from .batch import amend_archive
        result=amend_archive(a.run,a.file,a.engine_snapshot,a.output)
    elif a.command in ('extract','retry'):
        if not 1<=a.workers<=2:p.error('--workers: допустимо 1 или 2')
        if a.workers>1:
            from .batch import extract_parallel
            result=extract_parallel(a.workers,a.command=='retry',a.ids)
        else:result=service.extract(a.command=='retry',a.ids)
    elif a.command=='claim':result=service.claim(a.worker,a.count)
    elif a.command=='show':
        if not a.id:p.error('--id обязателен')
        result=service.show(a.id,a.offset,a.limit)
    elif a.command=='reject-source':
        if not a.id or not a.file:p.error('--id и --file отчёта обязательны')
        result=service.reject_source(a.id,a.file,a.worker)
    elif a.command in ('preview','submit'):
        if not a.file:p.error('--file обязателен')
        cards=[json.loads(x) for x in Path(a.file).read_text('utf-8-sig').splitlines() if x.strip()]
        result=[service.preview(c,a.worker) if a.command=='preview' else service.submit(c,a.worker) for c in cards]
    elif a.command=='profiles':
        from .engine.profiles import catalog
        result=catalog(active())
    elif a.command=='export':
        from .books import export
        result=export(active(),a.output)
    elif a.command=='backup':
        from .archives import archive_tree
        layout();result=archive_tree(DATA,a.output or DATA/'backups'/'hr-hive-backup.tar.gz')
    elif a.command=='restore':
        if not a.file:p.error('--file обязателен')
        from .archives import restore
        result=restore(a.file,a.destination)
    print(json.dumps(result,ensure_ascii=False,indent=2))

if __name__=='__main__':
    try:main()
    except Exception as exc:
        print(json.dumps({'error':str(exc),'type':type(exc).__name__},ensure_ascii=False),file=sys.stderr)
        sys.exit(1)
