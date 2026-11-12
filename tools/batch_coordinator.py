"""Single SQLite owner: dispatch full immutable packets and save returned cards.

Run inside Docker. Review helpers only read packets and write their own JSONL.
"""
import argparse
import json
from pathlib import Path
from app import service, books
from app.common import DATA, active, connect, dump, read


def packet(worker, count=5):
    run=active();ids=service.claim(worker,count);folder=run/'packets'/worker
    folder.mkdir(parents=True,exist_ok=True)
    source=service.source_rows(run)
    obligations=read(run/'review_obligations.json') if (run/'review_obligations.json').exists() else {}
    paths=[]
    for ident in ids:
        chunks=[];offset=0
        while True:
            part=service.show(ident,offset,20000);chunks.append(part)
            if part['next_offset'] is None:break
            offset=part['next_offset']
        c=connect(run/'state.sqlite')
        try:r=c.execute('SELECT files,documents FROM candidates WHERE id=?',(ident,)).fetchone()
        finally:c.close()
        dest=folder/(ident+'.json')
        values=source[ident][1]
        identity={'surname':values[5],'name':values[6],'patronymic':values[7],'phone':values[10]}
        documents=[]
        for document in json.loads(r['documents']):
            info={key:value for key,value in document.items() if key!='pages'}
            info['pages']=[{'page':page['page'],'method':page['method'],
                            'characters':len(page['text'])} for page in document['pages']]
            documents.append(info)
        dump(dest,{'run':run.name,'worker':worker,'id':ident,'source_identity':identity,'files':json.loads(r['files']),
                   'documents':documents,'chunks':chunks,'review_obligations':obligations.get(ident,[]),
                   'instruction':'Прочитать ВСЕ chunks до next_offset=null, все файлы/страницы. Не использовать прежние оценки.'})
        paths.append(str(dest))
    result={'run':run.name,'worker':worker,'ids':ids,'packets':paths,
            'cards_directory':str(run/'cards'/worker)}
    dump(folder/'assignment.json',result)
    return result


def save_ready(worker, ids=None):
    run=active();folder=run/'cards'/worker;results=[]
    for path in sorted(folder.glob('*.jsonl')):
        ident=path.stem
        if ids is not None and ident not in ids:continue
        c=connect(run/'state.sqlite')
        try:r=c.execute('SELECT review_status FROM candidates WHERE id=?',(ident,)).fetchone()
        finally:c.close()
        if r and r['review_status']=='reviewed':continue
        try:
            cards=[json.loads(line) for line in path.read_text('utf-8-sig').splitlines() if line.strip()]
            if len(cards)!=1 or str(cards[0]['id'])!=ident:raise ValueError('Один ID и одна карточка на файл')
            card=cards[0]
            packet=read(run/'packets'/worker/(ident+'.json'))
            if card.get('read_parts')!=[[part['offset'],part['end']] for part in packet['chunks']]:
                raise ValueError('Карточка не подтверждает чтение всех частей назначенного пакета')
            if card.get('read_files')!=[file['path'] for file in packet['files']]:
                raise ValueError('Карточка не подтверждает проверку всех файлов кандидата')
            result=service.preview(card,worker)
            dump(run/'previews'/worker/(ident+'.json'),result)
            # submit revalidates the immutable source, ownership and full read.
            saved=service.submit(card,worker)
            results.append({'id':ident,'saved':True,'department':result['department'],
                            'comment':result['comment'],'ban':result['ban'],
                            'ban_comment':result['ban_comment']})
        except Exception as error:
            results.append({'id':ident,'saved':False,'error':str(error)})
    if any(r.get('saved') for r in results):
        exported=books.export(run)
    else:exported=None
    result={'run':run.name,'worker':worker,'results':results,'export':exported}
    dump(run/'worker_reports'/(worker+'.json'),result)
    return result


def main():
    p=argparse.ArgumentParser(__doc__);p.add_argument('command',choices=['packet','save-ready'])
    p.add_argument('--worker',required=True);p.add_argument('--count',type=int,default=5)
    p.add_argument('--run',required=True);p.add_argument('--ids',nargs='+');a=p.parse_args()
    if active().name!=a.run:raise ValueError('Активный запуск изменён; остановка координатора')
    if not read(active()/'portable.json').get('import_complete'):raise ValueError('Запуск не является завершённым импортом архива')
    if a.count!=5:raise ValueError('Пакет должен содержать пять кандидатов')
    if Path(a.worker).name!=a.worker or a.worker in ('.','..'):raise ValueError('Недопустимое имя worker')
    if a.command=='packet' and a.ids:raise ValueError('--ids допустим только для save-ready')
    print(json.dumps(packet(a.worker,a.count) if a.command=='packet' else save_ready(a.worker,a.ids),ensure_ascii=False,indent=2))


if __name__=='__main__':main()
