"""Repair presentation text in migrated copies; never touch source workspace."""
import json
import re
import sqlite3
from pathlib import Path
from app.common import dump,read,stamp

data=Path('data')
changes=[]
for run in (data/'state/runs').iterdir():
    if not (run/'portable.json').exists():continue
    mapping_file=data/'state'/('comment_translations.json' if run.name.startswith('incoming') else 'full_comment_translations.json')
    if not (run.name.startswith('incoming') or run.name.startswith('full_')):continue
    mapping=read(mapping_file)
    c=sqlite3.connect(run/'state.sqlite');c.row_factory=sqlite3.Row
    for row in c.execute("SELECT id,result FROM candidates WHERE review_status='reviewed'").fetchall():
        result=json.loads(row['result']);old=result.get('comment','')
        if row['id'] not in mapping:continue
        new=mapping[row['id']]
        if old==new:continue
        result['comment']=new
        card=result.get('fact_card',{})
        if card:card['comment']=new
        result['presentation_correction']={'created':stamp(),'kind':'restore_specific_russian_comment','assessment_scores_changed':False}
        with c:
            c.execute('INSERT INTO portable_revisions(id,kind,payload,created) VALUES(?,?,?,?)',(row['id'],'comment_before_repair',row['result'],stamp()))
            c.execute('UPDATE candidates SET result=? WHERE id=?',(json.dumps(result,ensure_ascii=False),row['id']))
            fc=c.execute('SELECT card FROM fact_cards WHERE id=?',(row['id'],)).fetchone()
            if fc:
                item=json.loads(fc[0]);item['comment']=new
                c.execute('UPDATE fact_cards SET card=? WHERE id=?',(json.dumps(item,ensure_ascii=False),row['id']))
        changes.append({'run':run.name,'id':row['id'],'old':old,'new':new})
    c.close()
    journal=run/'semantic_reviews.jsonl'
    if journal.exists():
        items=[json.loads(x) for x in journal.read_text('utf-8-sig').splitlines() if x.strip()]
        for item in items:
            if str(item.get('id')) in mapping:item['comment']=mapping[str(item['id'])]
        journal.write_text(''.join(json.dumps(x,ensure_ascii=False)+'\n' for x in items),encoding='utf-8')
dump(data/'state/comment_restoration_report.json',{'created':stamp(),'changed':len(changes),'changes':changes,'scope':'Presentation comments only; original evidence/assessments and all historical files retained.'})
print('Specific comments restored:',len(changes))
