"""Commit semantic reviews with cited facts and profile-level component scores.

The JSONL is an assessment journal, not a table of routing exceptions. No ID is
used in the scoring or ban rules. A new CV requires a complete semantic review.
"""
import json,re,hashlib,shutil
from types import SimpleNamespace
import openpyxl
base = SimpleNamespace(openpyxl=openpyxl)



# Subject-teaching profiles and explicit instructor profiles require a
# pedagogical qualification.  Administrative and technical SАОК АП roles
# (for example ПТО, КИПиА, ОТ/ПБ, монтаж and maintenance) use their own
# education requirements from the profile row.
TEACHING=set(range(91,105))|set(range(126,132))|set(range(147,150))
ALIASES={'Департамент логистики':'ДЛ','ЦС (проектный блок)':'ЦС (ПБ)','HR':'Управление HR'}
WEIGHTS=[40,30,20,10]

def catalog(run):
 # Independent reader: no import or execution of the previous automatic matcher.
 book=base.openpyxl.load_workbook(next((run/'inputs'/'clusters').glob('*.xlsx')),read_only=True,data_only=True)
 result={}
 for rowno,v in enumerate(book['Кого берём'].iter_rows(values_only=True),1):
  if len(v)<10 or not v[2] or not v[3]:continue
  if str(v[2]).strip()=='Кластер/проект':continue
  if not any(v[k] for k in range(5,10)):continue
  if any(re.search(r'\b(?:образец|демонстрационный)\b',str(v[k] or ''),re.I) for k in range(1,5)):continue
  key=f'N{rowno:03}'
  result[key]=dict(key=key,row=rowno,module=v[1],cluster=str(v[2]).strip(),profile=str(v[3]).strip(),grade=v[4],education=v[5],experience=v[6],skills=v[7],stops=v[8],comment=v[9])
 book.close()
 result['USER_EXIM']=dict(key='USER_EXIM',cluster='УПИ',profile='Эксим: продажи',education='Любое',experience='Любой',skills='Клиенты, переговоры, сделки, коммерческие результаты',source='Согласованное пользовательское дополнение')
 result['USER_TOP']=dict(key='USER_TOP',cluster='Топ-менеджер',profile='Управление предприятием / запуск и масштабирование бизнеса',education='Любое',experience='Подтвержденные результаты',skills='Управление бизнесом, запуск, масштабирование, измеримые результаты',source='Согласованное пользовательское дополнение')
 result['USER_STA']=dict(key='USER_STA',cluster='СТА',profile='Резервное направление для фактических строительных работ',education='Любое',experience='Любой',skills='Практические строительные операции, материалы, инструменты, качество работ',source='Согласованное правило: СТА как резерв при реальном строительном опыте')
 for cluster,key in [('УПИ','N058'),('ССР','N166'),('УЦТ','N208'),('Цифровизация','N063'),('УКС','N341'),('САОК АП','N091'),('КБ','N255'),('АлМа','N247'),('Завод №2','N301'),('АГД','N016'),('УЭиЭ','N226')]:
  cluster=result[key]['cluster']
  result['QA_'+key]=dict(key='QA_'+key,cluster=cluster,profile='Тестирование ПО',education='Любое',experience='Любой',skills='Тестирование ПО: сценарии, дефекты, API/UI, регрессия',source='Согласованное правило QA; профиль разработки '+key)
 return result

def unrestricted(value):return not str(value or '').strip() or bool(re.match(r'^\s*(?:любо[йеая]+|не важно|не обязательно|не требуется|без опыта|без образования|[-–—])(?:\s*(?:[;,.]|$))',str(value or ''),re.I))

def assess(card,profiles):
 assessed=[];rejected=[]
 rejected.extend(card.get('exclusions',[]))
 for review in card.get('reviews',[]):
  key,components,family,basis,*extra=review
  assert key in profiles,key
  p=profiles[key];assert family in card['tracks'],(card['id'],family)
  if any(x.get('profile')==key for x in rejected):continue
  missing=extra[0] if extra else ''
  if (key=='QA_N091' or key.startswith('N') and int(key[1:]) in TEACHING) and not card.get('pedagogy'):
   rejected.append(dict(profile=key,reason='Нет подтвержденной педагогической квалификации'));continue
  if p['cluster']=='КБ' and not key.startswith('QA_') and family not in ['mechanical_design','mechanical_technology','chemical_technology','cnc','electronics','embedded','rf','aerodynamics','metallurgy','hardware_test','software_engineering','industrial_engineering','fpga','electronics_engineering','aerospace_engineering']:
   rejected.append(dict(profile=key,reason='Нет содержательной связи с инженерной работой'));continue
  if p['cluster']=='СТА' and not card.get('construction_work'):
   rejected.append(dict(profile=key,reason='Нет фактической строительной деятельности'));continue
  if key=='USER_TOP' and not card.get('business_results'):
   rejected.append(dict(profile=key,reason='Нет конкретных результатов бизнеса / предприятия'));continue
  assert len(components)==4
  active=[True,True,not unrestricted(p.get('experience')),not unrestricted(p.get('education'))]
  total=0;denom=0;detail=[]
  for category,weight,value,use in zip(['tasks','skills','experience','education'],WEIGHTS,components,active):
   assert 0<=value<=weight,(card['id'],key,category,value)
   if use:total+=value;denom+=weight
   detail.append(dict(category=category,base_weight=weight,points=value if use else None,used=use,effective_weight=0))
  for d in detail:d['effective_weight']=round(d['base_weight']*100/denom,6) if d['used'] else 0
  score=round(100*total/denom)
  assessed.append(dict(profile_key=key,profile=p['profile'],cluster=ALIASES.get(p['cluster'],p['cluster']),score=score,family=family,components=detail,basis=basis,missing=missing,evidence=card['tracks'][family][2],profile_requirements=p))
 # Evidence of two different professions, not one profession across departments.
 # The reviewer explicitly groups independent directions; no age is read.
 independent=card.get('independent_directions',[])
 if len(independent)>=2:
  best=[]
  used_families=set()
  for group in independent:
   eligible=[r for r in assessed if r['family'] in group and r['family'] not in used_families and r['score']>=60 and r['cluster'] not in ['Топ-менеджер','100 Лидеров']]
   if eligible:
    item=max(eligible,key=lambda r:r['score']);best.append(item);used_families.update(group)
  best.sort(key=lambda r:r['score'],reverse=True)
  if len(best)>=2:assessed.append(dict(profile_key='USER_100',profile='Два независимых профессиональных направления',cluster='100 Лидеров',score=min(best[0]['score'],best[1]['score']),family='multifunctional',basis='; '.join(r['family']+': '+r['basis'] for r in best[:2]),evidence=[r['evidence'] for r in best[:2]],missing='',direction_scores=[r['score'] for r in best[:2]]))
 clusters={}
 for r in assessed:
  if r['score']>=60 and (r['cluster'] not in clusters or r['score']>clusters[r['cluster']]['score']):clusters[r['cluster']]=r
 ranked=sorted(clusters.values(),key=lambda r:(-r['score'],r['cluster']))
 # Construction fallback only if no stronger construction direction is available.
 if any(r['cluster'] in ['АлДе','АлМа','УКС','УГП','СРНП','Департамент №4'] and r['family'].startswith('construction') for r in ranked):ranked=[r for r in ranked if r['cluster']!='СТА']
 chosen=ranked[:3]
 return assessed,rejected,chosen
