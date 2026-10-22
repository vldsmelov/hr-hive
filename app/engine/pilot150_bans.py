"""Screen reviewed employment records, never free text clients or education."""
import re,calendar
from datetime import date
from .ban_engine import BanEngine,company_name,norm,RT,LOCAL,OTHER_REGION,add_months,DamerauLevenshtein

def end_date(value):
 if not value or value=='present':return None
 parts=str(value).split('-')
 if len(parts)==1:return date(int(parts[0]),12,31)
 if len(parts)==2:return date(int(parts[0]),int(parts[1]),calendar.monthrange(int(parts[0]),int(parts[1]))[1])
 return date.fromisoformat(value)

class StructuredBans:
 def __init__(self,run):
  self.base=BanEngine(run);self.as_of=self.base.as_of
  self.rules={r['row']:r for r in self.base.company_rules}
  # The table names the holding; its explicit group name in a subsidiary name
  # is an affiliation match, not a guessed relationship.
  for r in self.base.company_rules:
   if 'таграс холдинг' in norm(r['name']):self.base.company_patterns[r['row']].append(('таграс / tagras',re.compile(r'\b(?:таграс|tagras)\b')))
   if r['row']==255:self.base.company_patterns[r['row']].append(('ПТФК ЗТЭО / завод ТЭО',re.compile(r'\bптфк\s+(?:зтэо|завод тэо)\b')))
   if r['row']==85:self.base.company_patterns[r['row']].append(('НГДУ Елховнефть',re.compile(r'\bелховнефть\b')))
 def finding(self,rule,job,certainty='confirmed',reason='Совпадение работодателя'):
  employer,start,end,role,region,kind=job[:6]
  finish=end_date(end);months=rule['months']
  expiry=add_months(finish,months) if finish and months else None
  active=months!=0 and (months is None or finish is None or expiry>self.as_of)
  if rule.get('term_unknown'):certainty='probable';reason+='; в бан-листе не указан срок'
  if months and not end:certainty='probable';reason+='; дата увольнения неизвестна'
  if len(job)>6 and isinstance(job[6],dict) and job[6].get('employment_uncertain'):
   certainty='probable';reason+='; '+job[6]['employment_uncertain']
  return dict(type='employment',row=rule['row'],rule=rule['name'],stage=rule['stage'],months=months,employer=employer,role=role,region=region,organization_type=kind,period=[start,end],evidence=' | '.join(str(x) for x in job[:6]),basis=reason,certainty=certainty,active=active,termination_upper_bound=finish.isoformat() if finish else None,expires=expiry.isoformat() if expiry else None,calculation_date=self.as_of.isoformat(),source='user_addendum' if str(rule['row']).startswith('USER_') else 'ФИО.csv' if str(rule['row']).startswith('ФИО:') else 'Компании.csv')
 def assess(self,card,source_values):
  # Only reviewed candidate name and candidate contact columns, not references.
  identity='\n'.join([card.get('name',''),*card.get('previous_names',[]),*card.get('phones',[])])
  sv=list(source_values)
  parts=card.get('name','').split()
  if len(parts)>=2 and not sv[5] and not sv[6]:sv[5:8]=[parts[0],parts[1],' '.join(parts[2:]) or None]
  findings,phones,names=self.base.identities(identity,sv)
  for f in findings:f.update(source='ФИО.csv',stage='100%',expires=None)
  for job in card['jobs']:
   employer,start,end,role,region,kind=job[:6]
   ne=company_name(employer);geo=norm(region);rt=bool(re.search(RT,geo));local=bool(re.search(LOCAL,geo))
   elsewhere=bool(geo and not rt and (re.search(OTHER_REGION,geo) or geo not in ['неизвестно','рф','россия','unknown']))
   if re.search(r'транснефт|transneft',ne):
    findings.append(self.finding(dict(row='USER_TRANSNEFT_GROUP',name='Вся группа Транснефть',stage='100%',months=None),job))
   for r in self.base.company_rules:
    if r['generic'] or r['months']==0:continue
    # A duplicated company in the personal list has no stated term. Use the
    # explicit term in the company table instead of inventing an indefinite ban.
    if r.get('term_unknown') and any(a in x['aliases'] for a in r['aliases'] for x in self.base.company_rules if not x.get('term_unknown') and x['months']!=0):continue
    hit=None
    for alias,pattern in self.base.company_patterns[r['row']]:
     if pattern.search(ne):hit=('probable' if alias=='ПТФК ЗТЭО / завод ТЭО' else 'confirmed',alias);break
     if len(alias)>=6 and ' ' not in alias and alias not in ['импульс','синтез','фаворит','основание','активы']:
      if any(len(w)>=5 and DamerauLevenshtein.distance(alias,w)==1 for w in ne.split()):hit=('probable',alias)
    if not hit:continue
    if isinstance(r['row'],int) and 89<=r['row']<=238 and elsewhere:continue
    reason='Совпадение работодателя: '+hit[1]
    if hit[1]=='ПТФК ЗТЭО / завод ТЭО':reason+='; уточнить юридическое совпадение сокращений ЗТЭО/ТЭО'
    findings.append(self.finding(r,job,hit[0],reason))
   applicable=[]
   # Organization types are reviewed from work history; absence alone is not a prohibition.
   if kind.startswith('security'):
    if 'росгвард' in ne:applicable=[(17 if local else 242 if rt else 87,'confirmed' if region else 'probable')]
    elif local:applicable=[(15,'confirmed')]
    elif rt:applicable=[(240 if 'мчс' in ne else 239,'confirmed')]
    elif elsewhere:applicable=[(88,'confirmed')]
    else:applicable=[(15,'probable')]
   elif kind in ['public_edu','edu_unknown','private_edu']:
    if kind!='private_edu':
     if rt:applicable=[(12,'confirmed' if kind=='public_edu' else 'probable')]
     elif elsewhere:applicable=[(264,'confirmed' if kind=='public_edu' else 'probable')]
     else:applicable=[(12,'probable')]
   elif kind in ['public_med','med_unknown','private_med']:
    if kind!='private_med':
     if rt:applicable=[(91,'confirmed' if kind=='public_med' else 'probable')]
     elif elsewhere:applicable=[(263,'confirmed' if kind=='public_med' else 'probable')]
     else:applicable=[(91,'probable')]
   elif kind=='public':
    if rt:applicable=[(2,'confirmed')]
    elif not elsewhere:applicable=[(2,'probable')]
   if re.search(r'\b(?:ржд|российские железные дороги)\b',ne) and not elsewhere:applicable.append((393,'confirmed' if rt else 'probable'))
   for row,certainty in applicable:
    reason='Трудовая деятельность по отраслевому правилу'
    if certainty=='probable':reason+='; требуется уточнить '+('регион работодателя' if not rt and not elsewhere else 'государственный или частный статус организации')
    findings.append(self.finding(self.rules[row],job,certainty,reason))
  unique={str((f['row'],f.get('employer'),f.get('period'),f.get('basis'))):f for f in findings}
  findings=list(unique.values());active=[f for f in findings if f['active']]
  people={p['row']:p for p in self.base.people}
  for f in findings:
   if f['source']=='user_addendum':
    f['rule_record']={'rule':'Вся группа Транснефть','term':'100%','source':'Согласованное дополнение пользователя'}
    f['source_snapshot']='inputs/resume_audit/test150_overrides.json'
   else:
    f['rule_record']=people.get(f['row'],{}) if f['type']=='person' else self.rules.get(f['row'],{})
    f['source_snapshot']='inputs/banlist/'+f['source']
   key=next((k for k in self.base.meta['input_hashes'] if k.replace('\\','/').endswith(f['source_snapshot'][7:])),None)
   f['source_sha256']=self.base.meta['input_hashes'].get(key)
  active.sort(key=lambda f:(f.get('months') is None,f.get('expires') or '9999',f['certainty']=='confirmed'),reverse=True)
  return dict(ban='да' if active else 'нет',findings=findings,active=active,candidate_names=names,phones=sorted(p for p in phones if p))

def brief(b):
 if not b['active']:return ''
 f=b['active'][0]
 label=('До уточнения: ' if f['certainty']=='probable' else '')
 label+=(f.get('employer') or f.get('candidate_name') or f['rule'])
 source='уточнение пользователя: вся группа Транснефть' if f['source']=='user_addendum' else f['source']+', строка '+str(f['row'])
 label+=' — '+source
 if f['certainty']=='probable':
  label+='; '+f['basis']
 if f.get('expires'):label+='; до '+f['expires']
 elif f['certainty']=='probable':label+='; применимость и срок требуют уточнения'
 elif f.get('months'):label+='; '+str(f['months'])+' мес. после увольнения, работа продолжается'
 else:label+='; бессрочно'
 return label
