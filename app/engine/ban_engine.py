"""Conservative, evidence-recording identity/employment exclusion screening."""
import calendar,collections,csv,json,re,unicodedata,bisect,functools
from datetime import date
from pathlib import Path
from rapidfuzz.distance import DamerauLevenshtein
from rapidfuzz import process as fuzzy_process
from ..common import active,read
TRANSLATE=str.maketrans({'ё':'е','Ё':'Е','«':' ','»':' ','“':' ','”':' ','\xa0':' '})
def norm(s):return re.sub(r'[^a-zа-я0-9]+',' ',str(s).lower().translate(TRANSLATE)).strip()
def compact(s):return norm(s).replace(' ','')
def normalized_positions(text):
 words=[];starts=[];positions=[];offset=0
 for m in re.finditer(r'[a-zа-яё0-9]+',text,re.I):
  value=norm(m.group());starts.append(offset);positions.append(m.start());words.append(value);offset+=len(value)+1
 normalized=' '.join(words)
 return normalized,lambda p: positions[max(0,bisect.bisect_right(starts,p)-1)] if positions else 0
def phone(s):
 d=re.sub(r'\D','',str(s))
 if len(d)==11 and d[0] in '78':d='7'+d[1:]
 elif len(d)==10:d='7'+d
 else:return ''
 if len(set(d[-10:]))<4 or re.search(r'(\d)\1{5,}',d):return ''
 return d
MONTHS={'январ':1,'феврал':2,'март':3,'апрел':4,'ма':5,'июн':6,'июл':7,'август':8,'сентябр':9,'октябр':10,'ноябр':11,'декабр':12}
MONTHS.update(dict(zip(['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec'],range(1,13))))
MONTH_RE=r'(?:январ[ья]|феврал[ья]|март[а]?|апрел[ья]|ма[йя]|июн[ья]|июл[ья]|август[а]?|сентябр[ья]|октябр[ья]|ноябр[ья]|декабр[ья]|jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|jun(?:e)?|jul(?:y)?|aug(?:ust)?|sep(?:tember)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?)'
DATE_RE=r'(?:'+MONTH_RE+r'\s+(?:19|20)\d{2}|\d{1,2}[./]\d{1,2}[./](?:19|20)\d{2}|\d{1,2}[./](?:19|20)\d{2}|(?:19|20)\d{2})'
RANGE_RE=re.compile(r'(?P<start>'+DATE_RE+r')\s*(?:г\.?\s*)?(?:—|–|-|по|до|to)\s*(?P<end>present|current|настоящ[а-я ]{0,12}|наст\.?\s*вр\.?|н\.?\s*в\.?|сейчас|'+DATE_RE+r')',re.I)
EMPLOY=re.compile(r'^\s*(?:work experience|professional experience|employment history|experience|опыт\s*работы|трудов[а-я -]{0,15}деятельност|работа\s+в\s+прошлом|профессиональн[а-я -]{0,12}опыт)',re.I|re.M)
STOP=re.compile(r'(?:^|[ \t]{4,})(?:образование|образованиe|education|повышение квалификации|курсы|дополнительная информация|рекомендации|другие сведения)(?:\s*[:,]?\s*(?:\n|$))',re.I|re.M)
HR_METADATA=re.compile(r'(?im)^\s*(?:Комментарии\s+к\s+резюме|История\s+общения\s+с\s+кандидатом|История\s+откликов|Recruiter notes|Application history)\b')
def candidate_content(text):
 if not HR_METADATA.search(text):return text
 # Applicant-tracking exports append recruiter names and applications, which are
 # neither the candidate's identity nor employment or professional skills.
 return '\n'.join(HR_METADATA.split(part,maxsplit=1)[0] for part in re.split(r'(?m)(?=^FILE )',text))
def work_text(text):
 if len(re.findall(r'(?m)^FILE ',text))>1:
  parts=[part for part in re.split(r'(?m)^FILE [^\n]*\n',text) if part.strip()]
  sections=[work_text(part) for part in parts]
  return '\n\n'.join(s[0] for s in sections),any(s[1] for s in sections)
 t=re.sub(r'FILE [^\n]+\n|PAGE \d+\n','',text).replace('\r','\n')
 m=EMPLOY.search(t)
 if m:
  start=m.start();end=STOP.search(t,m.end())
  current=[x.group() for x in re.finditer(r'(?im)^.*(?:с\s+\d{4}\s+года\s+работает|в настоящее время работает).{5,250}',t[:start])]
  return '\n'.join(current+[t[start:end.start() if end else len(t)]]),True
 # Freely formatted CVs retain duty sections, removing clearly marked education paragraphs.
 t=re.sub(r'(?ims)^\s*образование\s*:?\s*\n.*?(?=^\s*(?:опыт|навыки|профессиональн|дополнительн|о себе|работа)|\Z)','',t)
 return t,False
def parse_end(s):
 s=norm(s)
 if any(x in s for x in ['наст','сейчас','н в','present','current']):return None
 nums=re.findall(r'\d+',s)
 if not nums:return None
 if re.match(MONTH_RE,s):
  year=int(nums[-1]);month=next(v for k,v in MONTHS.items() if s.startswith(k));return date(year,month,calendar.monthrange(year,month)[1])
 if len(nums)==3:
  try:return date(int(nums[2]),int(nums[1]),int(nums[0]))
  except ValueError:return None
 if len(nums)==2:
  year=int(nums[1]);month=int(nums[0])
  try:return date(year,month,calendar.monthrange(year,month)[1])
  except ValueError:return None
 if len(nums)==1:return date(int(nums[0]),12,31)
 return None
def add_months(d,n):
 m=d.year*12+d.month-1+n;y,m=divmod(m,12);m+=1
 return date(y,m,min(d.day,calendar.monthrange(y,m)[1]))
@functools.lru_cache(maxsize=32)
def employment_periods(text):
 result=[{'pos':m.start(),'stop':m.end(),'start':m.group('start'),'end':m.group('end'),'evidence':m.group()} for m in RANGE_RE.finditer(text)]
 # PDF tables put the employer between the start and end dates, on adjacent lines.
 split=re.compile(r'(?P<start>'+DATE_RE+r')\s*[—–-][^\n\r]{0,160}[\r\n]+\s*(?P<end>настоящее\s+время|по\s+н\.?\s*в\.?|'+DATE_RE+r')',re.I)
 # Word tables often put two date cells on consecutive lines without a dash.
 adjacent=re.compile(r'(?m)^\s*(?P<start>'+DATE_RE+r')\s*[\r\n]+\s*(?P<end>По\s+н\.?\s*в\.?|'+DATE_RE+r')\s*$',re.I)
 for pattern in [split,adjacent]:
  for m in pattern.finditer(text):
   if not any(abs(m.start()-p['pos'])<8 for p in result):result.append({'pos':m.start(),'stop':m.end(),'start':m.group('start'),'end':m.group('end'),'evidence':m.group()})
 return sorted(result,key=lambda p:p['pos'])
def period_at(text,pos):
 periods=employment_periods(text)
 after=[p for p in periods if 0<p['pos']-pos<300]
 if after and re.search(r'(?i)(?:стаж работы|период работы|период трудоустройства|срок работы)\s*:?\s*с?\s*$',text[pos:after[0]['pos']]):return after[0]
 before=[p for p in periods if p['pos']<=pos]
 if before:
  p=before[-1]
  if pos-p['pos']<6500:return p
 after=[p for p in periods if 0<p['pos']-pos<120]
 return after[0] if after else None
def surrounding_period(text,pos):
 closest=period_at(text,pos)
 if not closest:return None,'Дата увольнения не установлена',None
 end=parse_end(closest['end'])
 if end is None:return None,'Работает по настоящее время либо дата окончания не подтверждена',closest['evidence']
 return end,'',closest['evidence']
LEGAL=r'\b(?:ооо|оао|пао|зао|ао|ип|фгуп|муп|гауз|гбуз|гбу|мбу|фбу|фгбу|ано|чоу|гку|мку|чпу|чоп|чоо|пкф|тд|гк)\b'
def company_name(s):
 s=norm(s);s=re.sub(LEGAL,' ',s);return re.sub(r'\s+',' ',s).strip()
SPECIAL_ALIASES={
 'камаз':['kamaz','камский автомобильный завод'],
 'татнефть':['tatneft','татнефть добыча','татнипинефть'],
 'калашников':['kalashnikov','концерн калашников'],
 'нижнекамскшина':['kama tyres','кама тайрс','кама шины'],
 'тэмпо':['темпо','tempo'],
 'таиф':['taif'],
 'сибур холдинг':['сибур','sibur'],
 'танеко':['taneco'],
 'узга':['уральский завод гражданской авиации'],
 'ижевский механический завод':['ижмех'],
 'казанский вертолетный завод':['квз'],
 'производственное объединение елабужский автомобильный завод':['елаз','елабужский автомобильный завод'],
 'хайер индастри рус':['haier','хайер','хаер'],
 'джошкуноз алабуга сoskunoz':['джошкуноз','coskunoz'],
 'кастамону интегрейтед вуд индастри':['кастамону','kastamonu'],
 'хаят кимья':['хаят','hayat'],
 'роквул волга':['роквул','rockwool'],
 'шишеджам аутоматив рус':['шишеджам','sisecam'],
 'сетевая компания елабужские электросети':['елабужские электросети'],
 'ецрб':['елабужская црб','елабужская центральная районная больница'],
 'тн сервис':['тн сервис','тнсервис'],
 'тн алабуга':['тн алабуга','технониколь алабуга']}
SPECIAL_ALIASES.update({'татнефть':['tatneft','татнефть добыча','татнипинефть','альметьевнефть','азнакаевскнефть','лениногорскнефть','джалильнефть','бавлынефть'],
 'нижнекамскнефтехим':['нкнх','nknh'], 'казаньоргсинтез':['казанский органический синтез','kazanorgsintez'],
 'камаз':['kamaz','камский автомобильный завод','камаз мастер'],
 'хаят кимья':['хаят','hayat','хаят кимья алабуга']})
RT=r'татарстан|\bрт\b|казан|елабуг|набережн.{0,10}челн|нижнекам|менделеевск|мамадыш|альметьевск|бугульм|зеленодольск|чистопол|заинск'
LOCAL=r'елабуг|набережн.{0,10}челн|нижнекам|менделеевск|мамадыш'
PUBLIC=r'государствен|муниципальн|\b(?:гауз|гбуз|гбу|гбоу|мбу|мбоу|мбдоу|мкдоу|мбдоу|мадоу|фгбоу|гбпоу|фгбу|гку|мку|мфц)\b|министерств|исполком|администраци[яи] (?:город|район)|налогов.{0,15}инспекц|\bфнс\b'
PRIVATE=r'частн|\b(?:ооо|ип|чоу|ано)\b'
SECURITY=r'\bмвд\b|\bгибдд\b|\bмчс\b|\bфсин\b|\bфсб\b|\bфтс\b|прокуратур|следственн.{0,12}комитет|росгвард|оперуполномоч|участков.{0,20}полиц|инспектор.{0,10}дпс|\bумвд\b|\bуфсин\b|\bуфсб\b'
SECURITY+=r'|\bовд\b|\bувд\b|\bомвд\b|\bгумвд\b|\bфскн\b|\bуфскн\b'
EDU=r'университет|институт|\bшкол[аыуе]\b|детск.{0,8}сад|колледж|техникум|воспитател|учител|преподавател'
MED=r'больниц|поликлиник|диспансер|родильн.{0,8}дом|медицинск.{0,20}(?:центр|учрежден)|медсестр|врач|фельдшер'
OTHER_REGION=r'москв|петербург|ленинградск|свердловск|екатеринбург|пермь|пермск|ижевск|удмурт|уфа\b|башкортостан|башкир|самар|саратов|оренбург|ульяновск|пенза|челябинск|ростов|краснодар|красноярск|новосибирск|томск|омск|тюмен|ханты|сургут|калуг|тула\b|тульск|воронеж|брянск|нижн.{0,8}новгород|дальневост|хабаровск|владивосток|сахалин|дагестан|чуваш|чебоксар|йошкар|марий|белгород|беларус|казахстан|крым|ялт|севастопол|волгоград|волжск'
OTHER_REGION+=r'|карели|усть луга|сатка|центральн.{0,10}федеральн.{0,8}округ|минстрой россии|министерств.{0,120}российской федерации'
ROMAN_MAP=dict(zip('абвгдеёжзийклмнопрстуфхцчшщъыьэюя',['a','b','v','g','d','e','e','zh','z','i','i','k','l','m','n','o','p','r','s','t','u','f','kh','ts','ch','sh','shch','','y','','e','yu','ya']))
def roman(s):return ''.join(ROMAN_MAP.get(ch,ch) for ch in s)
def mention_context(work,pos):
 p=period_at(work,pos);start=p['pos'] if p else max(0,pos-200)
 following=[x['pos'] for x in employment_periods(work) if x['pos']>pos]
 stop=min(following) if following else len(work)
 return work[start:min(stop,max(start+750,pos+300))]
TRAINING_REFERENCE=re.compile(r'(?:развити[еяю].{0,80}навыков|обучени[еяю]\s+(?:персонала|сотрудников|работников))\s*$')
def nonemployment_mention(work,pos):
 prefix=norm(work[max(0,pos-140):pos])
 line=norm(work[work.rfind('\n',0,pos)+1:pos])
 return bool(TRAINING_REFERENCE.search(prefix) or re.search(r'(?:клиент(?:ы|ами|ов|а)?|заказчик(?:и|ами|ов|а)?|для клиентов|дилер|поставки в|партнеры|по заказу|обучение в|курсы в|практика в|оборудование|автомобил(?:я|е|и|ей|ях|ь)|грузовик[аи]?|техник[аи]|шасси|самосвал[аы]?)(?:\s+\w+){0,5}\s*$',prefix)
  or re.search(r'взаимодействие с|представлен.{0,15}интересов в|согласован.{0,30}с\s*$',prefix)
  or re.search(r'работ.{0,20}(?:автомобил|автотранспорт)|ремонт.{0,25}(?:автомобил|техник)|вождение|водитель автомашины',line))
class BanEngine:
 def __init__(self,run=None):
  self.run=run or active();self.meta=read(self.run/'manifest.json');self.as_of=date.fromisoformat(self.meta['as_of'])
  files=list((self.run/'inputs'/'banlist').glob('*.csv'))
  if len(files)<2:raise ValueError('Нужны бан-листы ФИО.csv и Компании.csv')
  people=next((p for p in files if unicodedata.normalize('NFC',p.stem).lower()=='фио'),max(files,key=lambda p:p.stat().st_size))
  companies=next((p for p in files if unicodedata.normalize('NFC',p.stem).lower()=='компании'),min(files,key=lambda p:p.stat().st_size))
  if people==companies:raise ValueError('Не удалось различить бан-листы ФИО и компаний')
  self.people=[];self.phone_index=collections.defaultdict(list);self.surname_index=collections.defaultdict(list);self.ignored_person_rows=[]
  for rowno,r in enumerate(csv.reader(people.read_text('utf-8-sig').splitlines(),delimiter=';'),1):
   if rowno==1 or not r or not r[0].strip():continue
   raw=re.sub(r'(?<=[а-яё])(?=[А-ЯЁ])',' ',r[0])
   name=norm(re.sub(r'\d{1,2}[./]\d{1,2}[./]\d{2,4}','',raw));parts=[p for p in name.split() if not p.isdigit()]
   if re.search(r'проверка интеграции|тест|группа компаний|\bооо\b|\bпао\b|\bоао\b',name):
    self.ignored_person_rows.append({'row':rowno,'value':r[0]})
    contact=phone(r[1]) if len(r)>1 else ''
    if contact:self.phone_index[contact].append({'row':rowno,'name':r[0]})
    continue
   if len(parts)==3 and re.search(r'(?:вич|вна|ична)$',parts[1]):parts=[parts[2],parts[0],parts[1]]
   name=' '.join(parts)
   item={'row':rowno,'name':r[0],'normalized':name,'parts':parts,'phone':phone(r[1]) if len(r)>1 else ''}
   self.people.append(item)
   if item['phone']:self.phone_index[item['phone']].append(item)
   if parts:self.surname_index[parts[0]].append(item)
  self.given_names={p['parts'][1] for p in self.people if len(p['parts'])==3 and re.search(r'(вич|вна|ична)$',p['parts'][2])}
  self.roman_index=collections.defaultdict(list)
  for p in self.people:
   if p['parts']:self.roman_index[roman(p['parts'][0])].append(p)
  geo=self.run/'engine'/'geography'/'ru_cities.json'
  if not geo.exists():geo=Path(__file__).resolve().parent/'geography'/'ru_cities.json'
  self.cities=read(geo)['cities'] if geo.exists() else {}
  self.company_rules=[]
  for rowno,r in enumerate(csv.reader(companies.read_text('utf-8-sig').splitlines(),delimiter=';'),1):
   if rowno==1 or len(r)<2:continue
   stage=r[1].strip();months=None if stage=='100%' else 0 if stage=='Свободно' else int(re.search(r'\d+',stage).group())*(12 if 'год' in stage else 1)
   n=company_name(r[0]);aliases=[n]+SPECIAL_ALIASES.get(n,[])
   if re.search(r'т[эе]м[ -]?по',n):aliases+=['тэмпо','темпо','тэм по','технотрон' if 'технотрон' in n else n]
   if 'калашников' in n:aliases+=['калашников','kalashnikov']
   if 'менделеевский нефтехим' in n:aliases+=['менделеевский нпз']
   if 'квз ремсервис' in n:aliases+=['квз ремсервис']
   # Parentheses often contain a region or generic description, not a company alias.
   # Only explicit curated abbreviations are accepted.
   generic=bool(re.search(r'гос/муниципал|медицина (?:РТ|РФ)|образование (?:РТ|РФ)|МВД,|Росгвардия|МЧС по|Резиденты ОЭЗ|Практика/|Подрядная|Частн|Дилеры',r[0]))
   self.company_rules.append({'row':rowno,'name':r[0],'stage':stage,'months':months,'aliases':list(set(a for a in aliases if a)),'generic':generic})
  for item in self.ignored_person_rows:
   if re.search(r'группа компаний|\bооо\b|\bпао\b|\bоао\b',norm(item['value'])):
    alias=company_name(re.sub(r'(?i)группа компаний','',item['value']))
    self.company_rules.append({'row':'ФИО:'+str(item['row']),'name':item['value'],'stage':'срок не указан','months':None,'aliases':[alias],'generic':False,'term_unknown':True})
  self.company_patterns={}
  for rule in self.company_rules:
   compiled=[]
   for a in sorted(rule['aliases'],key=len,reverse=True):
    suffix=r'(?:а|е|у|ом)?' if a in ['камаз','калашников','танеко','хаят','таиф','кориб','сибур'] else ''
    compiled.append((a,re.compile(r'(?<!\w)'+re.escape(a).replace(r'\ ',r'\s*')+suffix+r'(?!\w)')))
   self.company_patterns[rule['row']]=compiled
 def identities(self,text,source_values=None):
  findings=[];clean=re.sub(r'FILE [^\n]+|PAGE \d+',' ',text)
  clean=re.split(r'(?i)\n\s*рекомендации',clean)[0]
  head=clean[:1800];tail=clean[-1500:]
  nh=norm(head);words=set(nh.split());candidate_names=[]
  for area in [head,tail]:
   nt=norm(area);tokens=nt.split()
   for i,t in enumerate(tokens):
    if re.search(r'(?:вич|вна|ична)$',t) and i>=2 and tokens[i-1] in self.given_names:
     candidate_names.append([tokens[i-2],tokens[i-1],t])
    elif re.search(r'(?:вич|вна|ична)$',t) and i>=1 and i+1<len(tokens) and tokens[i-1] in self.given_names:
     candidate_names.append([tokens[i+1],tokens[i-1],t])
   for line in area.splitlines():
    parts=norm(line).split()
    if len(parts)==2:
     if parts[1] in self.given_names:candidate_names.append(parts)
     elif parts[0] in self.given_names:candidate_names.append([parts[1],parts[0]])
  header_lines=[line.strip() for line in head.splitlines() if line.strip()][:6]
  for line in header_lines:
   line=re.sub(r'(?i)^(?:ф\.?и\.?о\.?|full name|name)\s*:?\s*','',line)
   if not re.fullmatch(r'[A-ZА-ЯЁ][a-zA-Zа-яА-ЯёЁ-]+(?:\s+[A-ZА-ЯЁ][a-zA-Zа-яА-ЯёЁ-]+){1,2}',line):continue
   if re.search(r'(?i)resume|curriculum|vitae|experience|developer|engineer|senior|junior|frontend|backend|software|резюме|личная|желаемая|должность|инженер|специалист|образование|опыт|РАДИОТЕЛЕФОНИСТ',line):continue
   parts=norm(line).split();candidate_names.append(parts)
   if re.search(r'[a-z]',parts[0]):candidate_names.append([parts[-1],*parts[:-1]])
  ph=set(filter(None,[phone(p) for p in re.findall(r'(?:\+?7|8)[\s().-]*(?:\d[\s().-]*){10}',head+'\n'+tail)]))
  # Source table contains placeholders; only real, non-placeholder contact numbers are usable.
  if source_values:
   ph.add(phone(source_values[10] or ''))
   nm=' '.join(str(v) for v in source_values[5:8] if v)
   if nm and source_values[5] and source_values[6]:candidate_names.append(norm(nm).split())
  candidate_names=[list(p) for p in dict.fromkeys(tuple(p) for p in candidate_names)]
  candidate_names=[p for p in candidate_names if not (len(p)==2 and any(len(q)>=3 and q[:2]==p for q in candidate_names))]
  for p in ph:
   for item in self.phone_index.get(p,[]):findings.append({'type':'person','row':item['row'],'rule':item['name'],'basis':'Совпадение контактного телефона, окончание '+p[-4:],'active':True,'certainty':'confirmed'})
  possible=[]
  for cn in candidate_names:
   index=self.roman_index if re.search(r'[a-z]',cn[0]) else self.surname_index
   surnames=fuzzy_process.extract(cn[0],index.keys(),scorer=DamerauLevenshtein.distance,score_cutoff=1 if len(cn[0])>=5 else 0,limit=None)
   for item in [item for surname,_,_ in surnames for item in index[surname]]:
    parts=item['parts']
    if not parts:continue
    original_cn=cn
    if re.search(r'[a-z]',cn[0]):
     parts=[roman(part) for part in parts]
    same_surname=parts[0]==cn[0]
    fuzzy_surname=len(parts[0])>=5 and len(cn[0])>=5 and DamerauLevenshtein.distance(parts[0],cn[0])==1
    if not same_surname and not fuzzy_surname:continue
    if len(parts)==1:
     if same_surname:possible.append((item,cn,'probable','Совпала фамилия с неполной записью бан-листа'))
     continue
    if len(cn)<2:continue
    given_distance=DamerauLevenshtein.distance(parts[1],cn[1])
    if given_distance and parts[1] in self.given_names and cn[1] in self.given_names:continue
    if given_distance>1 or (given_distance and min(len(parts[1]),len(cn[1]))<4):continue
    if fuzzy_surname and given_distance:continue
    if len(parts)>=3 and len(cn)>=3:
     middle_distance=DamerauLevenshtein.distance(parts[2],cn[2])
     if middle_distance>1:continue
    else:middle_distance=None
    certainty='confirmed' if same_surname and given_distance==0 and middle_distance==0 else 'probable'
    basis='Совпадение ФИО' if certainty=='confirmed' else 'Вероятное совпадение ФИО: неполное имя или опечатка'
    possible.append((item,cn,certainty,basis))
  for item,cn,certainty,basis in possible:
   findings.append({'type':'person','row':item['row'],'rule':item['name'],'candidate_name':' '.join(cn),'basis':basis,'active':True,'certainty':certainty})
  return findings,ph,candidate_names
 def assess(self,text,source_values=None):
  text=candidate_content(text)
  findings,phones,candidate_names=self.identities(text,source_values)
  work,structured=work_text(text);nw,original_position=normalized_positions(work)
  word_positions=collections.defaultdict(list)
  for token in re.finditer(r'\b[a-zа-я]{4,}\b',nw):word_positions[token.group()].append(token.start())
  # map normalized tokens back to their surrounding original text for date attribution
  for rule in self.company_rules:
   if rule['generic'] or rule['months']==0:continue
   hits=[]
   for a,pattern in self.company_patterns[rule['row']]:
    for m in pattern.finditer(nw):hits.append((m.start(),a,'confirmed'))
    # One edit in a distinctive company token, restricted to employer header/name context.
    if not hits and len(a)>=5 and ' ' not in a and a not in ['актив','активы','синтез','импульс','основание','фаворит']:
     for word,distance,_ in fuzzy_process.extract(a,word_positions.keys(),scorer=DamerauLevenshtein.distance,score_cutoff=1,limit=None):
      if distance!=1:continue
      for offset in word_positions[word]:
       pos=original_position(offset);p=period_at(work,pos)
       if (p and pos-p['pos']<250) or re.search(LEGAL,norm(work[max(0,pos-20):pos])):hits.append((offset,a,'probable'))
   table_event=None
   if re.search(r'СНИЛС|сведения о трудовой деятельности',text,re.I):
    events=[]
    for offset,a,certainty in hits:
     pos=original_position(offset)
     e=re.search(r'(\d{2}\.\d{2}\.\d{4})\s+(УВОЛЬНЕНИЕ|ПРИЕМ|ПРИЁМ|ПЕРЕВОД)',work[pos:pos+800],re.I)
     if e:
      dt=parse_end(e.group(1))
      if dt:events.append((dt,e.group(2).upper(),e.group()))
    if events:
     latest=max(events,key=lambda e:e[0]);table_event={'date':latest[0] if latest[1]=='УВОЛЬНЕНИЕ' else None,'evidence':latest[2]}
   for offset,a,certainty in hits:
    pos=original_position(offset);context=mention_context(work,pos)
    if nonemployment_mention(work,pos):continue
    if isinstance(rule['row'],int) and 89<=rule['row']<=238 and re.search(OTHER_REGION,norm(context)) and not re.search(RT,norm(context)):continue
    if a in ['основание','активы','импульс','зис','синтез','фаворит'] and not re.search(r'\b(?:ООО|АО|ОАО|ПАО|ИП)\s*[«"\s]*'+re.escape(a),context,re.I):continue
    if re.search(r'представитель бренда|работа на.*объекте|проектная деятельность',norm(context)):certainty='probable'
    findings.append(self.company_finding(rule,work,pos,context,certainty,table_event))
  # Sector-wide rules operate only on employment evidence, never on a degree or relocation wish.
  for m in re.finditer(SECURITY+'|'+PUBLIC+'|'+EDU+'|'+MED,nw):
   pos=original_position(m.start());ctx=norm(mention_context(work,pos))
   if nonemployment_mention(work,pos):continue
   tail=nw[m.start():m.start()+130]
   if re.match(r'(?:государствен|муниципальн)[а-я]*\s+(?:норматив|стандарт|контракт|закуп|пошлин|реестр|заказ|экзамен|аттестац|аккредитац|акт|экспертиз|регистрац)',tail):continue
   period=period_at(work,pos)
   if period and pos-period['pos']>400 and m.group() in ['государствен','муниципальн'] and re.search(PRIVATE,norm(work[period['pos']:period['pos']+250])) and not re.search(r'работал.{0,30}$|работала.{0,30}$',nw[max(0,m.start()-80):m.start()]):continue
   if re.search(r'курсы|повышение квалификации|обучение в|закончил|окончил',ctx) and not re.search(r'должност|работал|работала|сотрудник|обязанност',ctx):continue
   sec=bool(re.search(SECURITY,m.group()));pub=bool(re.search(PUBLIC,ctx));edu=bool(re.search(EDU,m.group()));med=bool(re.search(MED,m.group()))
   private=bool(re.search(PRIVATE,ctx));rt=bool(re.search(RT,ctx));local=bool(re.search(LOCAL,ctx))
   regions=set();words=ctx[:700].split()
   for size in range(1,5):
    for i in range(len(words)-size+1):regions.update(self.cities.get(' '.join(words[i:i+size]),[]))
   if regions=={'Tatarstan Republic'}:rt=True
   known_elsewhere=bool(regions) and 'Tatarstan Republic' not in regions
   if re.search(r'министерств|университет|институт',m.group()) and re.search(r'реконструкц|заказчик|фасадов|награды|награжден',ctx) and not re.search(r'должност|работал|работала|прием на работу|уволен',ctx):continue
   wanted=[]
   if sec:
    wanted=[2,15 if local else 239] if rt else [88]
   elif pub and rt:wanted=[2]
   elif (edu or med) and pub:wanted=[12 if edu and rt else 91 if med and rt else 264 if edu else 263]
   elif (edu or med) and not private:wanted=[12 if edu and rt else 91 if med and rt else 264 if edu else 263]
   region_unknown=(pub or sec) and not rt and not known_elsewhere and not re.search(OTHER_REGION,ctx)
   if region_unknown:wanted.append(2)
   if not wanted:continue
   for rowno in wanted:
    rule=next((r for r in self.company_rules if r['row']==rowno),None)
    if rule:
     ambiguous=region_unknown or not structured or ((edu or med) and not pub) or bool(re.search(r'проектная деятельность|учебный проект|производственная практика|стажиров',ctx))
     finding=self.company_finding(rule,work,pos,mention_context(work,pos),'probable' if ambiguous else 'confirmed')
     if region_unknown:finding['date_note']+='; регион работодателя не установлен — уточнить применимость запрета РТ'
     if (edu or med) and not pub:finding['date_note']+='; уточнить государственную или частную форму работодателя'
     if re.search(r'проектная деятельность|учебный проект|производственная практика|стажиров',ctx):finding['date_note']+='; уточнить трудоустройство или учебную/проектную практику'
     findings.append(finding)
  # Missing identity/work evidence cannot establish clearance. It is reported as uncertainty, not an actual blacklist match.
  body=norm(text)
  has_name=bool(candidate_names)
  if source_values and source_values[5] and source_values[6]:has_name=True
  # A two-part header with a contact can be checked against the partial-name index.
  if not has_name and phones and len(re.findall(r'\b[А-ЯЁ][а-яё]{2,}\b',text[:700]))>=2:has_name=True
  if not has_name:findings.append({'type':'uncertainty','rule':'Проверка личности','basis':'ФИО кандидата не удалось уверенно установить; требуется уточнение','active':True,'certainty':'probable'})
  if len(re.sub(r'\W','',work))<120:findings.append({'type':'uncertainty','rule':'Полнота резюме','basis':'Недостаточно сведений для проверки трудовой истории','active':True,'certainty':'probable'})
  if not re.search(r'опыт.{0,20}(?:отсутствует|нет)|без опыта|нигде не работ|не работал|первая работа',body) and not employment_periods(work) and not re.search(r'работал|работала|работает|работаю|стаж|\bооо\b|\bоао\b|\bпао\b|трудовая деятельность',norm(work)):
   findings.append({'type':'uncertainty','rule':'Полнота трудовой истории','basis':'Работодатели и периоды работы не установлены; бан до уточнения трудовой истории','active':True,'certainty':'probable'})
  unique={json.dumps({k:f.get(k) for k in ['type','rule','stage','end_date','certainty']},sort_keys=True,ensure_ascii=False):f for f in findings}
  findings=list(unique.values());active_hits=[f for f in findings if f['active']]
  active_hits.sort(key=lambda f:(f['type']!='person',f.get('months') is not None,f['certainty']!='confirmed'))
  display_hits=[];display_seen=set()
  for f in active_hits:
   key=(f['type'],company_name(f['rule']),f.get('stage'))
   if key not in display_seen:display_hits.append(f);display_seen.add(key)
  def brief(f):
   if f['type']=='person':return f["basis"]+': '+f['rule']
   if f['type']=='uncertainty':return f['basis']
   prefix='Вероятный бан: ' if f['certainty']=='probable' else ''
   return prefix+f['rule']+' — '+f['stage']+('; '+f['date_note'] if f.get('date_note') else '')
  comment='; '.join(brief(f) for f in display_hits[:3])
  if len(display_hits)>3:comment+='; другие основания — в журнале'
  if not active_hits and findings:
   expired=[f for f in findings if not f['active']]
   comment='Срок бана истёк: '+'; '.join(f['rule']+' до '+f.get('expires','') for f in expired[:2])
  return {'ban':'да' if active_hits else 'нет','ban_comment':comment,'ban_findings':findings,'identity_checked':has_name,'candidate_names':candidate_names,'employment_section_detected':structured}
 def company_finding(self,rule,work,pos,context,certainty,table_event=None):
  end,note,period=surrounding_period(work,pos)
  if table_event:end=table_event['date'];period=table_event['evidence'];note='Последнее кадровое событие: '+period
  months=rule['months'];expires=add_months(end,months) if end and months else None
  active=months is None or end is None or expires>self.as_of
  if months is None:note='бессрочный запрет'
  elif expires:note='срок до '+expires.strftime('%d.%m.%Y')
  else:certainty='probable'
  if rule.get('term_unknown'):note='Компания указана в списке ФИО без срока; бан до уточнения';certainty='probable'
  return {'type':'company','row':rule['row'],'rule':rule['name'],'stage':rule['stage'],'months':months,'active':active,'certainty':certainty,
   'end_date':end.isoformat() if end else None,'expires':expires.isoformat() if expires else None,'date_note':note,'period_evidence':period,'context':context.strip(),'matched_text':work[pos:pos+110]}
