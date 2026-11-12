"""Attach verbatim source windows to reviewed findings; never infer employment."""
import re,math
from .ban_engine import norm,company_name,normalized_positions

STOP=set('государственное федеральное учреждение бюджетное автономное общество ограниченной ответственностью компания предприятие филиал имени российской федерации настоящее время название указано'.split())
def source_window(text,label,period=None):
 normalized,position=normalized_positions(text)
 tokens=[w for w in company_name(label).split() if len(w)>2 and w not in STOP]
 candidates=[]
 for token in tokens:
  for m in re.finditer(r'\b'+re.escape(token)+r'\b',normalized):
   left=max(0,m.start()-150);right=min(len(normalized),m.end()+500)
   window=normalized[left:right]
   score=sum((1+math.log(len(w)))/(1+math.log(1+len(re.findall(r'\b'+re.escape(w)+r'\b',normalized)))) for w in set(tokens) if re.search(r'\b'+re.escape(w)+r'\b',window))
   if period:
    score+=sum(0.4 for d in period if d and d[:4].isdigit() and d[:4] in window)
   candidates.append((score,-m.start(),m.start()))
 if not candidates:return None
 anchor=position(max(candidates)[2]);start=max(0,text.rfind('\n',0,max(0,anchor-180)))
 end=text.find('\n',min(len(text),anchor+1000));end=len(text) if end<0 else end
 filemarks=list(re.finditer(r'(?m)^FILE (.+)$',text[:anchor+1]))
 pagemarks=list(re.finditer(r'(?m)^PAGE (\d+)',text[:anchor+1]))
 return dict(source_excerpt=text[start:end],excerpt_char_start=start,excerpt_char_end=end,source_file=filemarks[-1].group(1).strip() if filemarks else None,source_page=int(pagemarks[-1].group(1)) if pagemarks else None,quote_method='verbatim_window_around_reviewed_employer_or_candidate',quote_anchor=anchor)

def attach(ban,card,text):
 for f in ban['findings']:
  original_label=f.get('employer') or f.get('candidate_name') or card.get('name','')
  label=f.get('source_cv_anchor') or card.get('ban_quote_anchors',{}).get(original_label,original_label)
  evidence=source_window(text,label,f.get('period'))
  if evidence:
   f.update(evidence)
   assert text[f['excerpt_char_start']:f['excerpt_char_end']]==f['source_excerpt']
  elif original_label in card.get('source_row_evidence',{}):
   f.update(source_excerpt=card['source_row_evidence'][original_label],source_file='Матка №1.xlsx',source_page=None,quote_method='verbatim_source_row_cell',quote_anchor=None)
  else:
   f['quote_issue']='Нужна ручная привязка к фрагменту текста'
 return ban
