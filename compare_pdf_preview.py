"""Compare character placement; generated report contains metrics, not document text."""
import difflib,json,statistics,sys
from pathlib import Path
import pdfplumber
pdf,preview,out=map(Path,sys.argv[1:4])
p=json.loads(preview.read_text(encoding='utf-8'))
with pdfplumber.open(pdf) as d:
 page=d.pages[0];height=float(page.height)
 a=[dict(c,text=letter,unambiguous=len(c['text'])==1) for c in page.chars for letter in c['text'] if letter.isalnum()]
 # Match reading order for body followed by footer, rather than PDF stream order.
 a=[c for c in a if c['top']<height-55]+[c for c in a if c['top']>=height-55]
 b=[c for c in p['chars'] if len(c['text'])==1 and c['text'].isalnum()]
 matcher=difflib.SequenceMatcher(None,''.join(c['text'] for c in a),''.join(c['text'] for c in b),autojunk=False)
 pairs=[]
 for block in matcher.get_matching_blocks():
  for i in range(block.size):
   x,y=a[block.a+i],b[block.b+i]
   if not x['unambiguous']:continue
   pairs.append({'paragraph':y.get('paragraph'), 'dx':y['x']-x['x0'], 'dy':y['baseline']-(height-x['matrix'][5])})
 def stats(values):
  v=sorted(abs(x) for x in values)
  return {'median':statistics.median(v),'p90':v[int((len(v)-1)*.9)],'max':max(v)}
 groups={}
 for pair in pairs:groups.setdefault(str(pair['paragraph']),[]).append(pair)
 report={'pdf_pages':len(d.pages),'preview_page':p['page'],'matched_characters':len(pairs),
  'horizontal_error_pt':stats(v['dx'] for v in pairs),'baseline_error_pt':stats(v['dy'] for v in pairs),
  'paragraph_errors':{k:{'horizontal':stats(x['dx'] for x in v),'baseline':stats(x['dy'] for x in v)} for k,v in groups.items()}}
 out.write_text(json.dumps(report,indent=2),encoding='utf-8')
 print(json.dumps({k:v for k,v in report.items() if k!='paragraph_errors'},indent=2))
