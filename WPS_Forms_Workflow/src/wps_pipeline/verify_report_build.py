"""Independent read-only verifier of frozen-source to assembled-output content.

Does not import the assembler. Replays declared narrow text operations, then
compares ordered math/media/link/bookmark/table content and package invariants.
"""
from pathlib import Path
from zipfile import ZipFile
from copy import deepcopy
from collections import Counter
import json,hashlib,posixpath,re,argparse
from lxml import etree as E
from .verify_number_edits import heading_edit_error

N={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main','m':'http://schemas.openxmlformats.org/officeDocument/2006/math','wp':'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing','a':'http://schemas.openxmlformats.org/drawingml/2006/main','r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
N.update(v='urn:schemas-microsoft-com:vml',o='urn:schemas-microsoft-com:office:office',pic='http://schemas.openxmlformats.org/drawingml/2006/picture')
def checked_pictures(root):
 result=[]
 for n in root.iter():
  if n.tag=='{'+N['w']+'}drawing' and n.xpath('.//a:blip[@r:embed or @r:link]',namespaces=N):result.append(n)
  elif n.tag=='{'+N['w']+'}pict':
   equation=any(x.get('equationxml') is not None for x in n.iter())
   inside_object=any(a.tag=='{'+N['w']+'}object' for a in n.iterancestors())
   if not equation and not inside_object and n.findall('.//v:imagedata',N):result.append(n)
 return result

def q(k):a,b=k.split(':');return '{'+N[a]+'}'+b
def digest(b):return hashlib.sha256(b).hexdigest()
def tree(n):return [n.tag,sorted(n.attrib.items()),n.text if not len(n) else None,[tree(c) for c in n]]
def ordinary(n):return ''.join(n.xpath('.//w:t[not(ancestor::m:oMath)]/text()',namespaces=N))

def checked_equation_reference(body,number,excluded):
 if not isinstance(number,str) or not number.isdigit():return False
 token=re.escape(number)
 patterns=[r'(?:公式|(?<![格方模样形制程公])式)\s*[（(]\s*'+token+r'\s*[）)]',
           r'(?:参见|见|由|根据|代入|利用|按|如)\s*(?:公式|式)\s*'+token+r'(?![\d.eE×xX*·])',
           r'(?:^|[\s，。；;：:])式\s*'+token+r'(?![\d.eE×xX*·])']
 for root in body:
  if root is excluded:continue
  value=ordinary(root)
  for pattern in patterns:
   if any(not re.match(r'\s*(?:[×xX*·]\s*10|[eE][+\-]?\d)',value[m.end():]) for m in re.finditer(pattern,value)):return True
 return False

class Package:
 def __init__(self,path):
  self.path=Path(path)
  with ZipFile(path) as z:self.parts={n:z.read(n) for n in z.namelist()}
  self.root=E.fromstring(self.parts['word/document.xml']);self.body=self.root.find('w:body',N);self.rels={x.get('Id'):dict(x.attrib) for x in E.fromstring(self.parts['word/_rels/document.xml.rels'])}
 def relationship(self,rid):
  r=self.rels[rid];target=r['Target']
  if r.get('TargetMode')=='External':return [r['Type'],'External',target]
  path=target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join('word',target))
  return [r['Type'],'Internal',digest(self.parts[path])]
 def content(self,n,bmids=None,bmnames=None):
  bmids=bmids or {};bmnames=bmnames or {}
  if n.tag==q('w:tbl'):
   rows=[]
   for tr in n.findall('w:tr',N):
    cells=[]
    for tc in tr.findall('w:tc',N):
     span=tc.find('w:tcPr/w:gridSpan',N);vm=tc.find('w:tcPr/w:vMerge',N)
     cells.append([int(span.get(q('w:val'))) if span is not None else 1,None if vm is None else vm.get(q('w:val'),'continue'),[self.content(c,bmids,bmnames) for c in tc if c.tag in (q('w:p'),q('w:tbl'))]])
    rows.append(cells)
   return ['table',len(n.findall('w:tblGrid/w:gridCol',N)),rows]
  events=[]
  def push(k,v):
   if k=='text' and events and events[-1][0]=='text':events[-1][1]+=v
   else:events.append([k,v])
  def walk(x):
   local=E.QName(x).localname;space=E.QName(x).namespace
   if space==N['w'] and local in ('pPr','rPr'):return
   if space==N['m'] and local in ('oMath','oMathPara'):
    m=deepcopy(x)
    for col in m.findall('.//w:color',N):col.attrib.clear();col.set(q('w:val'),'000000')
    push(local,tree(m));return
   if space==N['wp'] and local in ('inline','anchor'):
    push('drawing',[local,[self.relationship(v) for v in x.xpath('.//a:blip/@r:embed',namespaces=N)]]);return
   if x.tag in (q('w:pict'),q('w:object')):
    raw=deepcopy(x)
    for prop in list(raw.xpath('.//w:rPr|.//w:pPr',namespaces=N)):prop.getparent().remove(prop)
    for n in raw.iter():
     for key,rid in list(n.attrib.items()):
      if key.startswith('{'+N['r']+'}'):n.set(key,json.dumps(self.relationship(rid),ensure_ascii=False,sort_keys=True))
    push('legacy',tree(raw));return
   if x.tag==q('w:t'):push('text',x.text or '');return
   if x.tag in (q('w:bookmarkStart'),q('w:bookmarkEnd')):
    attrs=dict(x.attrib);old=attrs.get(q('w:id'));attrs[q('w:id')]=bmids.get(old,old)
    if q('w:name') in attrs:attrs[q('w:name')]=bmnames.get(attrs[q('w:name')],attrs[q('w:name')])
    push(local,sorted(attrs.items()));return
   if x.tag==q('w:hyperlink'):
    attrs=dict(x.attrib);rid=attrs.pop(q('r:id'),None)
    if q('w:anchor') in attrs:attrs[q('w:anchor')]=bmnames.get(attrs[q('w:anchor')],attrs[q('w:anchor')])
    push('link_start',[sorted(attrs.items()),self.relationship(rid) if rid else None])
    for c in x:walk(c)
    push('link_end',None);return
   if x.tag==q('w:fldSimple'):
    copy=deepcopy(x)
    for prop in copy.findall('.//w:rPr',N):prop.getparent().remove(prop)
    push(local,tree(copy));return
   if space==N['w'] and local in ('tab','br','cr','noBreakHyphen','softHyphen','fldChar','instrText'):push(local,tree(x));return
   for c in x:walk(c)
  if n.tag==q('w:p'):
   for x in n:walk(x)
  else:walk(n)
  return [E.QName(n).localname,events]

def checked_fields(node):
 answer=[]
 for p in ([node] if node.tag==q('w:p') else node.findall('.//w:p',N)):
  start=None;middle=None;items=[]
  for i,child in enumerate(p):
   if child.tag==q('w:fldSimple'):items.append((i,dict(simple=child,first=child,last=child,codes=[],texts=child.findall('.//w:t',N))))
   flags=child.findall('w:fldChar',N) if child.tag==q('w:r') else []
   for flag in flags:
    kind=flag.get(q('w:fldCharType'))
    if kind=='begin':start=i;middle=None
    elif kind=='separate':middle=i
    elif kind=='end' and start is not None and middle is not None:
     span=list(p);items.append((start,dict(simple=None,first=span[start],last=child,codes=[t for c in span[start:middle] for t in c.findall('w:instrText',N)],texts=[t for c in span[middle:i+1] for t in c.findall('w:t',N)])));start=None
  for _,f in sorted(items,key=lambda x:x[0]):
   f['p']=p;f['instruction']=f['simple'].get(q('w:instr'),'') if f['simple'] is not None else ''.join(t.text or '' for t in f['codes']);f['cache']=''.join(t.text or '' for t in f['texts']);answer.append(f)
 return answer


def checked_caption_prefix(text):
 # Independent source lexer: do not import the recognizer/writer span.
 return re.match(r'^\s*(?P<label>[图表])\s*(?P<number>\d+(?:[.．\-－–—]\d+)*)(?:[.．、:：]?[^\S\r\n]*)',text)


def checked_field_span(p,field):
 positions={};offset=0
 for t in p.xpath('.//w:t[not(ancestor::m:oMath)]',namespaces=N):
  positions[t]=(offset,offset+len(t.text or ''));offset+=len(t.text or '')
 if not field['texts'] or any(t not in positions for t in field['texts']):raise ValueError('Caption field is outside ordinary text')
 a,b=positions[field['texts'][0]][0],positions[field['texts'][-1]][1]
 if ordinary(p)[a:b]!=field['cache']:raise ValueError('Caption field result is not a contiguous span')
 return a,b


def replay(node,change,converted_nodes=None):
 if change['operation']=='unresolved_numbering_notice':
  paragraphs=list(node.iter(q('w:p')))
  if change['prefix']!='[原自动编号待核] ' or type(change['paragraph_index']) is not int or not 0<=change['paragraph_index']<len(paragraphs):raise ValueError('Invalid local numbering notice')
  p=paragraphs[change['paragraph_index']];run=E.Element(q('w:r'));text=E.SubElement(run,q('w:t'));text.text=change['prefix'];text.set('{http://www.w3.org/XML/1998/namespace}space','preserve');p.insert(next((i for i,n in enumerate(p) if n.tag!=q('w:pPr')),len(p)),run)
  return
 if change['operation']=='normalize_embedded_story_styles':
  paragraphs=[p for p in node.findall('.//w:p',N) if any(n.tag in (q('w:pict'),q('w:object'),q('w:txbxContent')) for n in p.iterancestors())]
  if len(paragraphs)!=len(change['paragraphs']):raise ValueError('Embedded paragraph normalization count mismatch')
  for p,record in zip(paragraphs,change['paragraphs']):
   st=p.find('w:pPr/w:pStyle',N);old=st.get(q('w:val')) if st is not None else None
   if old!=record['source_style_id'] or record['target_style_id']!=change['target_style_id']:raise ValueError('Embedded style source binding mismatch')
  return
 if change['operation']=='detach_unmatched_bookmark':
  marks=[n for n in node.iter() if n.tag in (q('w:bookmarkStart'),q('w:bookmarkEnd')) and n.get(q('w:id'))==change['bookmark_id']]
  if len(marks)!=change['count'] or (node in marks)!=change['root_marker']:raise ValueError('Bookmark marker detachment does not match source')
  for mark in marks:
   if mark is node:node.tag=q('w:p');node.attrib.clear()
   else:
    parent=mark.getparent();at=parent.index(mark)
    for child in list(mark):parent.insert(at,child);at+=1
    parent.remove(mark)
  return
 if change['operation']=='missing_link_placeholder_as_embed':
  nodes=node.xpath(change['path'],namespaces=N)
  if len(nodes)!=1 or nodes[0].tag!=q('a:blip') or nodes[0].get(q('r:link'))!=change['relationship_id'] or nodes[0].get(q('r:embed')):raise ValueError('Invalid missing external image placeholder')
  nodes[0].set(q('r:embed'),nodes[0].get(q('r:link')));nodes[0].attrib.pop(q('r:link'));return
 if change['operation']=='drop_image_link':
  nodes=node.xpath(change['path'],namespaces=N)
  if len(nodes)!=1 or nodes[0].tag!=q('a:blip') or nodes[0].get(q('r:link'))!=change['relationship_id'] or not nodes[0].get(q('r:embed')):raise ValueError('Invalid image link removal')
  nodes[0].attrib.pop(q('r:link'));return
 if change['operation']=='unlink_excluded_frontmatter_anchor':
  links=[n for n in node.findall('.//w:hyperlink',N) if n.get(q('w:anchor'))==change['anchor']]
  if not links:raise ValueError('Missing excluded navigation link')
  link=links[-1];parent=link.getparent();at=parent.index(link)
  for child in list(link):parent.insert(at,child);at+=1
  parent.remove(link);return
 if change['operation'] in ('freeze_unbound_caption_seq','freeze_excluded_frontmatter_reference','freeze_caption_chapter_field','freeze_unresolved_reference'):
  f=checked_fields(node)[change['field_index']]
  if change['operation']=='freeze_unbound_caption_seq':allowed=bool(re.fullmatch(r'\s*SEQ\s+\w+(?:\s+\\[rs]\s+\d+|\s+\\\*\s+(?:ARABIC|MERGEFORMAT))*\s*',f['instruction'],re.I))
  elif change['operation']=='freeze_caption_chapter_field':
   prefix=checked_caption_prefix(ordinary(node));a,b=checked_field_span(node,f)
   allowed=(prefix is not None and prefix.start('number')<=a<b<=prefix.end('number') and
            bool(re.fullmatch(r'\d+(?:[.．\-－–—]\d+)*',f['cache'])) and
            bool(re.fullmatch(r'\s*STYLEREF\s+(?:"[^"\r\n]+"|[^\s"\\]+)(?:\s+\\[nrswtl])*\s*',f['instruction'],re.I)))
  elif change['operation']=='freeze_unresolved_reference':allowed=bool(re.match(r'\s*(?:REF|PAGEREF|NOTEREF)\b',f['instruction'],re.I))
  else:allowed=bool(re.match(r'\s*(?:REF|PAGEREF|NOTEREF)\s+'+re.escape(change['bookmark'])+r'(?=\s|$)',f['instruction'],re.I))
  if f['instruction']!=change['instruction'] or f['cache']!=change['old'] or change['new']!=change['old'] or not allowed:raise ValueError('Invalid field display preservation')
  p=f['p'];start=p.index(f['first']);end=p.index(f['last']);runs=[]
  if f['simple'] is not None:runs=[deepcopy(x) for x in f['simple']]
  else:
   display=True
   for original in list(p)[start:end+1]:
    copy=deepcopy(original)
    if original.tag!=q('w:r'):
     if display:runs.append(copy)
     continue
    for child in list(copy):
     if child.tag==q('w:fldChar'):
      state=child.get(q('w:fldCharType'));display=state!='begin';copy.remove(child)
     elif child.tag!=q('w:rPr') and not display:copy.remove(child)
    if any(x.tag!=q('w:rPr') for x in copy):runs.append(copy)
  for child in list(p)[start:end+1]:p.remove(child)
  for offset,run in enumerate(runs):p.insert(start+offset,run)
  if converted_nodes is not None and change['operation'] in ('freeze_excluded_frontmatter_reference','freeze_unresolved_reference'):converted_nodes.update(t for run in runs for t in run.findall('.//w:t',N))
  return
 if change['operation'] in ('native_caption_seq','native_caption_ref','remap_native_reference_field'):
  f=checked_fields(node)[change['field_index']]
  if f['instruction']!=change['instruction'] or f['cache']!=change['old']:raise ValueError('Native field source binding changed')
  new=change['new_instruction']
  if change['operation']=='native_caption_seq':
   if not re.fullmatch(r'\s*SEQ\s+\w+(?:\s+\\[rs]\s+\d+|\s+\\\*\s+(?:ARABIC|MERGEFORMAT))*\s*',f['instruction'],re.I) or not re.fullmatch(r' SEQ [图表] \\\* ARABIC ',new) or not change['new'].isdigit():raise ValueError('Invalid native SEQ migration')
  else:
   old=re.fullmatch(r'(\s*(?:REF|PAGEREF|NOTEREF)\s+)(\w+)(.*)',f['instruction'],re.I|re.S);target=re.fullmatch(r'(\s*(?:REF|PAGEREF|NOTEREF)\s+)(\w+)(.*)',new,re.I|re.S)
   if not old or not target or old[1].strip()!=target[1].strip() or old[3].strip()!=target[3].strip():raise ValueError('Reference command or switches changed')
  if f['simple'] is not None:f['simple'].set(q('w:instr'),new);f['simple'].set(q('w:dirty'),'1')
  else:
   f['codes'][0].text=new
   for t in f['codes'][1:]:t.text=''
   f['first'].find('w:fldChar',N).set(q('w:dirty'),'1')
  f['texts'][0].text=change['new']
  for t in f['texts'][1:]:t.text=''
  return
 if change['operation'] in ('promote_caption_number','promote_caption_reference'):
  p=([node] if node.tag==q('w:p') else node.findall('.//w:p',N))[change['paragraph_index']];a,b=change['start'],change['end'];offset=0
  if ordinary(p)[a:b]!=change['new'] or not change['new'].isdigit():raise ValueError('Field promotion may consume only declared digits')
  expression=r' SEQ [图表] \\\* ARABIC ' if change['operation']=='promote_caption_number' else r' REF \w+ \\h '
  if not re.fullmatch(expression,change['instruction']):raise ValueError('Invalid generated field instruction')
  for t in p.xpath('.//w:t[not(ancestor::m:oMath)]',namespaces=N):
   value=t.text or '';end=offset+len(value)
   if offset<=a<end:
    run=t.getparent()
    if b>end or run.getparent() is not p or len(run.findall('w:t',N))!=1:raise ValueError('Unsafe source field span')
    at=p.index(run);p.remove(run)
    for piece in ('prefix','field','suffix'):
     if piece=='field':
      f=E.Element(q('w:fldSimple'));f.set(q('w:instr'),change['instruction']);f.set(q('w:dirty'),'1');r=E.SubElement(f,q('w:r'));rp=run.find('w:rPr',N)
      if rp is not None:r.append(deepcopy(rp))
      E.SubElement(r,q('w:t')).text=change['new'];p.insert(at,f);at+=1
     else:
      text=value[:a-offset] if piece=='prefix' else value[b-offset:]
      if text:r=deepcopy(run);r.find('w:t',N).text=text;p.insert(at,r);at+=1
    return
   offset=end
  raise ValueError('Missing generated field digits')
 if change['operation']=='add_caption_number_bookmark':
  fs=[f for f in checked_fields(node) if re.match(r'\s*SEQ\b',f['instruction'],re.I)]
  if len(fs)!=1:raise ValueError('Bookmark must enclose one native SEQ')
  f=fs[0];p=f['p'];a=p.index(f['first']);b=p.index(f['last']);s=E.Element(q('w:bookmarkStart'));s.set(q('w:id'),change['id']);s.set(q('w:name'),change['name']);e=E.Element(q('w:bookmarkEnd'));e.set(q('w:id'),change['id']);p.insert(b+1,e);p.insert(a,s);return
 if change['operation'] in ('materialize_caption_seq','materialize_caption_ref','remap_simple_reference_field'):
  fields=node.findall('.//w:fldSimple',N);field=fields[change['field_index']]
  instruction=field.get(q('w:instr'),'')
  if instruction!=change['instruction']:raise ValueError('Field instruction fingerprint mismatch')
  if change['operation']=='remap_simple_reference_field':
   old=re.fullmatch(r'(\s*(?:REF|PAGEREF|NOTEREF)\s+)([\w]+)(.*)',instruction,re.I|re.S)
   new=re.fullmatch(r'(\s*(?:REF|PAGEREF|NOTEREF)\s+)([\w]+)(.*)',change['new_instruction'],re.I|re.S)
   if not old or not new or old[1]!=new[1] or old[3]!=new[3]:raise ValueError('Reference rename changed field behavior')
   field.set(q('w:instr'),change['new_instruction']);return
  if ordinary(field)!=change['old'] or not re.fullmatch(r'\d+',change['new']):raise ValueError('Invalid field result conversion')
  if change['operation']=='materialize_caption_seq':
   if not re.fullmatch(r'\s*SEQ\s+\w+\s*(?:\\r\s+\d+\s*)?(?:\\\*\s+(?:ARABIC|MERGEFORMAT)\s*)*',instruction,re.I):raise ValueError('Unsupported SEQ')
  else:
   m=re.fullmatch(r'\s*REF\s+(\w+)\s*((?:\\h\s*|\\\*\s+MERGEFORMAT\s*)*)',instruction,re.I)
   if not m or m[1]!=change['bookmark']:raise ValueError('Unsupported REF')
   if bool(change.get('anchor'))!=bool(re.search(r'\\h\b',instruction,re.I)):raise ValueError('Reference navigation changed')
  texts=field.findall('.//w:t',N)
  if not texts:raise ValueError('Missing field result')
  if converted_nodes is not None:converted_nodes.update(texts)
  texts[0].text=change['new']
  for t in texts[1:]:t.text=''
  parent=field.getparent();at=parent.index(field);children=list(field)
  if change.get('anchor'):
   wrapper=E.Element(q('w:hyperlink'));wrapper.set(q('w:anchor'),change['anchor'])
   for child in children:wrapper.append(child)
   parent.insert(at,wrapper)
  else:
   for i,child in enumerate(children):parent.insert(at+i,child)
  parent.remove(field);return

 if change['operation'] in ('detach_automatic_heading_numbering','preserve_automatic_list_numbering'):
  if change.get('deleted_text_characters')!=0:raise ValueError('Numbering-property change must not delete text')
  return
 if change['operation']=='remove_standalone_page_break':
  if node.tag!=q('w:p') or ordinary(node).strip() or node.xpath('.//m:oMath|.//w:drawing|.//w:pict',namespaces=N):raise ValueError('Page break removal on content paragraph')
  brs=node.xpath('.//w:br[@w:type="page"]',namespaces=N)
  if len(brs)!=change['count']:raise ValueError('Page break count mismatch')
  for br in brs:br.getparent().remove(br)
  return
 if change['operation']=='promote_standalone_math':
  ms=node.findall('m:oMath',N)
  if ordinary(node).strip() or len(ms)!=change['count'] or node.findall('m:oMathPara',N):raise ValueError('Invalid standalone math promotion')
  at=node.index(ms[0]);mp=E.Element(q('m:oMathPara'));pr=E.SubElement(mp,q('m:oMathParaPr'));jc=E.SubElement(pr,q('m:jc'));jc.set(q('m:val'),'center')
  for m in ms:node.remove(m);mp.append(m)
  node.insert(at,mp);return
 ps=[node] if node.tag==q('w:p') else node.findall('.//w:p',N);p=ps[change['paragraph_index']]
 text=ordinary(p);a=change['start'];b=change['end'];old=change['old'];new=change['new']
 if text[a:b]!=old:raise ValueError('Declared source text does not match')
 if converted_nodes:
  cursor=0
  for t in p.xpath('.//w:t[not(ancestor::m:oMath)]',namespaces=N):
   end=cursor+len(t.text or '')
   if t in converted_nodes and a<end and b>cursor:raise ValueError('Text edit reinterprets an already materialized field number')
   cursor=end
 kind=change['operation']
 if kind=='source_title_presentation':
  value=old.strip()
  if value.startswith('【') and value.endswith('】'):value=value[1:-1].strip()
  if value!=new or a!=0 or b!=len(text):raise ValueError('Title normalization must preserve its exact words')
 if kind=='source_display_title' and (a!=0 or b!=len(text) or not new or len(new)>200):raise ValueError('Invalid explicit display title')
 if kind=='mapped_object_reference':
  refs=list(re.finditer(r'(?<![子附])(?P<label>[图表])\s*(?P<number>\d+(?:[.．\-－–—]\d+)*)(?!\d|[.．\-－–—]\d)',text))
  if not new.isdigit() or not any(m.span('number')==(a,b) for m in refs):raise ValueError('Reference operation splits its complete source ordinal')
 if kind in ('heading_number','source_group_number') and not (re.fullmatch(r'\d+(?:\.\d+)* ',new) or (new=='' and re.fullmatch(r'\d+(?:\.\d+)*',change.get('native_number','')))):raise ValueError('Invalid heading number')
 if kind=='manual_list_prefix':
  from .verify_number_edits import list_edit_error
  error=list_edit_error(text,change)
  if error:raise ValueError(error)
 if kind=='caption_number':
  prefix=checked_caption_prefix(text)
  if not re.fullmatch(r'[图表] \d+　',new) or a!=0 or b!=(prefix.end() if prefix else 0):raise ValueError('Caption edit must consume exactly the old label/ordinal/separator')
 if kind in ('caption_text_number','caption_text_label','caption_text_separator'):
  prefix=checked_caption_prefix(text)
  if prefix is None or checked_fields(node):raise ValueError('Plain caption edit requires an ordinary source ordinal')
  if kind=='caption_text_number':valid=(a,b)==prefix.span('number') and new.isdigit()
  elif kind=='caption_text_label':valid=a==0 and b==prefix.start('number') and new in ('图 ','表 ')
  else:valid=a==prefix.end('number') and b==prefix.end() and new=='　'
  if not valid:raise ValueError('Plain caption edit exceeds its exact source prefix')
 if kind in ('caption_number_before_seq','caption_number_after_seq'):
  prefix=checked_caption_prefix(text);fields=[f for f in checked_fields(node) if re.fullmatch(r' SEQ [图表] \\\* ARABIC ',f['instruction']) and f['cache'].isdigit()]
  if prefix is None or len(fields)!=1:raise ValueError('Caption prefix edit requires its unique target SEQ')
  lo,hi=checked_field_span(p,fields[0])
  if not prefix.start('number')<=lo<hi<=prefix.end('number'):raise ValueError('SEQ is outside caption ordinal')
  valid=(a==0 and b==lo and new in ('图 ','表 ')) if kind=='caption_number_before_seq' else (a==hi and b==prefix.end() and new=='　')
  if not valid:raise ValueError('Native caption edit exceeds its old ordinal')
 if kind=='mapped_heading_reference' and not all(re.fullmatch(r'\d+(?:\.\d+)*',x) for x in (old,new)):raise ValueError('Invalid chapter reference change')
 if kind=='caption_label' and not (re.fullmatch(r'\s*[图表]\s*',old) and new in ('图 ','表 ')):raise ValueError('Invalid caption label')
 if kind=='caption_separator' and (not re.fullmatch(r'[.．、]?\s*',old) or new!='　'):raise ValueError('Invalid caption spacing')
 if kind=='materialize_inherited_bullet' and (old or new!='• '):raise ValueError('Invalid list marker')
 if kind=='remove_unused_equation_label' and (new or not re.fullmatch(r'\s*式\s*\d+\s*',old)):raise ValueError('Invalid label removal')
 if kind=='remove_equation_number_tail' and (new or not re.fullmatch(r'\s*[（(]\s*\d+\s*[）)]\s*',old) or not (p.findall('.//m:oMath',N) or any(n.get('equationxml') is not None or n.get('ProgID','').lower().startswith(('equation.','mathtype.')) for n in p.iter())) or text!=old):raise ValueError('Invalid equation number removal')
 if kind=='adjacent_equation_reference' and (not re.fullmatch(r'式\s*\d+',old) or new!='上述关系'):raise ValueError('Invalid adjacent reference migration')
 if kind=='missing_media_notice' and (old or a!=b or new!=' [图片资料缺失：'+change['source_id']+':'+change['block_id']+'；关系'+change['relationship_id']+']'):raise ValueError('Invalid missing-media notice')
 if kind not in ('source_display_title','source_title_presentation','mapped_heading_reference','caption_label','caption_separator','caption_text_number','caption_text_label','caption_text_separator','caption_number_before_seq','caption_number_after_seq','mapped_object_reference','heading_number','source_group_number','manual_list_prefix','caption_number','materialize_inherited_bullet','remove_unused_equation_label','remove_equation_number_tail','adjacent_equation_reference','missing_media_notice'):raise ValueError('Unknown text operation')
 if a==b and a>0:
  offset=0
  for t in p.xpath('.//w:t[not(ancestor::m:oMath)]',namespaces=N):
   value=t.text or ''
   if offset<=a<offset+len(value):
    t.text=value[:a-offset]+new+value[a-offset:];return
   offset+=len(value)
  if a!=offset:raise ValueError('Insertion beyond paragraph')
  run=E.SubElement(p,q('w:r'));E.SubElement(run,q('w:t')).text=new;return
 if a==b:
  run=E.Element(q('w:r'));t=E.SubElement(run,q('w:t'));t.text=new;t.set('{http://www.w3.org/XML/1998/namespace}space','preserve');p.insert(1 if p.find('w:pPr',N) is not None else 0,run);return
 pos=0;first=True
 for t in p.xpath('.//w:t[not(ancestor::m:oMath)]',namespaces=N):
  value=t.text or '';start=pos;pos+=len(value)
  left=max(a,start);right=min(b,pos)
  if left<right:t.text=value[:left-start]+(new if first else '')+value[right-start:];first=False

def bookmark_texts(body):
 values={};active={}
 for node in body.iter():
  if node.tag==q('w:bookmarkStart'):
   name=node.get(q('w:name'));active[node.get(q('w:id'))]=name;values[name]=''
  elif node.tag==q('w:bookmarkEnd'):active.pop(node.get(q('w:id')),None)
  elif node.tag==q('w:t'):
   for name in active.values():values[name]+=node.text or ''
 return values


def checked_source_layout(package,p):
 """Read effective source properties independently of scan/recognition."""
 if not hasattr(package,'layout_styles'):
  package.layout_styles={s.get(q('w:styleId')):s for s in E.fromstring(package.parts['word/styles.xml']).findall('w:style',N)}
 styles=package.layout_styles
 default=next((sid for sid,s in styles.items() if s.get(q('w:default'))=='1' and s.get(q('w:type'))=='paragraph'),'Normal')
 st=p.find('w:pPr/w:pStyle',N);sid=st.get(q('w:val')) if st is not None else default
 nodes=[p];seen=set();current=sid
 while current in styles and current not in seen:
  seen.add(current);s=styles[current];nodes.append(s);base=s.find('w:basedOn',N);current=base.get(q('w:val')) if base is not None else None
 def effective(path,attribute='val'):
  return next((n.get(q('w:'+attribute)) for owner in nodes for n in owner.findall(path,N) if n.get(q('w:'+attribute)) is not None),None)
 fields={k:float(effective('w:pPr/w:ind',k) or 0) for k in ('firstLine','firstLineChars','hanging','hangingChars','left','leftChars','right','rightChars')}
 shape=[effective('w:pPr/w:jc') or 'left']
 for key in ('firstLine','hanging','left','right'):
  c,t=fields[key+'Chars'],fields[key];shape.append(('chars',c) if c else ('twips',t) if t else ('zero',0))
 size=None
 first=p.find('w:r/w:rPr/w:sz',N)
 if first is not None:size=float(first.get(q('w:val')))/2
 elif effective('w:rPr/w:sz') is not None:size=float(effective('w:rPr/w:sz'))/2
 two=not any(fields[k] for k in ('hanging','hangingChars','left','leftChars','right','rightChars')) and (fields['firstLineChars']==200 or (not fields['firstLineChars'] and size is not None and abs(fields['firstLine']-2*size*20)<1))
 name=styles[sid].find('w:name',N) if sid in styles else None
 outline=effective('w:pPr/w:outlineLvl')
 return tuple(shape),two,name.get(q('w:val'),'').casefold() if name is not None else '',int(outline) if outline is not None else None


def checked_caption_role(source,row,candidate,baseline,known_ids):
 """Require source role evidence, not the writer's auto_rule assertion."""
 if row['role'] not in ('figure_caption','table_caption'):return None
 decision=candidate.get('decision') or {}
 if decision:
  if (decision.get('block_id')==row['block_id'] and decision.get('role')==row['role'] and decision.get('relates_to')==row.get('relates_to')
      and decision.get('origin') in ('model','user') and decision.get('decision_id') and decision.get('reason')
      and decision.get('evidence_ids') and set(decision['evidence_ids'])<=known_ids):return None
  return 'caption_role_invalid_local_decision'
 p=source.body[row['source_index']];text=ordinary(p).strip();shape,two,style,_=checked_source_layout(source,p)
 prefix=checked_caption_prefix(text);seq=False
 remainder=text[prefix.end():].lstrip() if prefix else text
 explanatory=prefix is not None and bool(re.match(r'^(?:[中内上下](?:可见|显示|所示)|所示(?:为|的)|显示(?:了|出|不同|该|这|其|一个|一种)|说明了|表明|给出(?:了|的是)|展示(?:了|出)|反映(?:了|出)|可见|可以|分别为|是|为|对应于|包含了|采用了)',remainder) or
     (re.search(r'[。；;！？!?]$',text) and (re.search(r'[，,；;]',remainder) or re.match(r'^的.*(?:一致|不同|相同|表明|显示|说明|可见|不得)',remainder))))
 if len(text)>100 or (len(text)>30 and re.search(r'[。；;！？!?]',text)) or explanatory:return 'caption_role_prose_requires_local_decision'
 if prefix:
  for f in checked_fields(p):
   if re.match(r'\s*SEQ\s+',f['instruction'],re.I) and f['cache'].isdigit():
    a,b=checked_field_span(p,f)
    seq=seq or prefix.start('number')<=a<b<=prefix.end('number')
 styled=style in ('图题','表题','caption','题注','图注','表注')
 distinct=shape[1:]!=baseline[1:] and not two
 if not (styled or seq or shape[0]=='center' or distinct):return 'compact_caption_needs_independent_adjacency' if prefix else 'caption_role_no_distinct_source_layout'
 if not styled and not seq:
  target=candidate.get('relates_to')
  if not target:return 'caption_role_missing_target'
  return 'layout_caption_needs_independent_adjacency' # Caller resolves source block identities.
 return None

def verify(report_path):
 r=json.loads(Path(report_path).read_text(encoding='utf8'));out=Package(r['output']);template=Package(r['template']);errors=[];count=0
 def ck(v,name):
  nonlocal count
  count+=1
  if not v:errors.append(name)
 ck(digest(out.path.read_bytes())==r['output_sha256'],'output_hash')
 slots=template.body.xpath('.//w:sdt[w:sdtPr/w:tag[@w:val="wps-report-body-v1"]]',namespaces=N)
 anchors=template.body.xpath('.//w:bookmarkStart[@w:name="WPS_ReportBody"]',namespaces=N)
 prefix=[]
 if anchors:
  mark=anchors[0];paragraph=mark.getparent();ident=mark.get(q('w:id'))
  starts=template.body.xpath('.//w:bookmarkStart[@w:id=$ident]',ident=ident,namespaces=N)
  ends=template.body.xpath('.//w:bookmarkEnd[@w:id=$ident]',ident=ident,namespaces=N)
  valid=(not slots and len(anchors)==1 and len(starts)==1 and len(ends)==1 and ident is not None
         and paragraph.tag==q('w:p') and paragraph.getparent() is template.body
         and template.body[-2] is paragraph and ends[0].getparent() is paragraph
         and list(paragraph).index(mark)<list(paragraph).index(ends[0]) and not ''.join(paragraph.xpath('.//w:t/text()',namespaces=N)).strip())
  if valid:
   allowed={q('w:p'),q('w:r'),q('w:t'),q('w:pPr'),q('w:rPr')}
   valid=all(n in (mark,ends[0]) or n.tag in allowed or any(a.tag in (q('w:pPr'),q('w:rPr')) for a in n.iterancestors()) for n in paragraph.iter())
  ck(valid,'unambiguous_template_body_bookmark')
  if valid:prefix=list(template.body)[:list(template.body).index(paragraph)]
  ck(r.get('target_contract')=='preserved-frontmatter-v1','recognized_target_layout_contract')
  ck(not out.body.xpath('.//w:bookmarkStart[@w:name="WPS_ReportBody"]',namespaces=N),'template_body_marker_consumed')
 elif slots:
  valid=len(slots)==1 and len(template.body)>=2 and slots[0].getparent() is template.body and template.body[-2] is slots[0]
  ck(valid,'unambiguous_template_body_slot')
  if valid:prefix=list(template.body)[:list(template.body).index(slots[0])]
  ck(r.get('target_contract')=='preserved-frontmatter-v1','recognized_target_layout_contract')
 else:ck(r.get('target_contract')=='continuous-body-v1','recognized_target_layout_contract')
 ck(len(out.body)>=len(prefix)+1 and all(tree(a)==tree(b) for a,b in zip(prefix,list(out.body)[:len(prefix)])),'preserved_template_frontmatter')
 ck(digest(template.path.read_bytes())==r['template_sha256'],'template_hash')
 editable={'word/document.xml','word/_rels/document.xml.rels','[Content_Types].xml'}
 if r.get('numbering_definitions_imported'):editable.add('word/numbering.xml')
 ck(all(out.parts.get(p)==b for p,b in template.parts.items() if p not in editable),'preserve_only_template_parts')
 ck(tree(out.body.find('w:sectPr',N))==tree(template.body.find('w:sectPr',N)),'target_section')
 styles=E.fromstring(template.parts['word/styles.xml']);style_names={s.find('w:name',N).get(q('w:val')).lower():s.get(q('w:styleId')) for s in styles.findall('w:style',N) if s.find('w:name',N) is not None}
 known_ids=set(style_names.values())
 style_nodes={s.get(q('w:styleId')):s for s in styles.findall('w:style',N)}
 def effective_flag(p,key):
  direct=p.find('w:pPr/w:'+key,N)
  if direct is not None:return direct.get(q('w:val'),'1') not in ('0','false','off')
  st=p.find('w:pPr/w:pStyle',N);sid=st.get(q('w:val')) if st is not None else style_names['normal'];seen=set()
  while sid in style_nodes and sid not in seen:
   seen.add(sid);node=style_nodes[sid];flag=node.find('w:pPr/w:'+key,N)
   if flag is not None:return flag.get(q('w:val'),'1') not in ('0','false','off')
   base=node.find('w:basedOn',N);sid=base.get(q('w:val')) if base is not None else None
  return False
 report_paragraphs=[p for node in list(out.body)[len(prefix):-1] for p in node.iter(q('w:p'))]
 for p in report_paragraphs:
  st=p.find('w:pPr/w:pStyle',N);ck(st is not None and st.get(q('w:val')) in known_ids,'all_paragraph_styles_resolve')
  ck(not effective_flag(p,'pageBreakBefore'),'no_paragraph_forced_page_break')
 generated=r['generated_blocks'];generated_indices=[b['output_index'] for b in generated]
 physical_contract=r.get('object_numbering_contract')=='physical-objects-v1'
 nested_captions={out.body[o['caption_output_index']].xpath(o['caption_path'],namespaces=N)[0] for o in r.get('object_numbers',[]) if o['generated'] and o['caption_path']!='.'}
 ck(len(generated_indices)==len(set(generated_indices)),'unique_generated_block_indices')
 ck(len({s['source_id'] for s in r['sources']})==len(r['sources']),'unique_source_ids')
 for b in generated:
  ck(b.get('kind') in (None,'source_group_title','source_gap','source_gap_heading','object_number_caption','source_document_title','organization_group'),'known_generated_block_kind')
  p=out.body[b['output_index']]
  expected_style=('图题' if b.get('label')=='图' else '表题') if b.get('kind')=='object_number_caption' else 'normal' if b.get('kind')=='source_gap' else 'heading 1' if b.get('kind') in ('source_group_title','source_gap_heading') else 'title'
  if b.get('kind')=='organization_group':expected_style='heading '+str(len(b['prefix']))
  if b.get('kind') in ('source_group_title','source_gap_heading') and r.get('organization'):expected_style='heading '+str(len(b['source_prefix']))
  ps=p.find('w:pPr/w:pStyle',N)
  ck(ordinary(p)==b['text'] and ps is not None and ps.get(q('w:val'))==style_names[expected_style],'generated_title_style_and_text')
  if b.get('kind')=='source_group_title':
   ck(bool(r['master']),'source_group_only_in_master')
   src=next((s for s in r['sources'] if s['source_id']==b.get('source_id')),None)
   detail=src.get('generated_group_title') if src else None
   ck(bool(detail) and detail['text']==b['text']==(detail['label'] if r.get('native_chapter_num_id') else '.'.join(map(str,detail.get('source_prefix') or [detail['source_number']]))+' '+detail['label']),'source_group_label_provenance')
 gaps=r.get('gaps',[])
 gap_blocks=[b for b in generated if b.get('kind')=='source_gap']
 ck(len(gaps)==len(gap_blocks),'gap_records_match_visible_markers')
 ck(len({g['source_id'] for g in gaps})==len(gaps),'unique_gap_source_ids')
 ck(not {s['source_id'] for s in r['sources']}&{g['source_id'] for g in gaps},'gap_is_not_counted_as_imported')
 scope=sorted([(s.get('source_number',i),s['source_id']) for i,s in enumerate(r['sources'],1)]+[(g['source_number'],g['source_id']) for g in gaps])
 if 'selected_source_ids' in r:
  ck([sid for _,sid in scope]==r['selected_source_ids'],'all_selected_sources_imported_or_explicit_gap_in_order')
 for g in gaps:
  ck(Path(g['path']).is_file() and digest(Path(g['path']).read_bytes())==g['sha256'],g['source_id']+':gap_original_attached')
  ck(any(b['source_id']==g['source_id'] and b['output_index']==g['output_index'] for b in gap_blocks),g['source_id']+':gap_visible_in_document')
 for error in verify_retained_numbering(r,out):ck(False,error)
 output_bookmark_text=bookmark_texts(out.body)
 allout=[];counts=Counter();details=[]
 previous_source_end=-1
 for source_order,src in enumerate(r['sources'],1):
  source=Package(src['path']);sid=src['source_id'];ck(digest(source.path.read_bytes())==src['sha256'],sid+':source_hash')
  from .verify_native_numbering import labels,effective
  original_numbering=labels(source);original_numbering_property=effective(source.parts)
  source.proven_missing=set()
  for repair in src.get('missing_media_repairs',[]):
   rid=repair['relationship_id'];original=source.rels.get(rid)
   used=[(n,k) for n in source.root.iter() for k,v in n.attrib.items() if k.startswith('{'+N['r']+'}') and v==rid]
   image_uses=bool(used) and all((n.tag==q('a:blip') and k in (q('r:embed'),q('r:link'))) or (n.tag==q('v:imagedata') and k==q('r:id')) for n,k in used)
   missing=original is None or original.get('TargetMode')=='External'
   if original is not None and not missing:
    target=original.get('Target','');part=target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join('word',target));missing=part not in source.parts
   ck(image_uses and missing and repair.get('original')==original,sid+':missing_media_really_missing:'+rid)
   expected='bba2cd57914964c1ff1feacfd38cba211e887273b00062ee299c95cc5e2a8ad7'
   data=next((data for key,data in out.parts.items() if key.startswith('word/media/') and digest(data)==expected),None)
   ck(repair.get('placeholder_sha256')==expected and data is not None,sid+':explicit_program_placeholder:'+rid)
   if image_uses and missing and data is not None:
    part='word/__missing_'+rid+'.png';source.parts[part]=data;source.rels[rid]={'Id':rid,'Type':N['r']+'/image','Target':'/'+part};source.proven_missing.add(rid)
  ids=src.get('legacy_id_map',{})
  allowed={n.get(k) for n in source.root.iter() for k in (('id',) if n.tag in (q('v:shape'),q('v:shapetype')) else ('ObjectID',) if n.tag==q('o:OLEObject') else ()) if n.get(k)}
  ck(set(ids)<=allowed and all(v==sid+'_'+k for k,v in ids.items()),sid+':legacy_identifier_scope')
  for n in source.root.iter():
   if n.tag in (q('v:shape'),q('v:shapetype')) and n.get('id') in ids:n.set('id',ids[n.get('id')])
   if n.tag==q('v:shape') and n.get('type','').startswith('#') and n.get('type')[1:] in ids:n.set('type','#'+ids[n.get('type')[1:]])
   if n.tag==q('o:OLEObject'):
    for k in ('ShapeID','ObjectID'):
     if n.get(k) in ids:n.set(k,ids[n.get(k)])
  source_bookmark_text=bookmark_texts(source.body)
  maps=[x for x in r['mapping'] if x['source_id']==sid];exc=[x for x in r['exclusions'] if x['source_id']==sid]
  whole={};kept={};selected_indices={m['source_index'] for m in maps}
  for i,root in enumerate(source.body):
   for n in root.iter():
    if n.tag not in (q('w:bookmarkStart'),q('w:bookmarkEnd')):continue
    ident=n.get(q('w:id'));whole.setdefault(ident,[]).append(n)
    if i in selected_indices:kept.setdefault(ident,[]).append(n)
  name_counts=Counter(n.get(q('w:name')) for nodes in kept.values() for n in nodes if n.tag==q('w:bookmarkStart'))
  required=[]
  for ident,events in kept.items():
   originals=whole[ident];starts=[n for n in events if n.tag==q('w:bookmarkStart')];ends=[n for n in events if n.tag==q('w:bookmarkEnd')]
   valid=ident is not None and len(starts)==len(ends)==1 and len(originals)==2 and originals[0].tag==q('w:bookmarkStart') and originals[1].tag==q('w:bookmarkEnd') and starts[0].get(q('w:name')) and name_counts[starts[0].get(q('w:name'))]==1
   if not valid:required.append(dict(bookmark_id=ident,source_starts=sum(n.tag==q('w:bookmarkStart') for n in originals),source_ends=sum(n.tag==q('w:bookmarkEnd') for n in originals),kept_starts=len(starts),kept_ends=len(ends)))
  ck(src.get('bookmark_detachments',[])==required,sid+':only_unmatched_scope_bookmarks_detached')
  approved_detachments={x['bookmark_id'] for x in required}
  groups=[b for b in generated if b.get('kind')=='source_group_title' and b.get('source_id')==sid]
  needs_group=bool(r['master']) and not any(x['role']=='document_title' for x in maps)
  ck(len(groups)==int(needs_group),sid+':exactly_one_required_source_group')
  ck(bool(src.get('generated_group_title'))==needs_group,sid+':source_group_report_presence')
  if groups:
   group=groups[0];detail=src.get('generated_group_title') or {}
   ck(group.get('source_number')==detail.get('source_number')==src.get('source_number',source_order),sid+':source_group_ordinal')
   own_indices=[x['output_index'] for x in maps]+[b['output_index'] for b in generated if b.get('kind')=='object_number_caption' and b.get('source_id')==sid]
   ck(group['output_index']>previous_source_end and bool(maps) and group['output_index']+1==min(own_indices),sid+':source_group_precedes_own_content')
  if maps:
   ck(min(x['output_index'] for x in maps)>previous_source_end,sid+':source_groups_keep_selected_order')
   previous_source_end=max(x['output_index'] for x in maps)
  valid=[i for i,n in enumerate(source.body) if n.tag!=q('w:sectPr')];selected=[x['source_index'] for x in maps]
  ck(len(selected)==len(set(selected)),sid+':no_duplicate_source_blocks')
  ck(set(selected)|{x['source_index'] for x in exc}==set(valid),sid+':all_source_blocks_accounted')
  ck(not set(selected)&{x['source_index'] for x in exc},sid+':source_exclusion_does_not_overlap_content')
  scope=src.get('source_frontmatter') or {}
  for error in verify_source_frontmatter(Package(src['path']),scope):ck(False,sid+':'+error)
  ck(exc==[dict(x,source_id=sid,source_text=ordinary(source.body[x['source_index']])) for x in scope.get('exclusions',[])],sid+':exclusions_match_bound_source_scope')
  if any(b.get('kind')=='source_document_title' and b.get('source_id')==sid for b in generated):
   ck(not r['master'] and bool(scope.get('title')) and all(b['text']==scope['title']['text'] for b in generated if b.get('kind')=='source_document_title' and b.get('source_id')==sid),sid+':generated_title_from_actual_source')
  # Managed figure titles can cross empty paragraphs when moved directly
  # below their bound object. Object/role/placement checks remain independent;
  # all ordinary content (including orphan titles) must retain source order.
  moved_figure_titles={o.get('caption_block_id') for o in r.get('object_numbers',[]) if o['source_id']==sid and o['label']=='图'} if r.get('caption_role_contract')=='source-layout-v1' else set()
  natural=sorted(maps,key=lambda x:x['output_index']);without_moved=[x['source_index'] for x in natural if x['role']!='table_caption' and x['block_id'] not in moved_figure_titles]
  ck(without_moved==sorted(without_moved),sid+':source_order_except_declared_caption_moves')
  ck(digest(Path(src['candidate']).read_bytes())==src['candidate_sha256'],sid+':candidate_hash')
  cand=json.loads(Path(src['candidate']).read_text(encoding='utf8'));byindex={x['source_locator']['body_child_index']:x for x in cand['blocks']}
  source_body_shapes=[]
  for m in maps:
   node=source.body[m['source_index']]
   if node.tag!=q('w:p') or checked_caption_prefix(ordinary(node).strip()) or checked_pictures(node) or len(re.sub(r'\s+','',ordinary(node)))<=40:continue
   shape,_,_,outline=checked_source_layout(source,node)
   if outline not in range(6):source_body_shapes.append(shape)
  baseline=Counter(source_body_shapes).most_common(1)[0][0] if source_body_shapes else checked_source_layout(source,E.Element(q('w:p')))[0]
  source_caption_index={}
  for m in maps:
   c=byindex[m['source_index']]
   if c['role'] not in ('figure_caption','table_caption') or c.get('subcaptions'):continue
   source_prefix=checked_caption_prefix(ordinary(source.body[m['source_index']]))
   if source_prefix:
    key=(source_prefix['label'],source_prefix['number'].translate(str.maketrans('．－–—','.---')))
    source_caption_index.setdefault(key,[]).append(m)
  # The verifier inspects original physical objects, never the model's role
  # name alone. Missing associations must have a visible preservation record.
  held_roles=set()
  source_exceptions=[x for x in r['exceptions'] if x.get('source_id')==sid]
  for row in maps:
   if row['role'] not in ('figure_caption','table_caption'):continue
   c=byindex[row['source_index']];related=c.get('relates_to')
   if r.get('caption_role_contract')=='source-layout-v1':
    error=checked_caption_role(source,row,c,baseline,{x['block_id'] for x in cand['blocks']})
    if error in ('layout_caption_needs_independent_adjacency','compact_caption_needs_independent_adjacency'):
     conventional=error=='compact_caption_needs_independent_adjacency'
     related_row=next((m for m in maps if m['block_id']==related),None)
     error='caption_role_layout_is_not_adjacent'
     if related_row:
      a,b=sorted((row['source_index'],related_row['source_index']))
      preferred=(related_row['source_index']<row['source_index']) if row['role']=='figure_caption' else (related_row['source_index']>row['source_index'])
      if (not conventional or preferred) and all(n.tag==q('w:p') and not ordinary(n).strip() and not n.xpath('.//w:drawing|.//w:pict|.//w:object|.//m:oMath|.//w:fldChar|.//w:fldSimple|.//w:br',namespaces=N) for n in list(source.body)[a+1:b]):error=None
    ck(error is None,f'{sid}:{row["source_index"]}:independent_caption_role:{error}')
   ck(row.get('relates_to')==related,f'{sid}:{row["source_index"]}:caption_relation_record_matches_candidate')
   target=next((m for m in maps if m['block_id']==related),None)
   if related:
    physical=source.body[target['source_index']] if target is not None else None
    valid=physical is not None and related!=row['block_id'] and ((row['role']=='table_caption' and physical.tag==q('w:tbl') and not target.get('figure_layout')) or (row['role']=='figure_caption' and bool(checked_pictures(physical)) and (physical.tag==q('w:p') or (physical.tag==q('w:tbl') and target.get('role')=='figure' and target.get('figure_layout',{}).get('kind')=='table'))))
    ck(valid,f'{sid}:{row["source_index"]}:caption_target_physical_type')
   if not related or c['status'] in ('review','retained_unpaired'):
    if not physical_contract:held_roles.add(row['role'])
    ck(any(x.get('block_id')==row['block_id'] and x.get('kind')=='caption_association_unresolved' for x in source_exceptions),f'{sid}:{row["source_index"]}:unresolved_caption_is_reported')
  for role in held_roles:
   ck(any(x.get('kind')=='caption_numbering_group_retained' and x.get('caption_role')==role for x in source_exceptions),f'{sid}:{role}:retained_numbering_group_reported')
   ck(not any(c['operation'] in ('caption_number','materialize_caption_seq') for m in maps if m['role']==role for c in m['changes']),f'{sid}:{role}:no_partial_numbering_of_unresolved_group')
  # Also check for OMITTED reference edits. Replaying only declared edits
  # cannot detect a whole ordinary paragraph being silently skipped.
  for row in maps:
   if not physical_contract or row['role'] in ('figure_caption','table_caption','heading','document_title') or row.get('figure_layout'):continue
   original=source.body[row['source_index']];candidate=byindex[row['source_index']]
   paragraphs=[original] if original.tag==q('w:p') else original.findall('.//w:p',N)
   for pi,p in enumerate(paragraphs):
    # Fields, linked text and drawing stories have their own binding checks;
    # do not turn this plain-text completeness check into permission to edit.
    if p.xpath('.//w:fldChar|.//w:fldSimple|.//w:hyperlink|.//w:object|.//w:pict|.//w:drawing',namespaces=N):continue
    text=ordinary(p);own=checked_caption_prefix(text) if candidate['role'] in ('figure_caption','table_caption') else None
    for ref in re.finditer(r'(?<![子附])(?P<label>[图表])\s*(?P<number>\d+(?:[.．\-－–—]\d+)*)(?!\d|[.．\-－–—]\d)',text):
     if own and ref.start()<own.end():
      ck(any(x.get('kind')=='unmapped_object_reference_preserved' and (x.get('block_id')==row['block_id'] or x.get('source_index')==row['source_index']) and x.get('text')==ref[0] for x in source_exceptions),f'{sid}:{row["source_index"]}:retained_caption_prefix_is_reported')
      continue
     key=(ref['label'],ref['number'].translate(str.maketrans('．－–—','.---')))
     targets=source_caption_index.get(key,[])
     if len(targets)!=1:continue
     cap=targets[0]
     objects=[o for o in r.get('object_numbers',[]) if o['source_id']==sid and o.get('caption_block_id')==cap['block_id'] and o['label']==ref['label']]
     if len(objects)!=1 or ref['number']==objects[0]['number']:continue
     edits=[c for c in row['changes'] if c['operation']=='mapped_object_reference' and c.get('paragraph_index')==pi and (c.get('start'),c.get('end'))==ref.span('number') and c.get('old')==ref['number'] and c.get('new')==objects[0]['number'] and c.get('target_block_id')==cap['block_id']]
     ck(len(edits)==1,f'{sid}:{row["source_index"]}:unique_plain_object_reference_not_silently_omitted')
  for row in maps:
   si=row['source_index'];oi=row['output_index'];allout.append(oi);node=deepcopy(source.body[si]);target=out.body[oi]
   clean_target=deepcopy(target)
   added=[clean_target.xpath(o['caption_path'],namespaces=N)[0] for o in r.get('object_numbers',[]) if o['generated'] and o['caption_output_index']==oi and o['caption_path']!='.']
   for added_node in added:added_node.getparent().remove(added_node)
   try:
    converted_nodes=set()
    for change in row['changes']:
     if change['operation']=='unresolved_numbering_notice':
      original_node=source.body[si];ps=list(original_node.iter(q('w:p')));pi=change['paragraph_index'];fact=original_numbering.get(ps[pi],{}) if type(pi) is int and 0<=pi<len(ps) else {}
      nid=change['source_num_id'];raw_numbering=E.fromstring(source.parts.get('word/numbering.xml',b'<w:numbering xmlns:w="'+N['w'].encode()+b'"/>'))
      definition=next((n for n in raw_numbering.findall('w:num',N) if n.get(q('w:numId'))==nid),None);aid=definition.find('w:abstractNumId',N) if definition is not None else None
      abstract=next((n for n in raw_numbering.findall('w:abstractNum',N) if n.get(q('w:abstractNumId'))==aid.get(q('w:val'))),None) if aid is not None else None
      reason=change['reason'];allowed=False
      if reason=='missing_definition':allowed=abstract is None
      elif reason=='invalid_level':allowed=fact.get('ilvl') is None or not 0<=fact['ilvl']<=8
      elif reason=='dependent_definition':allowed=abstract is not None and any(n.tag in (q('w:numStyleLink'),q('w:styleLink'),q('w:pStyle'),q('w:lvlPicBulletId')) or any(k.startswith('{'+N['r']+'}') for k in n.attrib) for base in (abstract,definition) for n in base.iter())
      elif reason=='incomplete_instance':
       selected_indices={m['source_index'] for m in maps}
       for p,fact2 in original_numbering.items():
        if fact2.get('num_id')!=nid:continue
        owner=p
        while owner.getparent() is not source.body:owner=owner.getparent()
        if source.body.index(owner) not in selected_indices:allowed=True
      elif reason=='external_story_instance':
       allowed=any(original_numbering_property(p,'w:numPr/w:numId')==nid for part,data in source.parts.items() if re.fullmatch(r'word/(?:header[^/]*|footer[^/]*|footnotes|endnotes|comments)\.xml',part) for p in E.fromstring(data).iter(q('w:p')))
      ck(allowed and fact.get('status')=='unresolved' and fact.get('num_id')==nid,f'{sid}:{si}:numbering_notice_has_unresolved_source_evidence')
      ck(any(x.get('source_id')==sid and x.get('block_id')==row['block_id'] and x.get('kind')=='automatic_numbering_label_unavailable' and x.get('reason')==reason for x in source_exceptions),f'{sid}:{si}:numbering_display_gap_is_reported')
      target_ps=list(target.iter(q('w:p')))
      value=target_ps[pi].find('w:pPr/w:numPr/w:numId',N) if type(pi) is int and 0<=pi<len(target_ps) else None
      ck(value is not None and value.get(q('w:val'))=='0',f'{sid}:{si}:unsafe_numbering_is_detached_only_locally')
     if change['operation']=='normalize_embedded_story_styles':
      ck(change['target_style_id']==style_names['normal'],f'{sid}:{si}:embedded_story_uses_declared_target_style')
      for p in target.findall('.//w:p',N):
       if any(n.tag in (q('w:pict'),q('w:object'),q('w:txbxContent')) for n in p.iterancestors()):
        st=p.find('w:pPr/w:pStyle',N);ck(st is not None and st.get(q('w:val'))==style_names['normal'],f'{sid}:{si}:embedded_style_mapping_readback')
     if change['operation']=='detach_unmatched_bookmark':ck(change['bookmark_id'] in approved_detachments,f'{sid}:{si}:bookmark_detachment_has_original_scope_evidence')
     if change['operation']=='missing_link_placeholder_as_embed':ck(change['relationship_id'] in source.proven_missing,f'{sid}:{si}:placeholder_link_has_missing_source_data')
     if change['operation']=='drop_image_link':
      candidates=node.xpath(change['path'],namespaces=N);rid=candidates[0].get(q('r:embed')) if len(candidates)==1 else None
      rel=source.rels.get(rid);rel_target=rel.get('Target','') if rel else '';part=rel_target.lstrip('/') if rel_target.startswith('/') else posixpath.normpath(posixpath.join('word',rel_target))
      ck(rid not in source.proven_missing and rel is not None and rel.get('Type','').endswith('/image') and rel.get('TargetMode')!='External' and part in source.parts,f'{sid}:{si}:link_removed_only_with_available_embed')
     if change['operation']=='source_display_title':
      def titles(node):
       return [c.get('title') for c in node.get('children',[]) if c.get('source_id')==sid]+[t for c in node.get('children',[]) if c.get('kind')=='group' for t in titles(c)]
      ck(bool(r.get('organization')) and titles(r.get('organization',{}).get('tree',{}))==[change['new']] and row['role']=='document_title',f'{sid}:{si}:display_title_from_organization')
     if change['operation'] in ('freeze_excluded_frontmatter_reference','unlink_excluded_frontmatter_anchor'):
      removed={m.get(q('w:name')) for e in exc for m in source.body[e['source_index']].iter(q('w:bookmarkStart'))}
      ck(change.get('bookmark',change.get('anchor')) in removed,f'{sid}:{si}:reference_target_was_explicitly_excluded')
     if change['operation']=='missing_media_notice':ck(change['relationship_id'] in source.proven_missing and change['source_id']==sid and change['block_id']==row['block_id'],f'{sid}:{si}:bound_missing_media_notice')
     numeric_error=heading_edit_error(ordinary(source.body[si]),change)
     ck(numeric_error is None,f'{sid}:{si}:independent_number_guard:{numeric_error}')
     if change['operation']=='adjacent_equation_reference':
      ref=change['target_source_index'];number=re.search(r'\d+',change['old']).group()
      ck(ref==si-1 and bool(re.match(r'^\s*式\s*'+number+r'\s+',ordinary(source.body[ref]))),f'{sid}:{si}:adjacent_equation_reference_target')
     if change['operation'] in ('materialize_caption_seq','materialize_caption_ref','native_caption_seq','native_caption_ref'):
      target_row=next((m for m in maps if m['block_id']==change['target_block_id']),None)
      ck(target_row is not None and target_row['role'] in ('figure_caption','table_caption'),f'{sid}:{si}:field_target_caption')
      if target_row is not None:
       text=ordinary(out.body[target_row['output_index']]);caption=re.match(r'^[图表] (\d+)　',text)
       ck(caption is not None and caption[1]==change['new'],f'{sid}:{si}:field_number_matches_caption')
       if change['operation'] in ('materialize_caption_ref','native_caption_ref'):
        old_target=source.body[target_row['source_index']]
        marks=[m for m in old_target.findall('w:bookmarkStart',N) if m.get(q('w:name'))==change['bookmark']]
        ck(len(marks)==1 and change['old']==source_bookmark_text.get(change['bookmark']),f'{sid}:{si}:source_ref_bookmark_binding')
        mapped=src['bookmark_name_map'].get(change['bookmark'])
        ck(mapped is not None and output_bookmark_text.get(mapped)==change['new'],f'{sid}:{si}:output_bookmark_contains_reference_number')
        if change['operation']=='native_caption_ref':ck(re.match(r'\s*REF\s+'+re.escape(mapped or '?')+r'(?=\s|$)',change['new_instruction'],re.I) is not None,f'{sid}:{si}:native_ref_preserves_bound_target')
     if change['operation'] in ('remap_simple_reference_field','remap_native_reference_field'):
      expected_name=src['bookmark_name_map'].get(change['bookmark'])
      ck(expected_name is not None and re.match(r'\s*(?:REF|PAGEREF|NOTEREF)\s+'+re.escape(expected_name)+r'(?=\s|$)',change['new_instruction'],re.I) is not None,f'{sid}:{si}:reference_field_name_mapping')
     if change['operation']=='mapped_heading_reference':
      matches=[m for m in maps if m['role']=='heading' and (byindex[m['source_index']].get('number_candidate') or {}).get('numbers') and '.'.join(map(str,byindex[m['source_index']]['number_candidate']['numbers']))==change['old']]
      ck(len(matches)==1 and (r.get('native_chapter_num_id') or ordinary(out.body[matches[0]['output_index']]).startswith(change['new']+' ')),f'{sid}:{si}:chapter_reference_unique_target')
     if change['operation']=='mapped_object_reference':
      p=([node] if node.tag==q('w:p') else node.findall('.//w:p',N))[change['paragraph_index']]
      refs=list(re.finditer(r'(?<![子附])(?P<label>[图表])\s*(?P<number>\d+(?:[.．\-－–—]\d+)*)(?!\d|[.．\-－–—]\d)',ordinary(p)))
      ref=next((x for x in refs if x.span('number')==(change['start'],change['end'])),None)
      key=(ref['label'],ref['number'].translate(str.maketrans('．－–—','.---'))) if ref else None
      matches=source_caption_index.get(key,[])
      valid=len(matches)==1 and matches[0]['role'] in ('figure_caption','table_caption')
      if valid:
       cap=matches[0];target_prefix=checked_caption_prefix(ordinary(out.body[cap['output_index']]))
       valid=(target_prefix is not None and target_prefix['label']==ref['label'] and target_prefix['number']==change['new'] and
              change.get('target_block_id',cap['block_id'])==cap['block_id'])
       if physical_contract:valid=valid and any(o['source_id']==sid and o.get('caption_block_id')==cap['block_id'] and o['number']==change['new'] for o in r.get('object_numbers',[]))
      ck(valid,f'{sid}:{si}:complete_object_reference_unique_source_target')
     replay(node,change,converted_nodes)
    observed=out.content(clean_target)
    layout=row.get('figure_layout')
    ck(layout==byindex[si].get('figure_layout'),f'{sid}:{si}:figure_layout_matches_candidate')
    if layout and layout.get('kind')=='inline-row':
     drawings=node.findall('.//wp:inline',N)
     ck(node.tag==q('w:p') and not ordinary(node).strip() and 2<=len(drawings)<=4 and len(drawings)==layout.get('columns'),f'{sid}:{si}:physical_inline_group')
     ck(target.tag==q('w:tbl') and len(target.findall('w:tr',N))==1 and len(target.findall('w:tr/w:tc',N))==len(drawings),f'{sid}:{si}:inline_group_cell_count')
     events=[]
     for cell in target.findall('w:tr/w:tc',N):
      ps=cell.findall('w:p',N);ck(len(ps)==1 and len(ps[0].findall('.//wp:inline',N))==1,f'{sid}:{si}:one_picture_per_inline_group_cell')
      for p in ps:
       for kind,value in out.content(p)[1]:
        if kind=='text' and events and events[-1][0]=='text':events[-1][1]+=value
        else:events.append([kind,value])
     observed=['p',events]
    ck(source.content(node,src['bookmark_id_map'],src['bookmark_name_map'])==observed,f'{sid}:{si}:ordered_content_math_links_tables')
   except Exception as e:ck(False,f'{sid}:{si}:replay_error:{e}')
   ck(ordinary(clean_target)==ordinary(node) and ordinary(target)==row['output_text'],f'{sid}:{si}:text_readback')
   if row.get('retained_caption_role'):
    c=byindex[si]
    ck(c['role']==row['retained_caption_role'] and row['role']=='body' and any(x.get('block_id')==row['block_id'] and x.get('kind')=='caption_association_unresolved' for x in source_exceptions),f'{sid}:{si}:unpaired_title_retained_as_body')
   if target.tag==q('w:p'):
    if source.body[si].findall('.//m:oMath',N) and (not ordinary(source.body[si]).strip() or re.fullmatch(r'\s*[（(]\s*\d+\s*[）)]\s*',ordinary(source.body[si]))):
     original=source.body[si];tail=ordinary(original)
     retained=any(x.get('kind')=='referenced_equation_number_retained' and x.get('block_id')==row['block_id'] and
         re.fullmatch(r'\s*[（(]\s*'+re.escape(x.get('number','?'))+r'\s*[）)]\s*',tail) and
         checked_equation_reference(source.body,x.get('number'),original) for x in source_exceptions)
     pending=[x for x in source_exceptions if x.get('kind')=='equation_layout_retained' and x.get('block_id')==row['block_id']]
     layout_ok=bool(pending) and all((x.get('reason')=='nested_native_math_carrier' and not original.findall('m:oMath',N) and original.findall('.//m:oMath',N)) or
         (x.get('reason')=='retained_visible_content' and retained and ordinary(target)==tail) for x in pending)
     ck((bool(target.findall('m:oMathPara',N)) or layout_ok) and (not ordinary(target).strip() or (retained and ordinary(target)==tail)),f'{sid}:{si}:independent_standalone_equation_contract')
     ck(target.find('w:pPr/w:pStyle',N).get(q('w:val'))==style_names['公式'],f'{sid}:{si}:independent_formula_style')
    ps=target.find('w:pPr/w:pStyle',N);ck(ps is not None,f'{sid}:{si}:paragraph_style_present')
    depth=len(src.get('source_prefix') or [src.get('source_number',source_order)]) if r['master'] else 0
    roles={'document_title':'heading '+str(depth) if r['master'] else 'title','front_heading':'heading 2','body':'normal','spacer':'normal','figure':'图形段','figure_caption':'图题','table_caption':'表题','equation':'公式','equation_description':'normal','list_item':'list paragraph','reference':'参考条目','table_text':'list paragraph','boundary':'normal'}
    expected=('heading '+str(row['level']+depth) if row.get('level') in range(1,6) else 'normal') if row['role']=='heading_retained' else 'heading '+str(row['level']+depth) if row['role']=='heading' else roles.get(row['role'])
    if row['role']=='heading_retained' and expected not in style_names:
     ck(any(x.get('kind')=='retained_heading_depth_unavailable' and x.get('source_id')==sid and x.get('block_id')==row['block_id'] for x in r['exceptions']),f'{sid}:{si}:retained_heading_depth_reported');expected='normal'
    ck(expected is not None and ps is not None and ps.get(q('w:val'))==style_names[expected],f'{sid}:{si}:role_style')
    if row['role']=='heading' and not r.get('native_chapter_num_id'):ck(bool(re.match(r'^\d+(?:\.\d+)* \S',ordinary(target))),f'{sid}:{si}:heading_number_flush_left')
   if target.tag==q('w:tbl') and row.get('figure_layout'):
    layout=row['figure_layout'];ck(row['role']=='figure',f'{sid}:{si}:layout_is_figure')
    ck(bool(checked_pictures(source.body[si])),f'{sid}:{si}:source_layout_contains_images')
    ck(not any(m['role']=='table_caption' and m.get('relates_to')==row['block_id'] for m in maps),f'{sid}:{si}:figure_layout_has_no_table_caption')
    if layout['kind']=='table':ck(source.body[si].tag==q('w:tbl'),f'{sid}:{si}:source_is_image_table')
    pr=target.find('w:tblPr',N);borders=pr.find('w:tblBorders',N)
    ck(borders is not None and len(borders)==6 and all(n.get(q('w:val'))=='nil' for n in borders),f'{sid}:{si}:figure_layout_borderless')
    ck(pr.find('w:tblStyle',N) is None and pr.find('w:jc',N).get(q('w:val'))=='center',f'{sid}:{si}:no_data_table_style_on_figure')
    for tr in target.findall('.//w:tr',N):
     ck(not tr.findall('w:trPr/w:tblHeader',N) and not tr.findall('w:trPr/w:trHeight',N),f'{sid}:{si}:figure_no_header_or_fixed_height')
    for cell in target.findall('.//w:tc',N):
     cp=cell.find('w:tcPr',N);ck(cp.find('w:shd',N) is None and cp.find('w:tcBorders',N) is None,f'{sid}:{si}:figure_cell_no_data_decoration')
     for p in cell.findall('w:p',N):
      pp=p.find('w:pPr',N);ind=pp.find('w:ind',N)
      style='图形段' if checked_pictures(p) else '图题'
      ck(pp.find('w:pStyle',N).get(q('w:val'))==style_names[style],f'{sid}:{si}:figure_cell_role_style')
      ck(ind is not None and all(ind.get(q('w:'+key))=='0' for key in ('left','right','firstLine','hanging','leftChars','rightChars','firstLineChars','hangingChars')),f'{sid}:{si}:figure_cell_all_indents_zero')
      ck(pp.find('w:jc',N).get(q('w:val'))=='center',f'{sid}:{si}:figure_cell_center')
     cw=cp.find('w:tcW',N)
     if cw is not None:
      limit=max(200,int(cw.get(q('w:w')))-170)*635
      for image in cell.findall('.//wp:inline/wp:extent',N):ck(int(image.get('cx'))<=limit,f'{sid}:{si}:figure_fits_cell')
   elif target.tag==q('w:tbl'):
    for tr in target.findall('.//w:tr',N):
     flag=tr.find('w:trPr/w:cantSplit',N);ck(flag is None or flag.get(q('w:val'),'1') in ('0','false','off'),f'{sid}:{si}:table_rows_may_split')
    ck(target.find('w:tblPr/w:jc',N).get(q('w:val'))=='center',f'{sid}:{si}:table_center')
    for p in target.findall('.//w:tc/w:p',N):
     if p in nested_captions:continue
     pp=p.find('w:pPr',N);space=pp.find('w:spacing',N);ind=pp.find('w:ind',N)
     ck(pp.find('w:pStyle',N).get(q('w:val'))==style_names['表内正文'],f'{sid}:{si}:cell_style')
     ck(space is not None and space.get(q('w:line'))=='240' and space.get(q('w:lineRule'))=='auto',f'{sid}:{si}:cell_single_line')
     ck(ind is not None and ind.get(q('w:firstLine'))=='0' and ind.get(q('w:firstLineChars'))=='0',f'{sid}:{si}:cell_no_indent')
     ck(pp.find('w:jc',N).get(q('w:val'))=='center',f'{sid}:{si}:cell_horizontal_center')
     kn=pp.find('w:keepNext',N);ck(kn is not None and kn.get(q('w:val'))=='0',f'{sid}:{si}:no_table_keep_chain')
    ck(all(c.find('w:tcPr/w:vAlign',N).get(q('w:val'))=='center' for c in target.findall('.//w:tc',N)),f'{sid}:{si}:cell_vertical_center')
   # Direct emphasis is compared per character after replay, with hyperlinks
   # and math still included in the independent ordered-content comparison.
   if row['role'] in ('body','list_item','reference','table_text','table'):
    def emphasis(n,package):
     definitions={s.get(q('w:styleId')):s for s in E.fromstring(package.parts['word/styles.xml']).findall('w:style',N)}
     answer=[]
     for p in ([n] if n.tag==q('w:p') else n.findall('.//w:p',N)):
      tokens=[]
      for run in p.xpath('.//w:r[not(ancestor::m:oMath)]',namespaces=N):
       attrs=[]
       style_ref=run.find('w:rPr/w:rStyle',N);current=style_ref.get(q('w:val')) if style_ref is not None else None
       ancestry=[];visited=set()
       while current in definitions and current not in visited:
        visited.add(current);style=definitions[current];ancestry.insert(0,style)
        base=style.find('w:basedOn',N);current=base.get(q('w:val')) if base is not None else None
       for key in ('b','i','u','strike','dstrike','vertAlign'):
        inherited=None
        for style in ancestry:
         flag=style.find('w:rPr/w:'+key,N)
         if flag is not None:
          value=flag.get(q('w:val'),'1')
          if key in ('b','i','strike','dstrike'):
           if value not in ('0','false','off'):inherited='0' if inherited=='1' else '1'
          else:inherited=value
        e=run.find('w:rPr/w:'+key,N)
        effective=e.get(q('w:val'),'1') if e is not None else inherited
        if effective is not None and effective not in ('0','false','off','none'):attrs.append((key,effective))
       for ch in ''.join(run.xpath('./w:t/text()',namespaces=N)):tokens.append([ch,attrs])
      answer.append(tokens)
     return answer
    ck(emphasis(node,source)==emphasis(clean_target,out),f'{sid}:{si}:effective_character_style_emphasis')
   srcimg=node.findall('.//wp:inline',N);dstimg=target.findall('.//wp:inline',N)
   ck(len(srcimg)==len(dstimg),f'{sid}:{si}:drawing_count')
   for a,b in zip(srcimg,dstimg):
    ae=a.find('wp:extent',N);be=b.find('wp:extent',N);ax,ay=int(ae.get('cx')),int(ae.get('cy'));bx,by=int(be.get('cx')),int(be.get('cy'))
    ck(bx<=ax and by<=ay and abs(bx*ay-by*ax)<=max(ax,ay),f'{sid}:{si}:image_proportions')
   counts['blocks']+=1
  for key,path in [('tables','.//w:tbl'),('math','.//m:oMath'),('drawings','.//wp:inline')]:
   tag={'tables':q('w:tbl'),'math':q('m:oMath'),'drawings':q('wp:inline')}[key]
   counts[key]+=sum(sum(1 for _ in source.body[index].iter(tag)) for index in selected)
  details.append({'source':sid,'mapped':len(maps),'excluded':len(exc)})
 ck(len(allout)==len(set(allout)),'unique_output_mapping')
 ck(not set(allout)&set(generated_indices),'generated_blocks_do_not_overlap_source_blocks')
 ck(set(allout)|{x['output_index'] for x in r['generated_blocks']}|set(range(len(prefix)))=={i for i,n in enumerate(out.body) if n.tag!=q('w:sectPr')},'all_output_blocks_accounted')
 ck(not (set(allout)|set(generated_indices))&set(range(len(prefix))),'report_mapping_does_not_overlap_frontmatter')
 caption_ordinals=Counter()
 for row in sorted(r['mapping'],key=lambda x:x['output_index']):
  if not physical_contract and any(c['operation'] in ('caption_number','materialize_caption_seq','native_caption_seq','promote_caption_number') for c in row['changes']):
   label='图' if row['role']=='figure_caption' else '表';caption_ordinals[label]+=1
   ck(ordinary(out.body[row['output_index']]).startswith(label+' '+str(caption_ordinals[label])+'　'),'caption_numbers_follow_output_order')
 for instr in out.body.xpath('.//w:fldSimple/@w:instr',namespaces=N):
  target=re.match(r'\s*(?:REF|PAGEREF|NOTEREF)\s+(\w+)',instr,re.I)
  if target:ck(target[1] in output_bookmark_text,'remaining_reference_field_target_exists')
 starts=out.body.findall('.//w:bookmarkStart',N);ends=out.body.findall('.//w:bookmarkEnd',N)
 ids=[x.get(q('w:id')) for x in starts];names=[x.get(q('w:name')) for x in starts]
 ck(len(ids)==len(set(ids))==len(ends) and set(ids)=={x.get(q('w:id')) for x in ends},'bookmark_id_pairs_unique')
 ck(len(names)==len(set(names)),'bookmark_names_unique')
 original_unresolved=[]
 for src in r['sources']:
  p=Package(src['path']);bn={x.get(q('w:name')) for x in p.body.findall('.//w:bookmarkStart',N)}
  original_unresolved.extend(x.get(q('w:anchor')) for x in p.body.findall('.//w:hyperlink',N) if x.get(q('w:anchor')) and x.get(q('w:anchor')) not in bn)
 ck(all(x.get(q('w:anchor')) in set(names)|set(original_unresolved) for x in out.body.findall('.//w:hyperlink',N) if x.get(q('w:anchor'))),'no_new_internal_link_breaks')
 for x in out.body.iter():
  for k,v in x.attrib.items():
   if k.startswith('{'+N['r']+'}'):ck(v in out.rels,'relationship_exists:'+v)
 for key,path in [('tables','.//w:tbl'),('math','.//m:oMath'),('drawings','.//wp:inline')]:
  local=path.split(':')[-1]
  namespace=N['wp'] if key=='drawings' else N['m'] if key=='math' else N['w']
  preserved=sum(sum(1 for _ in node.iter('{'+namespace+'}'+local)) for node in prefix)
  generated_layouts=sum(row.get('figure_layout',{}).get('kind')=='inline-row' for row in r['mapping'] if row.get('figure_layout')) if key=='tables' else 0
  ck(len(out.body.findall(path,N))==counts[key]+preserved+generated_layouts,'total_'+key)
 promoted=sum(c['operation']=='promote_standalone_math' for row in r['mapping'] for c in row['changes'])
 ck(len(out.body.findall('.//m:oMathPara',N))==sum(len(Package(s['path']).body.findall('.//m:oMathPara',N)) for s in r['sources'])+promoted+sum(sum(1 for _ in node.iter(q('m:oMathPara'))) for node in prefix),'display_math_wrappers')
 from .verify_native_numbering import verify as verify_native_output
 from .forms_organization import verify_organization
 for error in verify_organization(r,out):ck(False,error)
 for error in verify_native_output(r,out):ck(False,error)
 for error in verify_layout_feedback(r,out):ck(False,error)
 if physical_contract:
  for error in verify_object_numbers(r,out,style_names,len(prefix)):ck(False,error)
 return {'output':r['output'],'status':'PASS' if not errors else 'FAIL','checks':count,'errors':errors,'counts':dict(counts),'source_details':details,'existing_unresolved_internal_links':original_unresolved,'exceptions':r['exceptions'],'wps_status':r['wps_status'],'visual_status':'SKIPPED_BY_USER','scope':'Declared source-bound content transformations, object/link/table integrity, explicit emphasis, template package preservation. Not semantic truth or WPS acceptance.'}


def verify_layout_feedback(report,output):
 """Recheck output scope and final properties without calling the formatter."""
 errors=[]
 def check(ok,name):
  if not ok:errors.append('layout_feedback:'+name)
 document=report.get('output_document')
 requests=report.get('layout_requests',[]);applied=report.get('layout_feedback',[])
 if not requests and not applied:return errors
 check(bool(document),'document_identity_present')
 check((document=='master')==bool(report['master']),'document_identity_matches_master')
 if not report['master']:check(document in report.get('selected_source_ids',[]),'document_identity_matches_source')
 expected=[];properties={}
 for request in requests:
  check(bool(request.get('document')),'request_has_document_scope')
  if request.get('document')!=document:continue
  for row in report['mapping']:
   if (request['source_id'],request['block_id'])!=(row['source_id'],row['block_id']):continue
   expected.append(dict(request,output_index=row['output_index'],content_unchanged=True))
   properties.setdefault(row['output_index'],{}).update(request['settings'])
 check(applied==expected,'applied_requests_exactly_match_selected_document')
 for index,settings in properties.items():
  p=output.body[index];check(p.tag==q('w:p'),str(index)+':paragraph')
  paths={'alignment':('w:jc','val',str),'first_line_pt':('w:ind','firstLine',lambda v:str(round(v*20))),
         'space_before_pt':('w:spacing','before',lambda v:str(round(v*20))),
         'space_after_pt':('w:spacing','after',lambda v:str(round(v*20))),
         'line_multiple':('w:spacing','line',lambda v:str(round(v*240)))}
  for key,(path,attr,convert) in paths.items():
   if key not in settings:continue
   node=p.find('w:pPr/'+path,N)
   check(node is not None and node.get(q('w:'+attr))==convert(settings[key]),str(index)+':'+key)
  if 'line_multiple' in settings:
   spacing=p.find('w:pPr/w:spacing',N);check(spacing is not None and spacing.get(q('w:lineRule'))=='auto',str(index)+':line_rule')
  if 'font_pt' in settings:
   size=str(round(settings['font_pt']*2))
   for run in p.xpath('.//w:r[not(ancestor::m:oMath)]',namespaces=N):
    for key in ('sz','szCs'):
     node=run.find('w:rPr/w:'+key,N);check(node is not None and node.get(q('w:val'))==size,str(index)+':font_'+key)
 return errors


def verify_source_frontmatter(source,scope):
 """Validate each exclusion against frozen XML, without importing the selector."""
 if not scope:return []
 errors=[]
 def check(ok,name):
  if not ok:errors.append('source_frontmatter:'+name)
 check(scope.get('contract')=='source-frontmatter-v1','known_contract')
 check(scope.get('source_sha256')==digest(source.path.read_bytes()),'source_hash')
 exclusions=scope.get('exclusions',[]);indices=[e['source_index'] for e in exclusions]
 check(len(indices)==len(set(indices)),'unique_excluded_blocks')
 title=scope.get('title');title_index=None
 if title:
  nodes=source.root.xpath(title['source_path'],namespaces=N)
  check(len(nodes)==1 and nodes[0].tag==q('w:p'),'title_source_paragraph')
  if len(nodes)==1:
   node=nodes[0];raw=ordinary(node);value=raw.strip()
   if value.startswith('【') and value.endswith('】'):value=value[1:-1].strip()
   check(raw==title['raw_text'] and value==title['text'],'exact_source_title_words')
   while node.getparent() is not source.body and node.getparent() is not None:node=node.getparent()
   check(node.getparent() is source.body,'title_is_in_body_not_header')
   if node.getparent() is source.body:title_index=source.body.index(node)
 styles=E.fromstring(source.parts['word/styles.xml']);style_names={s.get(q('w:styleId')):(s.find('w:name',N).get(q('w:val'),'')) for s in styles.findall('w:style',N) if s.find('w:name',N) is not None}
 labels=re.compile(r'^(?:部门|作者|编制人?|编写人?|审核人?|批准人?|审定人?|编制单位|编写单位|单位|项目组|报告编号|文件编号|编制日期|完成日期|日期|时间)(?:\s*[:：].*|\s+\S.*)?$')
 date=re.compile(r'^(?:\d{4}\s*年\s*(?:\d{1,2}\s*月\s*(?:\d{1,2}\s*日)?)?|\d{4}[-/.]\d{1,2}(?:[-/.]\d{1,2})?)$')
 headings={'目录','目次','contents','tableofcontents','图目录','表目录','插图目录','表格目录','listoffigures','listoftables'}
 def lines(n):
  answer=[]
  for p in n.iter(q('w:p')):
   value=''.join((x.text or '') if x.tag==q('w:t') else '\t' if x.tag==q('w:tab') else '' for x in p.iter() if x.tag in (q('w:t'),q('w:tab'))).strip()
   if value:answer.append((p,value))
  return answer
 toc_entries=0;cover=[]
 for e in exclusions:
  index=e['source_index'];check(type(index) is int and 0<=index<len(source.body),'source_index')
  if type(index) is not int or not 0<=index<len(source.body):continue
  node=source.body[index];values=lines(node)
  check(e.get('source_sha256')==scope['source_sha256'] and e.get('xml_sha256')==digest(E.tostring(node,encoding='utf8',with_tail=False)),'excluded_xml_binding:'+str(index))
  check(e['kind'] in ('source_cover','source_toc'),'allowed_exclusion_kind')
  check(not node.findall('.//m:oMath',N) and not any(n.get('equationxml') is not None or n.get('ProgID','').lower().startswith(('equation.','mathtype.')) for n in node.iter()),'no_math_removed_as_frontmatter:'+str(index))
  if e['kind']=='source_cover':
   cover.append(index);reason=e.get('evidence',[])
   metadata=sum(bool(labels.fullmatch(v) or date.fullmatch(v)) for p,v in values)
   if 'leading_cover_metadata_container' in reason:check(metadata>=2,'cover_container_metadata:'+str(index))
   elif 'leading_cover_metadata' in reason:check(bool(values) and metadata==len(values),'cover_metadata_only:'+str(index))
   elif 'leading_empty_page_boundary' in reason:check(not values and not checked_pictures(node),'empty_cover_boundary:'+str(index))
   elif 'leading_empty_cover_paragraph' in reason:
    check(node.tag==q('w:p') and not values and not checked_pictures(node),'empty_leading_cover_paragraph:'+str(index))
   elif 'leading_cover_institution_with_metadata' in reason:
    institution=bool(re.fullmatch(r'(?=.{2,60}$)[^。；;！？!?：:0-9]*(?:研究院|研究所|实验室|技术中心|研发中心|事业部|部门|有限公司)',ordinary(node).strip()))
    corroboration=any(labels.fullmatch(ordinary(n).strip()) or date.fullmatch(ordinary(n).strip()) for n in list(source.body)[:80])
    check(len(values)==1 and institution and corroboration and bool(title),'leading_institution_with_metadata:'+str(index))
   elif 'leading_cover_graphic' in reason:
    check(not values and bool(checked_pictures(node)) and bool(title),'cover_graphic_with_source_title:'+str(index))
   elif 'leading_metadata_value' in reason:
    previous=ordinary(source.body[index-1]).strip() if index else ''
    check(bool(labels.fullmatch(previous)) and previous.endswith(('：',':')) and len(ordinary(node))<30,'metadata_value_after_label:'+str(index))
   else:check(False,'unknown_cover_reason:'+str(index))
  else:
   check(not checked_pictures(node),'no_figure_removed_as_toc:'+str(index))
   for p,value in values:
    compact=re.sub(r'\s+','',value).casefold()
    if compact in headings or re.match(r'^(?:错误[!！：:]|Error[!:])',value,re.I):continue
    st=p.find('w:pPr/w:pStyle',N);name=style_names.get(st.get(q('w:val')) if st is not None else '','')
    entry=bool(re.search(r'(?:\t+|[.．·…]{2,}| {2,})\s*(?:[0-9]+|[ivxlcdmIVXLCDM]+|[一二三四五六七八九十]+)\s*$',value))
    if not entry:entry=bool(re.match(r'^(?:\d+(?:\.\d+)*|[一二三四五六七八九十百]+[、.．]|第[一二三四五六七八九十百0-9]+[章节篇]).*\s+(?:\d+|[ivxlcdmIVXLCDM]+)\s*$',value))
    styled=bool(re.match(r'^(?:toc|目录|目次)\s*[1-9]',name,re.I))
    check(entry or styled,'toc_row_has_navigation_evidence:'+str(index));toc_entries+=int(entry or styled)
 if any(e['kind']=='source_toc' for e in exclusions):check(toc_entries>=1,'toc_has_entries')
 if cover:
  for i,node in enumerate(source.body):
   if i>max(cover):break
   if i in indices or i==title_index or not ordinary(node).strip():continue
   check(False,'cover_exclusion_crosses_body:'+str(i))
 return errors


def verify_object_numbers(report,out,styles,prefix_count=0):
 """Inventory original XML independently; do not trust title counts as objects."""
 errors=[];expected={};observed={};caption_nodes={};bookmarks=bookmark_texts(out.body)
 def check(ok,name):
  if not ok:errors.append('objects:'+name)
 def path(root,node):
  steps=[]
  while node is not root:
   parent=node.getparent();name=E.QName(node).localname
   index=1+sum(c.tag==node.tag for c in list(parent)[:list(parent).index(node)])
   steps.insert(0,'w:'+name+'['+str(index)+']');node=parent
  return './'+'/'.join(steps) if steps else '.'
 for source in report['sources']:
  sid=source['source_id'];original=Package(source['path'])
  for row in [m for m in report['mapping'] if m['source_id']==sid]:
   root=original.body[row['source_index']]
   if row.get('figure_layout') and checked_pictures(root):items=[('图',root)]
   elif root.tag!=q('w:p'):
    items=[]
    for n in root.iter():
     if n.tag==q('w:tbl'):items.append(('表',n))
     elif n.tag==q('w:p') and checked_pictures(n):items.append(('图',n))
   elif root.tag==q('w:p') and checked_pictures(root):items=[('图',root)]
   else:items=[]
   for label,node in items:
    key=(sid,row['source_index'],path(root,node),label)
    expected[key]=(row,node,source)
 for obj in report.get('object_numbers',[]):
  key=(obj['source_id'],obj['source_index'],obj['source_path'],obj['label'])
  check(key not in observed,'one_number_per_object:'+str(key));observed[key]=obj
 check(set(expected)==set(observed),'all_physical_objects_numbered_no_phantoms')
 check(len({o['object_id'] for o in observed.values()})==len(observed),'stable_occurrence_ids_unique')
 for key,obj in observed.items():
  if key not in expected:continue
  row,original,source=expected[key];label=obj['label'];number=obj['number']
  try:
   block=out.body[obj['object_output_index']];targets=block.xpath(obj['object_path'],namespaces=N)
   caps=out.body[obj['caption_output_index']].xpath(obj['caption_path'],namespaces=N)
   check(len(targets)==len(caps)==1,'unique_object_and_caption_locator:'+obj['object_id'])
   physical,cap=targets[0],caps[0]
   check(obj['object_output_index']==row['output_index'],'object_mapping_locator:'+obj['object_id'])
   if label=='图':
    check(bool(checked_pictures(physical)),'figure_has_images:'+obj['object_id'])
    check(len(checked_pictures(original))==len(checked_pictures(physical))==obj['image_instances'],'group_preserves_every_image_instance:'+obj['object_id'])
   else:check(original.tag==physical.tag==q('w:tbl'),'data_table_physical_type:'+obj['object_id'])
   check(obj['source_sha256']==source['sha256'],'source_binding:'+obj['object_id'])
   check(cap not in caption_nodes,'one_caption_per_logical_object:'+obj['object_id']);caption_nodes[cap]=obj
   check(physical.getparent() is cap.getparent() and list(cap.getparent()).index(cap)==list(physical.getparent()).index(physical)+(1 if label=='图' else -1),'caption_adjacent_correct_side:'+obj['object_id'])
   field=checked_fields(cap)
   check(len(field)==1 and re.fullmatch(r'\s*SEQ\s+'+label+r'\s+\\\*\s+ARABIC\s*',field[0]['instruction'],re.I) is not None and field[0]['cache']==number,'native_object_sequence:'+obj['object_id'])
   ps=cap.find('w:pPr/w:pStyle',N)
   check(ps is not None and ps.get(q('w:val'))==styles['图题' if label=='图' else '表题'],'caption_template_style:'+obj['object_id'])
   check(bookmarks.get(obj['bookmark_name'])==number,'number_bookmark_binding:'+obj['object_id'])
   if obj['generated']:
    check(ordinary(cap)==label+' '+number,'bare_number_no_invented_title:'+obj['object_id'])
    check(not obj.get('caption_block_id'),'generated_caption_has_no_title_source:'+obj['object_id'])
    if obj['caption_path']=='.':check(any(b.get('kind')=='object_number_caption' and b.get('object_id')==obj['object_id'] and b['output_index']==obj['caption_output_index'] for b in report['generated_blocks']),'generated_caption_accounted:'+obj['object_id'])
   else:
    linked=next((m for m in report['mapping'] if m['source_id']==obj['source_id'] and m['block_id']==obj['caption_block_id']),None)
    check(linked is not None and linked['role']==('figure_caption' if label=='图' else 'table_caption') and linked.get('relates_to')==obj['block_id'] and linked['output_index']==obj['caption_output_index'],'title_bound_to_this_object:'+obj['object_id'])
  except (KeyError,IndexError,ValueError,TypeError) as exc:check(False,'invalid_object_record:'+obj.get('object_id','?')+':'+str(exc))
 counters=Counter();seen=[]
 for top in list(out.body)[prefix_count:]:
  for p in top.iter(q('w:p')):
   obj=caption_nodes.get(p)
   managed=[f for f in checked_fields(p) if re.match(r'\s*SEQ\s+[图表](?=\s|$)',f['instruction'],re.I)]
   if managed:check(obj is not None,'no_unbound_seq_advances_target_series')
   if obj:
    counters[obj['label']]+=1;check(obj['number']==str(counters[obj['label']]),'global_numbering_in_output_order:'+obj['object_id']);seen.append(obj['object_id'])
 check(len(seen)==len(expected),'every_object_caption_visible_once')
 return errors


def verify_retained_numbering(report, output):
 errors=[]
 declarations=report.get('numbering_retention',[])
 declared={(d['source_id'],d['source_num_id'],d['target_num_id']) for d in declarations}
 for row in report['mapping']:
  for change in row['changes']:
   if change.get('operation')=='preserve_automatic_list_numbering' and change.get('label_resolved') is False:
    if (row['source_id'],change['source_num_id'],change['target_num_id']) not in declared:
     errors.append('unresolved_numbering_without_preservation_declaration')
 if not declarations:return errors
 target_root=E.fromstring(output.parts['word/numbering.xml'])
 all_targets=[d['target_num_id'] for d in declarations]
 if len(all_targets)!=len(set(all_targets)):errors.append('retained_numbering_instances_must_not_be_shared')
 def definitions(root,nid):
  num=next((n for n in root.findall('w:num',N) if n.get(q('w:numId'))==nid),None)
  if num is None:raise ValueError('Missing numbered instance')
  aid=num.find('w:abstractNumId',N).get(q('w:val'))
  abstract=next((n for n in root.findall('w:abstractNum',N) if n.get(q('w:abstractNumId'))==aid),None)
  if abstract is None:raise ValueError('Missing abstract definition')
  a=deepcopy(abstract);n=deepcopy(num)
  a.set(q('w:abstractNumId'),'IDENTITY');n.set(q('w:numId'),'IDENTITY');n.find('w:abstractNumId',N).set(q('w:val'),'IDENTITY')
  return tree(a),tree(n)
 def usage(package,nid):
  styles={s.get(q('w:styleId')):s for s in E.fromstring(package.parts['word/styles.xml']).findall('w:style',N)}
  default=next((sid for sid,s in styles.items() if s.get(q('w:default'))=='1' and s.get(q('w:type'))=='paragraph'),'Normal')
  def effective(p,key):
   nodes=[p];st=p.find('w:pPr/w:pStyle',N);sid=st.get(q('w:val')) if st is not None else default;seen=set()
   while sid in styles and sid not in seen:
    seen.add(sid);n=styles[sid];nodes.append(n);base=n.find('w:basedOn',N);sid=base.get(q('w:val')) if base is not None else None
   for n in nodes:
    prop=n.find('w:pPr/w:numPr/w:'+key,N)
    if prop is not None:return prop.get(q('w:val'))
   return None
  entries=[]
  for p in package.body.iter(q('w:p')):
   if effective(p,'numId')==nid:
    owner=p
    while owner.getparent() is not package.body:owner=owner.getparent()
    ps=list(owner.iter(q('w:p')))
    entries.append((package.body.index(owner),ps.index(p),int(effective(p,'ilvl') or 0)))
  return entries
 for d in declarations:
  try:
   sid=d['source_id'];source_info=next(s for s in report['sources'] if s['source_id']==sid)
   source=Package(source_info['path']);root=E.fromstring(source.parts['word/numbering.xml'])
   if definitions(root,d['source_num_id'])!=definitions(target_root,d['target_num_id']):raise ValueError('Numbering definition changed beyond identity')
   mapping={m['source_index']:m for m in report['mapping'] if m['source_id']==sid}
   expected=[]
   for index,paragraph,level in usage(source,d['source_num_id']):
    row=mapping[index]
    if any(c['operation'] in ('heading_number','detach_automatic_heading_numbering') for c in row['changes']):raise ValueError('Preserved instance contains a renumbered heading')
    expected.append((row['output_index'],paragraph,level))
   if not expected or expected!=sorted(expected) or usage(output,d['target_num_id'])!=expected:raise ValueError('Complete numbering instance order or level changed')
  except Exception as exc:errors.append('retained_numbering:'+str(d)+':'+str(exc))
 return errors
