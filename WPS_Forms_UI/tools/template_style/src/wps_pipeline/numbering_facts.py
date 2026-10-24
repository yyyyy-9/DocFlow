"""OOXML numbering observations. No heading/list classification."""
import re
from lxml import etree as E

W='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS={'w':W}
def value(node,path,default=None):
    n=node.find(path,NS) if node is not None else None
    return n.get('{'+W+'}val',default) if n is not None else default


class NumberingFacts:
    def __init__(self,parts):
        styles=E.fromstring(parts['word/styles.xml'])
        self.styles={s.get('{'+W+'}styleId'):s for s in styles.findall('w:style',NS)}
        self.default=next((sid for sid,s in self.styles.items() if s.get('{'+W+'}default')=='1' and s.get('{'+W+'}type')=='paragraph'),'Normal')
        self.root=E.fromstring(parts['word/numbering.xml']) if 'word/numbering.xml' in parts else E.Element('{'+W+'}numbering')
        self.nums={n.get('{'+W+'}numId'):n for n in self.root.findall('w:num',NS)}
        self.abstracts={n.get('{'+W+'}abstractNumId'):n for n in self.root.findall('w:abstractNum',NS)}
        self.counters={}

    def inherited(self,p,path):
        result=value(p,'w:pPr/'+path)
        if result is not None:return result
        sid=value(p,'w:pPr/w:pStyle',self.default);seen=set()
        while sid and sid not in seen:
            seen.add(sid);s=self.styles.get(sid);result=value(s,'w:pPr/'+path)
            if result is not None:return result
            sid=value(s,'w:basedOn')
        return None

    def inspect(self,p):
        nid=self.inherited(p,'w:numPr/w:numId');ilvl=self.inherited(p,'w:numPr/w:ilvl') or '0'
        result={'num_id':nid,'ilvl':int(ilvl),'status':'none' if nid in (None,'0') else 'unresolved',
                'label':None,'values':None,'definition':None}
        if result['status']=='none':return result
        num=self.nums.get(nid);aid=value(num,'w:abstractNumId');abstract=self.abstracts.get(aid)
        if abstract is None:return result
        result['abstract_id']=aid
        overrides={n.get('{'+W+'}ilvl'):n for n in num.findall('w:lvlOverride',NS)}
        levels={n.get('{'+W+'}ilvl'):n for n in abstract.findall('w:lvl',NS)}
        defs={}
        for key,lvl in levels.items():
            override=overrides.get(key)
            replacement=override.find('w:lvl',NS) if override is not None else None
            if replacement is not None:lvl=replacement
            defs[int(key)]={'format':value(lvl,'w:numFmt'),'text':value(lvl,'w:lvlText'),
                            'start':int(value(override,'w:startOverride',value(lvl,'w:start','1'))),
                            'suffix':value(lvl,'w:suff','tab'),'restart':value(lvl,'w:lvlRestart')}
        level=int(ilvl);d=defs.get(level);result['definition']=d
        if d is None or level>8:return result
        # Restart follows an ancestor's occurrence, including a transitive reset.
        # A never-restart level (0) does not inherit its parent's reset.
        if abstract.find('w:numStyleLink',NS) is not None or abstract.find('w:styleLink',NS) is not None or any(v['restart'] is not None and (not v['restart'].isdigit() or int(v['restart'])>k) for k,v in defs.items()):return result
        counters=self.counters.setdefault(nid,{})
        counters[level]=counters.get(level,d['start']-1)+1
        dirty={level}
        for key in sorted(k for k in defs if k>level):
            restart=defs[key]['restart']
            trigger=key-1 if restart is None else int(restart)-1
            if restart!='0' and trigger in dirty:dirty.add(key)
        for key in dirty-{level}:counters.pop(key,None)
        if d['format']=='bullet' and d['text'] and '%' not in d['text']:
            result.update(status='resolved',label=d['text'],values=[counters[level]])
            return result
        if not d['text']:return result
        refs=[int(x)-1 for x in re.findall(r'%(\d)',d['text'])]
        rendered={k:format_ordinal(counters.get(k,defs.get(k,{}).get('start',1)),defs.get(k,{}).get('format')) for k in refs}
        if any(v is None for v in rendered.values()):return result
        label=re.sub(r'%(\d)',lambda m:rendered[int(m.group(1))-1],d['text'])
        result.update(status='resolved',label=label,values=[counters.get(k,defs[k]['start']) for k in range(level+1)] if all(k in defs for k in range(level+1)) else [counters[level]])
        return result


def format_ordinal(number,fmt):
    if fmt=='decimal':return str(number)
    if fmt=='decimalZero':return str(number).zfill(2)
    if fmt in ('lowerLetter','upperLetter') and number>0:
        text='';value=number
        while value>0:value,rem=divmod(value-1,26);text=chr(97+rem)+text
        return text.upper() if fmt=='upperLetter' else text
    if fmt in ('lowerRoman','upperRoman') and number>0:
        value=number;text=''
        for amount,symbol in ((1000,'M'),(900,'CM'),(500,'D'),(400,'CD'),(100,'C'),(90,'XC'),(50,'L'),(40,'XL'),(10,'X'),(9,'IX'),(5,'V'),(4,'IV'),(1,'I')):
            while value>=amount:text+=symbol;value-=amount
        return text.lower() if fmt=='lowerRoman' else text
    if fmt=='chineseCounting' and 0<number<100:
        digits='零一二三四五六七八九';tens,ones=divmod(number,10)
        return (('' if tens==1 else digits[tens])+'十' if tens else '')+(digits[ones] if ones else '')
    return None
