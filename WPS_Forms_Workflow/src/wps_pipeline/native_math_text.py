"""Bounded literal text -> native OMML. No algebra, evaluation or LLM calls.

Only explicit relations and explicit scripts in prose are candidates. A literal
parser retains operator order, grouping and text; ambiguity is returned locally.
"""
import re,json,hashlib,argparse
from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED
from copy import deepcopy
from lxml import etree as E
from .math_scope_check import check as check_scope

W='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
M='http://schemas.openxmlformats.org/officeDocument/2006/math'
N={'w':W,'m':M};q=lambda x:'{'+N[x.split(':')[0]]+'}'+x.split(':')[1]
SUB=str.maketrans('₀₁₂₃₄₅₆₇₈₉₊₋₌₍₎ᵢⱼₙ','0123456789+-=()ijn')
SUP=str.maketrans('⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁼⁽⁾ⁿⁱ','0123456789+-=()ni')
SUBCH='₀₁₂₃₄₅₆₇₈₉₊₋₌₍₎ᵢⱼₙ';SUPCH='⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁼⁽⁾ⁿⁱ'
CHARS=r'A-Za-z0-9\u0370-\u03ff\u2100-\u214f₀-₉₊₋₌₍₎ᵢⱼₙ⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁼⁽⁾ⁿⁱ=≈≃≡≤≥<>∝+−–*/×·^_(){}\[\]|√∞∑Σ∫∂∆Δπ°%.,;′†∘∈∀∇± \'−'
REGION=re.compile('['+CHARS+']+')
RELATION=re.compile(r'[=≈≃≡≤≥∝]')
ATOM=re.compile(r'[A-Za-z\u0370-\u03ff\u2100-\u214f∂∆Δ∞0-9]+(?:\.[0-9]+)?')
def sha(b):return hashlib.sha256(b).hexdigest()
def rawtext(n):return ''.join(n.xpath('.//w:t[not(ancestor::m:oMath)]/text()',namespaces=N))
def canonical(s):
    # Independent verifier has its own reader; this only normalizes source
    # notation (spaces, minus variants and explicit script syntax).
    s=s.replace('−','-').replace('–','-').replace('×','*').replace('·','*')
    s=re.sub('['+SUBCH+']+',lambda m:'_{'+m.group().translate(SUB)+'}',s)
    s=re.sub('['+SUPCH+']+',lambda m:'^{'+m.group().translate(SUP)+'}',s)
    s=re.sub(r'([_^])([A-Za-z0-9α-ωΑ-Ω])',r'\1{\2}',s)
    return re.sub(r'\s+','',s)

def candidates(text):
    result=[]
    for match in REGION.finditer(text):
        value=match.group();left=len(value)-len(value.lstrip(' .,'));s=value.strip(' .,')
        if len(s)<3 or not re.search(r'[A-Za-zα-ωΑ-Ω\u2100-\u214f]',s):continue
        relation=RELATION.search(s)
        if s.startswith('/') or (not relation and re.fullmatch(r'(?:μm|mm|nm|pm|cm|m)/(?:min|s|ms|μs|mm|nm|m)',s)):continue
        chemical=s.translate(SUB)
        if not relation and re.fullmatch(r'(?:[A-Z][a-z]?\d*)+(?:/(?:[A-Z][a-z]?\d*)+)*',chemical):continue
        explicit_operation=bool(re.search(r'[A-Za-zα-ωΑ-Ω]\s*/\s*\(|(?:sin|cos|tan|exp|ln|log|asin|arcsin)\s*\(',s)) or ('/' in s and (bool(re.search('[α-ωΑ-Ω]',s)) or bool(re.fullmatch(r'[A-Za-z]{1,2}/[A-Za-z]{1,2}',s))))
        if not relation and not explicit_operation and not re.search(r'[_^]|['+SUBCH+SUPCH+']',s):continue
        if len(s)>400 or re.search(r'https?|www\.|\.json|\.csv|\.py\b|\.docx|\.png|\.zip|--\w',s,re.I):continue
        # Underscored software labels/path names are not equations.
        if not relation and re.search(r'[A-Za-z0-9]{3,}_|_[A-Za-z]{3,}',s):continue
        if not relation and not explicit_operation and not re.search(r'[α-ωΑ-Ω\u2100-\u214f_^]',s):continue
        if relation and not s[:relation.start()].strip():continue
        # Decimal relation operands must not capture an adjacent sentence period.
        result.append((match.start()+left,match.start()+left+len(s),s))
    return result

class Parser:
    def __init__(self,s):self.s=s;self.i=0
    def seq(self,stop=None):
        out=[]
        while self.i<len(self.s) and (stop is None or self.s[self.i]!=stop):
            if self.s[self.i]=='/':
                # Preserve ambiguous unit/implicit-product division literally;
                # never invent the scope of a denominator from typography.
                tail=self.s[self.i+1:]
                if re.match(r'\s*[A-Za-z0-9α-ωΑ-Ω.]+\s+[A-Za-zα-ωΑ-Ω]',tail):
                    out.append(('text','/'));self.i+=1;continue
                self.i+=1
                while out and out[-1][0]=='text' and out[-1][1].isspace():out.pop()
                if not out:raise ValueError('Missing fraction numerator')
                while self.i<len(self.s) and self.s[self.i].isspace():self.i+=1
                if self.i<len(self.s) and self.s[self.i] in '+-−':
                    sign=self.s[self.i];self.i+=1;den=('seq',[('text',sign),self.atom()])
                else:den=self.atom()
                out.append(('frac',out.pop(),den));continue
            out.append(self.atom())
        if stop is not None:
            if self.i>=len(self.s):raise ValueError('Unclosed delimiter')
            self.i+=1
        return ('seq',out)
    def atom(self):
        if self.i>=len(self.s):raise ValueError('Missing operand')
        ch=self.s[self.i];self.i+=1
        if ch in '([{|':
            close={'(':')','[':']','{':'}','|':'|'}[ch];base=('group',ch,close,self.seq(close))
        elif ch in ')]}':raise ValueError('Unmatched closing delimiter')
        elif ch=='√':base=('root',self.atom())
        else:
            a=ATOM.match(self.s,self.i-1)
            if a:base=('text',a.group());self.i=a.end()
            else:base=('text',ch)
        while self.i<len(self.s):
            c=self.s[self.i]
            if c in '_^':
                self.i+=1
                if self.i>=len(self.s):raise ValueError('Missing script')
                if self.s[self.i]=='{':self.i+=1;arg=self.seq('}')
                else:arg=('text',self.s[self.i]);self.i+=1
                base=('sub' if c=='_' else 'sup',base,arg)
            elif c in SUBCH+SUPCH:
                typ='sub' if c in SUBCH else 'sup';chars=SUBCH if typ=='sub' else SUPCH;start=self.i
                while self.i<len(self.s) and self.s[self.i] in chars:self.i+=1
                val=self.s[start:self.i].translate(SUB if typ=='sub' else SUP);base=(typ,base,('text',val))
            else:break
        return base

def emit(ast,rpr=None):
    kind=ast[0]
    if kind=='seq':return [n for child in ast[1] for n in emit(child,rpr)]
    if kind=='text':
        r=E.Element(q('m:r'))
        if rpr is not None:r.append(deepcopy(rpr))
        t=E.SubElement(r,q('m:t'));t.text=ast[1];t.set('{http://www.w3.org/XML/1998/namespace}space','preserve');return [r]
    tag={'group':'d','root':'rad','frac':'f','sub':'sSub','sup':'sSup'}[kind];n=E.Element(q('m:'+tag))
    def put(name,child):
        el=E.SubElement(n,q('m:'+name));el.extend(emit(child,rpr))
    if kind=='group':
        pr=E.SubElement(n,q('m:dPr'))
        for key,val in [('begChr',ast[1]),('endChr',ast[2])]:E.SubElement(pr,q('m:'+key)).set(q('m:val'),val)
        put('e',ast[3])
    elif kind=='root':
        pr=E.SubElement(n,q('m:radPr'));E.SubElement(pr,q('m:degHide')).set(q('m:val'),'1');E.SubElement(n,q('m:deg'));put('e',ast[1])
    elif kind=='frac':put('num',ast[1]);put('den',ast[2])
    else:put('e',ast[1]);put('sub' if kind=='sub' else 'sup',ast[2])
    return [n]

def make_math(s,rpr=None):
    parser=Parser(s);ast=parser.seq();m=E.Element(q('m:oMath'));m.extend(emit(ast,rpr));return m

def groups(p):
    group=[]
    for child in p:
        eligible=child.tag==q('w:r') and all(c.tag in (q('w:rPr'),q('w:t')) for c in child)
        if eligible:group.append(child)
        else:
            if group:yield group;group=[]
    if group:yield group

def convert(source,dest,plan_only=False,protected_prefix_blocks=0):
    source=Path(source);dest=Path(dest)
    with ZipFile(source) as z:parts={n:z.read(n) for n in z.namelist()}
    root=E.fromstring(parts['word/document.xml']);records=[];pending=[];promotions=[]
    body=root.find('w:body',N)
    if type(protected_prefix_blocks) is not int or not 0<=protected_prefix_blocks<len(body):
        raise ValueError('Invalid protected template prefix')
    protected_paragraphs={p for node in list(body)[:protected_prefix_blocks] for p in node.iter(q('w:p'))}
    styles=E.fromstring(parts['word/styles.xml'])
    style_names={s.get(q('w:styleId')):s.find('w:name',N).get(q('w:val'),'').casefold().replace(' ','')
                 for s in styles.findall('w:style',N) if s.find('w:name',N) is not None}
    protected_ids={sid for sid,name in style_names.items()
                   if re.fullmatch(r'(?:heading|标题)[1-9]',name) or name in ('title','参考条目','图题','表题')}
    for pi,p in enumerate(root.findall('.//w:body//w:p',N)):
        if p in protected_paragraphs:continue
        if any(n.tag in (q('w:pict'),q('w:object'),q('w:txbxContent')) for n in p.iterancestors()):continue
        style=p.find('w:pPr/w:pStyle',N);sid=style.get(q('w:val'),'') if style is not None else ''
        if sid in protected_ids:continue
        for group in list(groups(p)):
            text=''.join(rawtext(r) for r in group);spans=candidates(text)
            if not spans:continue
            if p.getparent().tag==q('w:body') and len(spans)==1 and spans[0][0]==0 and text==rawtext(p) and re.fullmatch(r'[.,;。；，\s]*',text[spans[0][1]:]):
                # A standalone equation's final punctuation stays within its
                # native display object, rather than forcing an inline wrapper.
                spans=[(0,len(text),text)]
            props=[E.tostring(r.find('w:rPr',N)) if r.find('w:rPr',N) is not None else b'' for r in group]
            if len(set(props))!=1:
                pending.extend({'paragraph':pi,'text':s,'reason':'mixed_character_format'} for _,_,s in spans);continue
            base=group[0];rpr=base.find('w:rPr',N);nodes=[];cursor=0;made=[]
            def run(t):
                if not t:return
                r=E.Element(q('w:r'))
                if rpr is not None:r.append(deepcopy(rpr))
                el=E.SubElement(r,q('w:t'));el.text=t;el.set('{http://www.w3.org/XML/1998/namespace}space','preserve');nodes.append(r)
            for a,b,raw in spans:
                try:
                    math=make_math(raw,rpr)
                    if not check_scope(raw,math):raise ValueError('independent_math_scope_mismatch')
                except ValueError as err:
                    pending.append({'paragraph':pi,'text':raw,'reason':str(err)});continue
                run(text[cursor:a]);nodes.append(math);cursor=b
                made.append({'paragraph':pi,'source_group_text':text,'group_start':a,'group_end':b,'text':raw,'canonical_source':canonical(raw),'_math':math})
            if made:
                run(text[cursor:]);index=p.index(group[0])
                for r in group:p.remove(r)
                for offset,n in enumerate(nodes):p.insert(index+offset,n)
                records.extend(made)
    for r in records:
        p=root.findall('.//w:body//w:p',N)[r['paragraph']];m=r.pop('_math');r['math_index_in_paragraph']=p.findall('.//m:oMath',N).index(m);r['native_math_sha256']=sha(E.tostring(m))
    formula_style=next(s.get(q('w:styleId')) for s in styles.findall('w:style',N) if s.find('w:name',N) is not None and s.find('w:name',N).get(q('w:val'))=='公式')
    for pi in sorted({r['paragraph'] for r in records}):
        p=root.findall('.//w:body//w:p',N)[pi]
        if p.getparent().tag==q('w:body') and not rawtext(p).strip() and not p.findall('m:oMathPara',N) and p.findall('m:oMath',N):
            ms=p.findall('m:oMath',N);at=p.index(ms[0]);mp=E.Element(q('m:oMathPara'));pr=E.SubElement(mp,q('m:oMathParaPr'));E.SubElement(pr,q('m:jc')).set(q('m:val'),'center')
            for m in ms:p.remove(m);mp.append(m)
            p.insert(at,mp);pp=p.find('w:pPr',N)
            for key,val in [('pStyle',formula_style),('keepLines','1'),('keepNext','0')]:
                e=pp.find('w:'+key,N)
                if e is None:e=E.SubElement(pp,q('w:'+key))
                e.set(q('w:val'),val)
            promotions.append(pi)
    report={'version':'0.1.0-experimental','source':str(source.resolve()),'source_sha256':sha(source.read_bytes()),'output':str(dest.resolve()),'conversions':records,'new_display_paragraphs':promotions,'pending':pending,'protected_prefix_blocks':protected_prefix_blocks,'script_model_calls':0,'scope':'Literal explicit mathematical expressions in report body, no inferred subscripts or algebra changes; template prefix preserved.'}
    if not plan_only:
        parts['word/document.xml']=E.tostring(root,encoding='UTF-8',xml_declaration=True,standalone=True)
        dest.parent.mkdir(parents=True,exist_ok=True)
        with ZipFile(dest,'x',ZIP_DEFLATED) as z:
            for n,data in parts.items():z.writestr(n,data)
        report['output_sha256']=sha(dest.read_bytes())
    with dest.with_suffix('.math.json').open('x',encoding='utf8') as f:json.dump(report,f,ensure_ascii=False,indent=2)
    return report
