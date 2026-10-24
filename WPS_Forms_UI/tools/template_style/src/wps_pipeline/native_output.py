"""Native output primitives used by the existing source-bound assembler.

No semantic recognition, scheduling, model calls or Office automation here.
"""
from copy import deepcopy
import re
from .docx_formatting import NS, q, elt, ordinary_text

SEQ_FIELD = re.compile(r'\s*SEQ\s+\w+\s*(?:\\r\s+\d+\s*)?(?:\\\*\s+(?:ARABIC|MERGEFORMAT)\s*)*',re.I)


def number_properties(props, num_id, level=0):
    for old in list(props.findall('w:numPr', NS)):props.remove(old)
    np=elt('w:numPr');np.extend([elt('w:ilvl',val=level),elt('w:numId',val=num_id)])
    props.append(np)


def field_views(node):
    """Complete simple fields and nonnested complex fields in document order.

    Unsupported nested/malformed fields remain untouched by the assembler.
    Complex spans must consist of sibling runs within one paragraph.
    """
    result=[]
    for p in ([node] if node.tag==q('w:p') else node.findall('.//w:p',NS)):
        active=None;depth=0;local=[]
        for child in p:
            if child.tag==q('w:fldSimple'):
                if active is not None:active['nested']=True
                local.append(dict(simple=child,start=child,end=child,
                    instruction=child.get(q('w:instr'),''),cache_nodes=child.findall('.//w:t',NS),code_nodes=[]))
                continue
            if child.tag!=q('w:r'):continue
            for event in child:
                if event.tag==q('w:fldChar'):
                    kind=event.get(q('w:fldCharType'))
                    if kind=='begin':
                        depth+=1
                        if active is None:active=dict(simple=None,start=child,end=None,code_nodes=[],cache_nodes=[],separated=False,nested=False,marker=event)
                        else:active['nested']=True
                    elif kind=='separate' and active is not None and depth==1:active['separated']=True
                    elif kind=='end' and active is not None:
                        depth-=1
                        if depth==0:
                            active['end']=child
                            if active['separated'] and not active['nested']:
                                active['instruction']=''.join(x.text or '' for x in active['code_nodes']);local.append(active)
                            active=None
                elif active is not None:
                    if event.tag==q('w:instrText') and not active['separated']:active['code_nodes'].append(event)
                    elif event.tag==q('w:t') and active['separated']:active['cache_nodes'].append(event)
        local.sort(key=lambda f:p.index(f['start']))
        for f in local:
            f['paragraph']=p;f['cache']=''.join(t.text or '' for t in f['cache_nodes']);result.append(f)
    return result


def update_field(view,instruction,cache):
    if not view['cache_nodes']:raise ValueError('Native field has no cached result')
    if view['simple'] is not None:
        view['simple'].set(q('w:instr'),instruction);view['simple'].set(q('w:dirty'),'1')
    else:
        if not view['code_nodes']:raise ValueError('Native field has no instruction')
        view['code_nodes'][0].text=instruction
        for t in view['code_nodes'][1:]:t.text=''
        view['marker'].set(q('w:dirty'),'1')
    view['cache_nodes'][0].text=cache
    for t in view['cache_nodes'][1:]:t.text=''


def promote_text_field(p,start,end,instruction,cache):
    """Replace an exact plain run span with an editable field, without guessing."""
    offset=0
    for t in p.xpath('.//w:t[not(ancestor::m:oMath)]',namespaces=NS):
        text=t.text or '';left=offset;offset+=len(text)
        if not left<=start<offset:continue
        run=t.getparent()
        if end>offset or run.tag!=q('w:r') or run.getparent() is not p or any(c.tag not in (q('w:rPr'),q('w:t')) for c in run) or len(run.findall('w:t',NS))!=1:
            raise ValueError('Field promotion needs a single ordinary run span')
        at=p.index(run);before=text[:start-left];after=text[end-left:]
        p.remove(run)
        if before:
            prefix=deepcopy(run);prefix.find('w:t',NS).text=before;p.insert(at,prefix);at+=1
        f=elt('w:fldSimple',instr=instruction,dirty=1);r=elt('w:r')
        rp=run.find('w:rPr',NS)
        if rp is not None:r.append(deepcopy(rp))
        rt=elt('w:t');rt.text=cache;r.append(rt);f.append(r);p.insert(at,f);at+=1
        if after:
            suffix=deepcopy(run);suffix.find('w:t',NS).text=after;p.insert(at,suffix)
        return f
    raise ValueError('Native field span outside ordinary text')


def sequence_field(p):
    views=field_views(p)
    if len(views)!=1:return None
    f=views[0]
    if f['simple'] is not None and p.findall('.//w:fldChar',NS):return None
    if not SEQ_FIELD.fullmatch(f['instruction']) or not f['cache'].isdigit():return None
    prefix=re.match(r'^\s*[图表]\s*(\d+)',ordinary_text(p))
    return f if prefix and prefix[1]==f['cache'] else None


def number_bookmarks(p,number):
    active={};found=[]
    for n in p.iter():
        if n.tag==q('w:bookmarkStart'):active[n.get(q('w:id'))]=[n.get(q('w:name')),'']
        elif n.tag==q('w:bookmarkEnd'):
            old=active.pop(n.get(q('w:id')),None)
            if old and old[1].strip()==number:found.append(old[0])
        elif n.tag==q('w:t'):
            for value in active.values():value[1]+=n.text or ''
    return found


def wrap_number_bookmark(view,ident,name):
    p=view['paragraph'];at=p.index(view['start']);end=p.index(view['end'])
    start=elt('w:bookmarkStart',id=ident);start.set(q('w:name'),name);stop=elt('w:bookmarkEnd',id=ident)
    p.insert(end+1,stop);p.insert(at,start)


def manual_list_definition(root,forms,start_values):
    """Write one independent list instance from already confirmed group facts."""
    aid=str(max([int(n.get(q('w:abstractNumId'))) for n in root.findall('w:abstractNum',NS)]+[-1])+1)
    nid=str(max([int(n.get(q('w:numId'))) for n in root.findall('w:num',NS)]+[0])+1)
    abstract=elt('w:abstractNum',abstractNumId=aid);abstract.append(elt('w:multiLevelType',val='multilevel'))
    for level,form in sorted(forms.items()):
        label='%'+str(level+1);fmt='decimal'
        if form.startswith('integer_delimited:'):label+=form.split(':',1)[1]
        elif form=='parenthesized':label='（'+label+'）'
        elif form in ('chinese','chinese_space'):fmt='chineseCounting';label+='、' if form=='chinese' else ''
        elif form=='decimal_path':label='.'.join('%'+str(i+1) for i in range(level+1))
        elif form not in ('integer_space','integer_compact'):raise ValueError('Unsupported confirmed list marker')
        lvl=elt('w:lvl',ilvl=level)
        lvl.extend([elt('w:start',val=start_values.get(level,1)),elt('w:numFmt',val=fmt),elt('w:lvlText',val=label),elt('w:suff',val='space')])
        if level:lvl.append(elt('w:lvlRestart',val=level))
        pp=elt('w:pPr');pp.append(elt('w:ind',left=0,right=0,firstLine=0,hanging=0,firstLineChars=0,leftChars=0,rightChars=0,hangingChars=0));lvl.append(pp)
        abstract.append(lvl)
    root.insert(len(root.findall('w:abstractNum',NS)),abstract)
    num=elt('w:num',numId=nid);num.append(elt('w:abstractNumId',val=aid));root.append(num)
    return nid
