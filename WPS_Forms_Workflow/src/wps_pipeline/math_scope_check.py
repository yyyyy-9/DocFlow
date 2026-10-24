"""Independent scope projection for the finite text-math grammar.

Keeps fractions, roots, scripts and visible delimiters as distinct nodes;
never proves equivalence by flattening operands to an ambiguous string.
"""
import re
from lxml import etree as E

NS={'m':'http://schemas.openxmlformats.org/officeDocument/2006/math'}
SUB=dict(zip('₀₁₂₃₄₅₆₇₈₉₊₋₌₍₎ᵢⱼₙ','0123456789+-=()ijn'))
SUP=dict(zip('⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁼⁽⁾ⁿⁱ','0123456789+-=()ni'))


def text(value):
    return ''.join(value.split()).replace('−','-').replace('–','-').replace('×','*').replace('·','*')


def sequence(items):
    result=[]
    for item in items:
        if item[0]=='text':
            value=text(item[1])
            if not value:continue
            if result and result[-1][0]=='text':result[-1]=('text',result[-1][1]+value)
            else:result.append(('text',value))
        else:result.append(item)
    return result[0] if len(result)==1 else ('seq',tuple(result))


def source_scope(value):
    cursor=0
    def atom():
        nonlocal cursor
        if cursor>=len(value):raise ValueError('Missing source operand')
        c=value[cursor];cursor+=1
        if c in '([{|':
            end={'(':')','[':']','{':'}','|':'|'}[c]
            out=('delimiter',c,end,group(end))
        elif c in ')]}':raise ValueError('Unexpected delimiter')
        elif c=='√':out=('root',atom())
        else:
            token=re.match(r'[A-Za-z\u0370-\u03ff\u2100-\u214f∂∆Δ∞0-9]+(?:\.[0-9]+)?',value[cursor-1:])
            if token:cursor+=len(token.group())-1;out=('text',token.group())
            else:out=('text',c)
        while cursor<len(value):
            c=value[cursor]
            if c in '_^':
                cursor+=1
                if cursor==len(value):raise ValueError('Missing script')
                if value[cursor]=='{':cursor+=1;arg=group('}')
                else:arg=('text',value[cursor]);cursor+=1
                out=('sub' if c=='_' else 'sup',out,arg)
            elif c in SUB or c in SUP:
                mapping=SUB if c in SUB else SUP;parts=[]
                while cursor<len(value) and value[cursor] in mapping:
                    parts.append(mapping[value[cursor]]);cursor+=1
                out=('sub' if mapping is SUB else 'sup',out,('text',''.join(parts)))
            else:break
        return out
    def group(close=None):
        nonlocal cursor
        items=[]
        while cursor<len(value):
            c=value[cursor]
            if close and c==close:cursor+=1;return sequence(items)
            if c=='/':
                # Whitespace-separated implicit products remain a literal slash.
                if re.match(r'\s*[A-Za-z0-9α-ωΑ-Ω.]+\s+[A-Za-zα-ωΑ-Ω]',value[cursor+1:]):
                    items.append(('text','/'));cursor+=1;continue
                while items and items[-1][0]=='text' and not items[-1][1].strip():items.pop()
                if not items:raise ValueError('Missing numerator')
                cursor+=1
                while cursor<len(value) and value[cursor].isspace():cursor+=1
                if cursor<len(value) and value[cursor] in '+-−':
                    sign=value[cursor];cursor+=1;den=sequence([('text',sign),atom()])
                else:den=atom()
                items.append(('fraction',items.pop(),den))
            else:items.append(atom())
        if close:raise ValueError('Unclosed delimiter')
        return sequence(items)
    def normalized(node):
        if node[0]=='text':return ('text',text(node[1]))
        if node[0]=='seq':return sequence([normalized(x) for x in node[1]])
        return tuple(normalized(x) if isinstance(x,tuple) else x for x in node)
    return normalized(group())


def omml_scope(node):
    tag=E.QName(node).localname
    def child(name):
        el=node.find('m:'+name,NS)
        if el is None:raise ValueError('Missing OMML operand '+name)
        return omml_scope(el)
    if tag in ('oMath','e','sub','sup','num','den','deg'):
        return sequence([omml_scope(c) for c in node])
    if tag=='r':return ('text',text(''.join(node.xpath('./m:t/text()',namespaces=NS))))
    if tag=='f':return ('fraction',child('num'),child('den'))
    if tag in ('sSub','sSup'):return ('sub' if tag=='sSub' else 'sup',child('e'),child('sub' if tag=='sSub' else 'sup'))
    if tag=='rad':
        if child('deg')!=('seq',()):raise ValueError('Indexed roots outside finite grammar')
        return ('root',child('e'))
    if tag=='d':
        pr=node.find('m:dPr',NS)
        def prop(k):return pr.find('m:'+k,NS).get('{'+NS['m']+'}val')
        return ('delimiter',prop('begChr'),prop('endChr'),child('e'))
    raise ValueError('Unsupported generated math node '+tag)


def check(raw,math):
    return source_scope(raw)==omml_scope(math)
