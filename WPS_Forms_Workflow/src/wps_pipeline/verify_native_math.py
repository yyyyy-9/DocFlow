"""Independent decoder and source-bound inverse check of newly authored math.

Does not import the math parser/writer. No expression is evaluated or simplified.
"""
from pathlib import Path
from copy import deepcopy
from collections import defaultdict
import json,re,hashlib
from lxml import etree as E
from .verify_report_build import Package,N,q,tree,ordinary,verify as verify_assembly
from .math_scope_check import check as check_scope

def source_notation(s):
    subs=dict(zip('₀₁₂₃₄₅₆₇₈₉₊₋₌₍₎ᵢⱼₙ','0123456789+-=()ijn'))
    sups=dict(zip('⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁼⁽⁾ⁿⁱ','0123456789+-=()ni'))
    result=[];i=0
    while i<len(s):
        ch=s[i]
        if ch in subs or ch in sups:
            mapping=subs if ch in subs else sups;kind='_' if ch in subs else '^';parts=[]
            while i<len(s) and s[i] in mapping:parts.append(mapping[s[i]]);i+=1
            result.append(kind+'{'+''.join(parts)+'}');continue
        if ch in '_^' and i+1<len(s) and s[i+1]!='{':result.append(ch+'{'+s[i+1]+'}');i+=2;continue
        result.append(ch);i+=1
    return ''.join(''.join(result).split()).replace('−','-').replace('–','-').replace('×','*').replace('·','*')

def decode(n):
    tag=E.QName(n).localname
    def child(name):
        c=n.find('m:'+name,N)
        if c is None:raise ValueError('Missing math child '+name)
        return decode(c)
    if tag in ('oMath','e','sub','sup','num','den','deg'):return ''.join(decode(c) for c in n)
    if tag=='r':
        if any(c.tag not in (q('m:t'),q('w:rPr')) for c in n):raise ValueError('Unexpected math run')
        return ''.join(n.xpath('./m:t/text()',namespaces=N))
    if tag=='f':return child('num')+'/'+child('den')
    if tag=='sSub':return child('e')+'_{'+child('sub')+'}'
    if tag=='sSup':return child('e')+'^{'+child('sup')+'}'
    if tag=='d':
        pr=n.find('m:dPr',N);a=pr.find('m:begChr',N).get(q('m:val'));b=pr.find('m:endChr',N).get(q('m:val'))
        return a+child('e')+b
    if tag=='rad':
        if child('deg'):raise ValueError('Unexpected indexed root')
        return '√'+child('e')
    raise ValueError('Unrecognized generated math '+tag)

def verify(math_report,assembly_report=None):
    r=json.loads(Path(math_report).read_text(encoding='utf8'));src=Package(r['source']);out=Package(r['output']);errors=[];checks=0
    def ck(ok,name):
        nonlocal checks
        checks+=1
        if not ok:errors.append(name)
    ck(hashlib.sha256(src.path.read_bytes()).hexdigest()==r['source_sha256'],'input_hash')
    ck(hashlib.sha256(out.path.read_bytes()).hexdigest()==r['output_sha256'],'output_hash')
    ck(set(src.parts)==set(out.parts) and all(src.parts[k]==out.parts[k] for k in src.parts if k!='word/document.xml'),'only_document_part_changed')
    before=src.body.findall('.//w:p',N);after=out.body.findall('.//w:p',N);ck(len(before)==len(after),'paragraph_count')
    style_names={s.get(q('w:styleId')):s.find('w:name',N).get(q('w:val'),'').casefold().replace(' ','')
                 for s in E.fromstring(src.parts['word/styles.xml']).findall('w:style',N) if s.find('w:name',N) is not None}
    restored=deepcopy(out.body);inverse=restored.findall('.//w:p',N);bypara=defaultdict(list)
    protected_prefix=0
    if assembly_report:
        assembly=json.loads(Path(assembly_report).read_text(encoding='utf8'))
        template=Package(assembly['template'])
        slots=template.body.xpath('./w:sdt[w:sdtPr/w:tag[@w:val="wps-report-body-v1"]]',namespaces=N)
        anchors=template.body.xpath('.//w:bookmarkStart[@w:name="WPS_ReportBody"]',namespaces=N)
        if anchors:
            mark=anchors[0];paragraph=mark.getparent();ident=mark.get(q('w:id'))
            starts=template.body.xpath('.//w:bookmarkStart[@w:id=$ident]',ident=ident,namespaces=N)
            ends=template.body.xpath('.//w:bookmarkEnd[@w:id=$ident]',ident=ident,namespaces=N)
            valid=(not slots and len(anchors)==1 and ident is not None and len(starts)==len(ends)==1
                   and paragraph.tag==q('w:p') and paragraph.getparent() is template.body
                   and template.body[-2] is paragraph and ends[0].getparent() is paragraph
                   and list(paragraph).index(mark)<list(paragraph).index(ends[0]) and not ordinary(paragraph).strip())
            ck(valid,'unambiguous_template_body_bookmark')
            if valid:protected_prefix=len(template.body)-2
        elif len(slots)==1 and template.body[-2] is slots[0]:protected_prefix=len(template.body)-2
    ck(r.get('protected_prefix_blocks',0)==protected_prefix,'template_prefix_scope_matches_frozen_template')
    ck(all(tree(a)==tree(b) for a,b in zip(list(src.body)[:protected_prefix],list(out.body)[:protected_prefix])),'math_does_not_change_template_frontmatter')
    for c in r['conversions']:bypara[c['paragraph']].append(c)
    promoted=set(r.get('new_display_paragraphs',[]));ck(len(promoted)==len(r.get('new_display_paragraphs',[])),'unique_promotions')
    for pi,rows in bypara.items():
        ck(not any(n.tag in (q('w:pict'),q('w:object'),q('w:txbxContent')) for n in before[pi].iterancestors()),f'{pi}:no_rewrite_inside_preserved_drawing_or_ole_story')
        pstyle=before[pi].find('w:pPr/w:pStyle',N)
        name=style_names.get(pstyle.get(q('w:val')) if pstyle is not None else None,'')
        ck(not (re.fullmatch(r'(?:heading|标题)[1-9]',name) or name in ('title','参考条目','图题','表题')),f'{pi}:no_math_rewrite_in_protected_style')
        used=set();ms=inverse[pi].findall('.//m:oMath',N)
        for c in rows:
            idx=c['math_index_in_paragraph'];ck(idx not in used,f'{pi}:unique_math_index');used.add(idx)
            a,b=c['group_start'],c['group_end'];raw=c['text'];group=c['source_group_text']
            ck(group[a:b]==raw and group in ordinary(before[pi]),f'{pi}:source_span')
            math=ms[idx]
            try:ck(source_notation(decode(math))==source_notation(raw),f'{pi}:native_math_decodes_to_source')
            except Exception as e:ck(False,f'{pi}:decode:{e}')
            try:ck(check_scope(raw,math),f'{pi}:independent_operand_scope')
            except Exception as e:ck(False,f'{pi}:scope:{e}')
            run=E.Element(q('w:r'));t=E.SubElement(run,q('w:t'));t.text=raw;math.getparent().replace(math,run)
        if pi in promoted:
            p=inverse[pi];mp=p.find('m:oMathPara',N)
            ck(mp is not None and not ordinary(after[pi]).strip(),f'{pi}:standalone_display')
            if mp is not None:
                at=p.index(mp)
                for child in list(mp):
                    if child.tag!=q('m:oMathParaPr'):p.insert(at,child);at+=1
                p.remove(mp)
            original_pp=before[pi].find('w:pPr',N);p.remove(p.find('w:pPr',N));p.insert(0,deepcopy(original_pp))
    def attrs(run):
        values=[]
        for key in ('b','i','u','strike','dstrike','vertAlign'):
            e=run.find('w:rPr/w:'+key,N)
            if e is not None and e.get(q('w:val'),'1') not in ('0','false','off','none'):values.append((key,e.get(q('w:val'),'1')))
        return values
    def emphasis(p,new_math):
        result=[]
        def walk(n):
            if n.tag==q('m:oMath'):
                if n in new_math:
                    runs=n.findall('.//m:r',N);properties=attrs(runs[0]) if runs else []
                    ck(all(attrs(x)==properties for x in runs),'new_math_emphasis_consistent')
                    result.extend((ch,properties) for ch in new_math[n])
                return
            if n.tag==q('w:r'):
                result.extend((ch,attrs(n)) for ch in ''.join(n.xpath('./w:t/text()',namespaces=N)));return
            for c in n:walk(c)
        walk(p);return result
    for pi,(a,b) in enumerate(zip(before,after)):
        ms=b.findall('.//m:oMath',N);new={ms[c['math_index_in_paragraph']]:c['text'] for c in bypara.get(pi,[])}
        ck(emphasis(a,{})==emphasis(b,new),f'{pi}:character_emphasis')
    ck(src.content(src.body)==out.content(restored),'inverse_whole_body_content')
    # Package.content expects individual p/table containers, so also compare
    # every top-level block explicitly, including all old native math objects.
    ck(len(src.body)==len(restored),'body_block_count')
    for i,(a,b) in enumerate(zip(src.body,restored)):
        ck(src.content(a)==out.content(b),f'{i}:source_bound_content')
    for i,(a,b) in enumerate(zip(before,inverse)):
        ap,bp=a.find('w:pPr',N),b.find('w:pPr',N)
        ck((tree(ap) if ap is not None else None)==(tree(bp) if bp is not None else None),f'{i}:paragraph_format')
    count_old=len(src.body.findall('.//m:oMath',N));count_new=len(out.body.findall('.//m:oMath',N))
    ck(count_new==count_old+len(r['conversions']),'old_plus_new_math_count')
    report_nodes=list(out.body)[protected_prefix:-1]
    retained_equations={}
    if assembly_report:
        for item in assembly.get('exceptions',[]):
            if item.get('kind')!='referenced_equation_number_retained':continue
            mapped=next((m for m in assembly['mapping'] if m['source_id']==item.get('source_id') and m['block_id']==item.get('block_id')),None)
            bound=next((s for s in assembly['sources'] if s['source_id']==item.get('source_id')),None)
            if not mapped or not bound:ck(False,'retained_equation_source_binding');continue
            original=Package(bound['path']);op=original.body[mapped['source_index']];number=item.get('number','')
            # Independent literal citation check; a suffix in 格式126 or
            # 公式3×10^-8 is not evidence for retaining equation number 3.
            number_token=re.escape(number);reference=False
            for n in original.body:
                if n is op:continue
                value=ordinary(n)
                matches=list(re.finditer(r'(?:公式|(?<![格方模样形制程公])式)\s*[（(]\s*'+number_token+r'\s*[）)]',value))
                matches+=list(re.finditer(r'(?:参见|见|由|根据|代入|利用|按|如)\s*(?:公式|式)\s*'+number_token+r'(?![\d.eE×xX*·])',value))
                matches+=list(re.finditer(r'(?:^|[\s，。；;：:])式\s*'+number_token+r'(?![\d.eE×xX*·])',value))
                reference=reference or any(not re.match(r'\s*(?:[×xX*·]\s*10|[eE][+\-]?\d)',value[m.end():]) for m in matches)
            carrier=bool(op.findall('.//m:oMath',N)) or any(n.get('equationxml') is not None or n.get('ProgID','').lower().startswith(('equation.','mathtype.')) for n in op.iter())
            valid=(bool(reference) and carrier and re.fullmatch(r'\s*[（(]\s*'+number_token+r'\s*[）)]\s*',ordinary(op)) is not None)
            ck(valid,'retained_equation_has_original_reference')
            if valid:retained_equations[mapped['output_index']]=ordinary(op)
    retained_layouts=set()
    if assembly_report:
        for item in assembly.get('exceptions',[]):
            if item.get('kind')!='equation_layout_retained':continue
            mapped=next((m for m in assembly['mapping'] if m['source_id']==item.get('source_id') and m['block_id']==item.get('block_id')),None)
            bound=next((s for s in assembly['sources'] if s['source_id']==item.get('source_id')),None)
            if not mapped or not bound:ck(False,'retained_layout_source_binding');continue
            original=Package(bound['path']).body[mapped['source_index']]
            valid=bool(original.findall('.//m:oMath',N)) and ((item.get('reason')=='nested_native_math_carrier' and not original.findall('m:oMath',N)) or
                 (item.get('reason')=='retained_visible_content' and mapped['output_index'] in retained_equations))
            ck(valid,'retained_equation_layout_has_original_container_evidence')
            if valid:retained_layouts.add(mapped['output_index'])
    for p in [p for node in report_nodes for p in node.iter(q('w:p'))]:
        flag=p.find('w:pPr/w:pageBreakBefore',N);ck(flag is None or flag.get(q('w:val'),'1') in ('0','false','off'),'no_forced_paragraph_break')
        tail=bool(p.findall('.//m:oMath',N) and re.fullmatch(r'\s*[（(]\s*\d+\s*[）)]\s*',ordinary(p)))
        owner=p
        while owner.getparent() is not out.body and owner.getparent() is not None:owner=owner.getparent()
        approved=owner.getparent() is out.body and retained_equations.get(out.body.index(owner))==ordinary(p)
        ck(not tail or approved,'no_equation_number_tail_without_bound_retention')
        if p.getparent().tag==q('w:body') and p.findall('.//m:oMath',N) and not ordinary(p).strip():
            ck(bool(p.findall('m:oMathPara',N)) or out.body.index(p) in retained_layouts,'standalone_math_uses_display_wrapper_or_bound_original_carrier')
    ck(not any(node.xpath('.//w:br[@w:type="page"]',namespaces=N) for node in report_nodes),'no_source_page_break_run')
    upstream=verify_assembly(assembly_report) if assembly_report else None
    if upstream:ck(upstream['status']=='PASS','frozen_source_assembly_checks')
    return {'status':'PASS' if not errors else 'FAIL','checks':checks,'errors':errors,'original_native_math':count_old,'new_native_math':len(r['conversions']),'total_native_math':count_new,'pending':r['pending'],'upstream':upstream,'scope':'Independent literal OMML decoding plus inverse mapping to the frozen-source-verified assembly; WPS layout not asserted.'}
