"""Independent source-side audit of fields consumed by numbered-unit recognition.

Does not import Scanner or NumberingFacts. Unsupported label semantics remain
unresolved; this is deliberately not a general Word numbering renderer.
"""
from zipfile import ZipFile
from lxml import etree as E

N={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
W='{'+N['w']+'}'


def audit(source,scan):
    from .verify_report_build import Package
    from .verify_native_numbering import labels
    package=Package(source);independent=labels(package)
    bypath={package.root.getroottree().getpath(p):f for p,f in independent.items()}
    with ZipFile(source) as z:parts={n:z.read(n) for n in z.namelist()}
    root=E.fromstring(parts['word/document.xml']);styles_root=E.fromstring(parts['word/styles.xml'])
    styles={s.attrib[W+'styleId']:s for s in styles_root.findall('w:style',N)}
    default=next((k for k,s in styles.items() if s.get(W+'default')=='1' and s.get(W+'type')=='paragraph'),'Normal')
    numbering=E.fromstring(parts.get('word/numbering.xml',('<w:numbering xmlns:w="'+N['w']+'"/>').encode()))
    errors=[];checks=0
    def check(ok,label):
        nonlocal checks
        checks+=1
        if not ok:errors.append(label)
    def val(n,path):
        found=n.find(path,N) if n is not None else None
        return found.get(W+'val') if found is not None else None
    def inherited(p,sid,path):
        nodes=[p];seen=set()
        while sid and sid not in seen:
            seen.add(sid);s=styles.get(sid)
            if s is None:break
            nodes.append(s);sid=val(s,'w:basedOn')
        for n in nodes:
            result=val(n,'w:pPr/'+path)
            if result is not None:return result
        return None
    def layout(p,sid):
        chain=[p];visited=set()
        while sid in styles and sid not in visited:
            visited.add(sid);style=styles[sid];chain.append(style);sid=val(style,'w:basedOn')
        def attribute(path,name):
            for node in chain:
                prop=node.find('w:pPr/'+path,N) if node is not None else None
                if prop is not None and prop.get(W+name) is not None:return prop.get(W+name)
            return None
        result={'alignment':attribute('w:jc','val') or 'left'}
        for key in ('firstLine','firstLineChars','hanging','hangingChars','left','leftChars','right','rightChars'):
            result[key]=attribute('w:ind',key) or '0'
        return result
    raw_auto={}
    for p in root.findall('.//w:p',N):
        sid=val(p,'w:pPr/w:pStyle') or default;nid=inherited(p,sid,'w:numPr/w:numId')
        try:level=int(inherited(p,sid,'w:numPr/w:ilvl') or 0)
        except (ValueError,TypeError):level=None
        expected={'num_id':nid,'ilvl':level,'status':'none' if nid in (None,'0') else 'unresolved','label':None,'values':None,'definition':None}
        if nid not in (None,'0') and level is not None and 0<=level<=8:
            num=next((x for x in numbering.findall('w:num',N) if x.get(W+'numId')==nid),None)
            aid=val(num,'w:abstractNumId');abstract=next((x for x in numbering.findall('w:abstractNum',N) if x.get(W+'abstractNumId')==aid),None)
            if abstract is not None:
                expected['abstract_id']=aid;definitions={}
                for entry in abstract.findall('w:lvl',N):
                    try:
                        k=int(entry.get(W+'ilvl'));ov=next((x for x in num.findall('w:lvlOverride',N) if x.get(W+'ilvl')==str(k)),None)
                        replacement=ov.find('w:lvl',N) if ov is not None else None
                        n=replacement if replacement is not None else entry
                        definitions[k]={'format':val(n,'w:numFmt'),'text':val(n,'w:lvlText'),'start':int(val(ov,'w:startOverride') or val(n,'w:start') or 1),'suffix':val(n,'w:suff') or 'tab','restart':val(n,'w:lvlRestart')}
                    except (ValueError,TypeError):definitions={};break
                d=definitions.get(level);expected['definition']=d
        independent_facts=bypath[root.getroottree().getpath(p)]
        for key in ('status','label','values'):expected[key]=independent_facts.get(key)
        raw_auto[root.getroottree().getpath(p)]=expected
    for b in scan['blocks']:
        if b['kind']!='paragraph':continue
        bid=b['block_id'];p=root.xpath(b['source_locator']['path'],namespaces=N)[0]
        sid=val(p,'w:pPr/w:pStyle') or default;outline=inherited(p,sid,'w:outlineLvl')
        check(b.get('style_id')==sid,bid+':style_id')
        check(b.get('style_name')==val(styles.get(sid),'w:name'),bid+':style_name')
        check(b.get('outline_level_raw')==(int(outline) if outline is not None else None),bid+':effective_outline')
        check(b.get('numbering')=={'num_id':val(p,'w:pPr/w:numPr/w:numId'),'ilvl':val(p,'w:pPr/w:numPr/w:ilvl')},bid+':direct_numbering')
        check(b.get('numbering_effective')==raw_auto[root.getroottree().getpath(p)],bid+':effective_numbering_and_label')
        if 'paragraph_format_effective' in b:check(b['paragraph_format_effective']==layout(p,sid),bid+':effective_paragraph_layout')
        if b.get('format_group'):check(b['format_group'][0]==sid,bid+':format_style_id')
    if 'body_width_emu' in scan:
        section=root.find('.//w:body//w:sectPr',N)
        page=section.find('w:pgSz',N) if section is not None else None
        margins=section.find('w:pgMar',N) if section is not None else None
        width=(int(page.get(W+'w'))-int(margins.get(W+'left'))-int(margins.get(W+'right')))*635 if page is not None and margins is not None else None
        check(scan['body_width_emu']==width,'first_section_body_width')
    if 'normal_paragraph_layout' in scan:
        normal=layout(None,default)
        check(scan['normal_paragraph_layout']==[normal['alignment']]+[normal[k] for k in ('firstLine','firstLineChars','hanging','hangingChars','left','leftChars','right','rightChars')],'default_paragraph_layout')
    return {'status':'FAIL' if errors else 'PASS','checks':checks,'errors':errors,'scope':'source style id/name, effective outline, direct/effective numbering, supported sequence formats and optional paragraph/first-section layout facts; no page verification'}
