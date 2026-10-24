"""Organization tree and prepared template; never changes source originals."""
from pathlib import Path
from copy import deepcopy
from zipfile import ZipFile,ZIP_DEFLATED
import hashlib,json
from lxml import etree as E
from .docx_formatting import q,NS,elt,pack
from .job_support import sha,write


def tree_plan(entries):
    root={'children':[]}
    for e in sorted(entries,key=lambda x:(x['order'],x['id'])):
        node=root
        for title in e['group']:
            match=next((n for n in node['children'] if n.get('kind')=='group' and n['title']==title),None)
            if match is None:
                match=dict(kind='group',title=title,children=[]);node['children'].append(match)
            node=match
        node['children'].append(dict(kind='source',source_id=e['id'],title=e['title']))
    sources=[];pending=[]
    def walk(node,prefix):
        for i,c in enumerate(node['children'],1):
            address=prefix+[i]
            if len(address)>8:raise ValueError('归类深度超过八层；请减少分类层级')
            if c['kind']=='group':
                pending.append(dict(title=c['title'],prefix=address));walk(c,address)
            else:
                sources.append(dict(source_id=c['source_id'],prefix=address,groups_before=deepcopy(pending)));pending.clear()
    walk(root,[])
    return dict(schema='organization-v1',tree=root,sources=sources)


def prepare_template(source,dest):
    """H7--H9 use the approved H6 appearance, with independent outline levels.

    This creates a separately frozen template. Original/template front matter
    and all parts except styles and the isolated chapter definition stay exact.
    """
    from .resolve_styles import inspect_template
    profile=inspect_template(source)
    with ZipFile(source) as z:parts={n:z.read(n) for n in z.namelist()}
    styles=E.fromstring(parts['word/styles.xml']);byid={s.get(q('w:styleId')):s for s in styles}
    base=byid[profile['styles']['Heading 6']]
    for level in range(7,10):
        matches=[s for s in styles if s.find('w:name',NS) is not None and s.find('w:name',NS).get(q('w:val')).casefold()=='heading '+str(level)]
        if len(matches)>1:raise ValueError('模板深层标题样式重名')
        old=matches[0] if matches else None;copy=deepcopy(base)
        copy.set(q('w:styleId'),old.get(q('w:styleId')) if old is not None else 'FormsHeading'+str(level))
        copy.find('w:name',NS).set(q('w:val'),'heading '+str(level))
        for n in copy.findall('w:link',NS):copy.remove(n)
        pp=copy.find('w:pPr',NS)
        if pp is None:pp=elt('w:pPr');copy.append(pp)
        outline=pp.find('w:outlineLvl',NS)
        if outline is None:outline=elt('w:outlineLvl');pp.append(outline)
        outline.set(q('w:val'),str(level-1))
        num=pp.find('w:numPr',NS)
        if num is not None:
            lvl=num.find('w:ilvl',NS)
            if lvl is None:lvl=elt('w:ilvl');num.append(lvl)
            lvl.set(q('w:val'),str(level-1))
        if old is not None:styles.replace(old,copy)
        else:styles.append(copy)
    parts['word/styles.xml']=pack(styles)
    if 'word/numbering.xml' in parts:
        from .numbering_facts import NumberingFacts
        nf=NumberingFacts(parts);p=elt('w:p');pp=elt('w:pPr');pp.append(elt('w:pStyle',val=profile['styles']['Heading 6']));p.append(pp)
        nid=nf.inherited(p,'w:numPr/w:numId');nr=E.fromstring(parts['word/numbering.xml'])
        n=next((v for v in nr.findall('w:num',NS) if v.get(q('w:numId'))==str(nid)),None)
        if n is not None:
            aid=n.find('w:abstractNumId',NS).get(q('w:val'));a=next((v for v in nr.findall('w:abstractNum',NS) if v.get(q('w:abstractNumId'))==aid),None)
            last=a.find('w:lvl[@w:ilvl="5"]',NS)
            if last is None:raise ValueError('模板章节编号缺少六级定义')
            for level in range(6,9):
                old=a.find('w:lvl[@w:ilvl="'+str(level)+'"]',NS);copy=deepcopy(last);copy.set(q('w:ilvl'),str(level))
                copy.find('w:lvlText',NS).set(q('w:val'),'.'.join('%'+str(i) for i in range(1,level+2)))
                restart=copy.find('w:lvlRestart',NS)
                if restart is not None:restart.set(q('w:val'),str(level))
                ps=copy.find('w:pStyle',NS)
                if ps is not None:copy.remove(ps)
                if old is not None:a.replace(old,copy)
                else:a.append(copy)
            parts['word/numbering.xml']=pack(nr)
    with ZipFile(dest,'x',ZIP_DEFLATED) as z:
        for name,data in parts.items():z.writestr(name,data)
    return dict(original=str(Path(source).resolve()),original_sha256=sha(source),prepared_sha256=sha(dest),
                method='H7-H9 inherit approved H6 appearance; source template preserved')


def verify_organization(report,out):
    """Independent recomputation from the tree, not emitter numbering claims."""
    org=report.get('organization');errors=[]
    if not org:
        if any(s.get('source_prefix') not in (None,[],[s['source_number']]) for s in report['sources']):errors.append('organization:undeclared_nested_prefix')
        return errors
    if org.get('schema')!='organization-v1':return ['organization:unknown_schema']
    sources=[];groups=[];prefixes={}
    def walk(node,prefix):
        children=node.get('children',[])
        if not isinstance(children,list):raise ValueError('children')
        for i,c in enumerate(children,1):
            address=prefix+[i]
            if len(address)>8:raise ValueError('depth')
            if c.get('kind')=='group':
                if not c.get('title') or not c.get('children'):raise ValueError('empty group')
                groups.append((c['title'],address));walk(c,address)
            elif c.get('kind')=='source':
                sid=c['source_id']
                if sid in prefixes:raise ValueError('duplicate source')
                prefixes[sid]=address;sources.append(sid)
            else:raise ValueError('node kind')
    try:walk(org['tree'],[])
    except Exception as exc:return ['organization:invalid_tree:'+str(exc)]
    if sources!=report.get('selected_source_ids',[]):errors.append('organization:source_order')
    if [s['source_id'] for s in org.get('sources',[])]!=sources:errors.append('organization:execution_source_order')
    for source in report['sources']:
        if source.get('source_prefix')!=prefixes.get(source['source_id']):errors.append('organization:source_prefix:'+source['source_id'])
    actual=[(x['label'],x['prefix']) for x in report['generated_blocks'] if x.get('kind')=='organization_group']
    if actual!=groups:errors.append('organization:group_coverage')
    return errors
