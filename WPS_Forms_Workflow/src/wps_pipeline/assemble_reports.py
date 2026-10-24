"""Template package assembly for user-selected sources, with per-block provenance.

Reuses the earlier style/layout helpers. Copies only referenced images and
hyperlinks, remaps bookmark identities, preserves native math and containers.
No model, COM, field refresh, source writes, or visual validation.
"""
import argparse,json,re,posixpath,hashlib
from pathlib import Path
from collections import Counter
from copy import deepcopy
from zipfile import ZipFile,ZIP_DEFLATED
from lxml import etree as E
from .docx_formatting import NS,q,xml,pack,elt,ordinary_text,style_paragraph,style_table,style_figure_table,picture_row_table,sha,normalize_embedded_story_styles
from .resolve_styles import inspect_template
from .numbering_facts import NumberingFacts
from .numbered_candidates import validate_edit,validate_list_edit,CONTRACT,select_section,caption_prefix,caption_number_key,CAPTION_REFERENCE
from .native_output import number_properties,manual_list_definition,field_views,update_field,freeze_field,promote_text_field,sequence_field,number_bookmarks,wrap_number_bookmark,SEQ_FIELD
from .object_transport import pictures,is_legacy_equation,image_attribute,part_name,missing_png,equation_references

REL='http://schemas.openxmlformats.org/package/2006/relationships'
CT='http://schemas.openxmlformats.org/package/2006/content-types'
ROLE_STYLE={'document_title':'Title','front_heading':'Heading 2','body':'Normal','spacer':'Normal','figure':'图形段','figure_caption':'图题','table_caption':'表题','equation':'公式','equation_description':'Normal','list_item':'List Paragraph','reference':'参考条目','table_text':'List Paragraph','boundary':'Normal'}

def text_edit(p,start,end,new):
    nodes=p.xpath('.//w:t[not(ancestor::m:oMath)]',namespaces=NS);cursor=0;inserted=False
    if start==end and start>0:
        for n in nodes:
            value=n.text or ''
            if cursor<=start<cursor+len(value):
                at=start-cursor;n.text=value[:at]+new+value[at:];n.set('{http://www.w3.org/XML/1998/namespace}space','preserve');return
            cursor+=len(value)
        if start!=cursor:raise ValueError('Insertion beyond content')
        run=elt('w:r');t=elt('w:t');t.text=new;run.append(t);p.append(run);return

    for n in nodes:
        value=n.text or '';a,b=cursor,cursor+len(value);cursor=b
        if end<=a or start>=b:continue
        lo=max(0,start-a);hi=min(len(value),end-a)
        n.text=value[:lo]+(new if not inserted else '')+value[hi:];inserted=True;n.set('{http://www.w3.org/XML/1998/namespace}space','preserve')
    if start==end:
        run=elt('w:r');t=elt('w:t');t.text=new;t.set('{http://www.w3.org/XML/1998/namespace}space','preserve');run.append(t);p.insert(1 if p.find(q('w:pPr')) is not None else 0,run)
    elif not inserted:raise ValueError('Text span outside content')


# Narrow, explicit field support. Managed SEQ and REF remain editable fields.
REF_FIELD=re.compile(r'\s*REF\s+([\w]+)\s*((?:\\h\s*|\\\*\s+MERGEFORMAT\s*)*)',re.I)

def caption_seq(p):
    return sequence_field(p)


def relative_path(root,node):
    """A structural source locator, independent of style/property insertions."""
    parts=[]
    while node is not root:
        parent=node.getparent()
        if parent is None:raise ValueError('Object outside source block')
        tag=E.QName(node);prefix=next(k for k,v in NS.items() if v==tag.namespace)
        siblings=[c for c in parent if c.tag==node.tag]
        parts.append(f'{prefix}:{tag.localname}[{siblings.index(node)+1}]');node=parent
    return './'+'/'.join(reversed(parts)) if parts else '.'


def physical_objects(rows,body):
    """One number per actual figure container/data table, never per title text.

    An image layout table is one group, including its internal layout cells.
    Data-table nesting remains data; image paragraphs inside data cells get
    their own figure number. Repeated media bytes are distinct occurrences.
    """
    answer=[]
    for row in rows:
        node=body[row['source_locator']['body_child_index']]
        if row.get('figure_layout') and pictures(node):
            nodes=[('图',node)]
        elif node.tag!=q('w:p'):
            nodes=[('表' if n.tag==q('w:tbl') else '图',n) for n in node.iter()
                   if n.tag==q('w:tbl') or (n.tag==q('w:p') and pictures(n))]
        elif node.tag==q('w:p') and pictures(node):nodes=[('图',node)]
        else:nodes=[]
        for label,target in nodes:
            path=relative_path(node,target)
            answer.append(dict(block_id=row['block_id'],source_index=row['source_locator']['body_child_index'],
                               source_path=path,label=label,
                               image_instances=len(pictures(target)) if label=='图' else 0))
    return answer

def protected_reference_span(p,start,end,anchors,converted_nodes=()):
    cursor=0
    if p.findall('.//w:fldChar',NS):return True
    for t in p.xpath('.//w:t[not(ancestor::m:oMath)]',namespaces=NS):
        left=cursor;cursor+=len(t.text or '')
        if start<cursor and end>left and (t in converted_nodes or any(a.tag==q('w:fldSimple') or (a.tag==q('w:hyperlink') and a.get(q('w:anchor')) in anchors) for a in t.iterancestors())):return True
    return False


class Assembler:
    def __init__(self,template,master=False,allow_pending=False,master_title='工程技术报告汇编'):
        self.template=Path(template);self.master=master;self.allow_pending=allow_pending
        with ZipFile(template) as z:self.parts={n:z.read(n) for n in z.namelist()}
        self.original_parts=dict(self.parts);self.root=xml(self.parts['word/document.xml']);self.body=self.root.find(q('w:body'));self.section=self.body.find(q('w:sectPr'))
        profile=inspect_template(template)
        self.styles=profile['styles']
        self.template_version=profile['template_version']
        from .resolve_styles import open_template_body
        self.template_profile=profile
        self.title_pattern=open_template_body(self.body,profile)
        pg=self.section.find(q('w:pgSz'));ma=self.section.find(q('w:pgMar'));self.width=int(pg.get(q('w:w')))-int(ma.get(q('w:left')))-int(ma.get(q('w:right')))
        self.rels=xml(self.parts['word/_rels/document.xml.rels']);self.types=xml(self.parts['[Content_Types].xml']);self.ids={r.get('Id') for r in self.rels}
        self.records=[];self.exclusions=[];self.sources=[];self.exceptions=[];self.images=[];self.math_colors=[];self.caption_count=Counter();self.mark_id=0;self.docpr=0;self.extra=[];self.object_numbers=[]
        self.mark_id=max([int(n.get(q('w:id'))) for n in self.body.iter(q('w:bookmarkStart'))]+[0])
        self.numbering_root=xml(self.parts['word/numbering.xml']) if 'word/numbering.xml' in self.parts else elt('w:numbering')
        self.numbering_changed=False;self.numbering_retention=[];self.gaps=[]
        heading_facts=NumberingFacts(self.parts);self.template_style_numbers={}
        for name,ident in self.styles.items():
            probe=elt('w:p');pp=elt('w:pPr');pp.append(elt('w:pStyle',val=ident));probe.append(pp)
            self.template_style_numbers[name]=heading_facts.inherited(probe,'w:numPr/w:numId')
        bindings=[self.template_style_numbers['Heading '+str(level)] for level in range(1,7)]
        self.chapter_num_id=bindings[0] if len(set(bindings))==1 and bindings[0] not in (None,'0') else None
        if self.chapter_num_id:
            for item in list(self.body)[:-1]:
                if any(heading_facts.inherited(p,'w:numPr/w:numId')==self.chapter_num_id for p in item.iter(q('w:p'))):raise ValueError('Report outline intersects protected template content; prepare an isolated template')
        if master:
            if not isinstance(master_title,str) or not master_title.strip() or len(master_title)>200:
                raise ValueError('Master title must contain 1..200 characters')
            p=deepcopy(self.title_pattern);old=ordinary_text(p);text_edit(p,0,len(old),master_title);pp=style_paragraph(p,self.styles['Title'],False);self.cancel_title_rule(pp);index=len(self.body)-1;self.body.insert(index,p);self.extra.append({'output_index':index,'text':master_title})
    def cancel_title_rule(self,pp):
        pb=elt('w:pBdr');pb.append(elt('w:bottom',val='nil'));pp.append(pb)
    def add_rel(self,rel,data=None,extension=None):
        i=1
        while f'rId{i}' in self.ids:i+=1
        r=deepcopy(rel);r.set('Id',f'rId{i}');self.ids.add(f'rId{i}');self.rels.append(r)
        return r
    def import_numbering(self,parts,nid,mapping,first_values=None,preserve_definition=False):
        if nid in mapping:return mapping[nid]
        root=xml(parts['word/numbering.xml']);old=next((n for n in root.findall('w:num',NS) if n.get(q('w:numId'))==nid),None)
        if old is None:raise ValueError('Missing source numbering definition')
        aid=old.find('w:abstractNumId',NS).get(q('w:val'));abstract=next((n for n in root.findall('w:abstractNum',NS) if n.get(q('w:abstractNumId'))==aid),None)
        if abstract is None or abstract.find('w:numStyleLink',NS) is not None:raise ValueError('Unsupported linked numbering')
        newaid=str(max([int(n.get(q('w:abstractNumId'))) for n in self.numbering_root.findall('w:abstractNum',NS)]+[-1])+1)
        newnid=str(max([int(n.get(q('w:numId'))) for n in self.numbering_root.findall('w:num',NS)]+[0])+1)
        a=deepcopy(abstract);a.set(q('w:abstractNumId'),newaid)
        # Keep label/format/start definitions, apply the target flush-left layout.
        for lvl in ([] if preserve_definition else a.findall('w:lvl',NS)):
            pp=lvl.find('w:pPr',NS)
            if pp is not None:lvl.remove(pp)
            pp=elt('w:pPr');pp.append(elt('w:ind',left=0,right=0,firstLine=0,hanging=0,leftChars=0,firstLineChars=0,hangingChars=0));lvl.append(pp)
            for child in list(lvl):
                if child.tag==q('w:pStyle'):lvl.remove(child)
                elif child.tag==q('w:rPr'):
                    fmt=lvl.find('w:numFmt',NS)
                    if fmt is not None and fmt.get(q('w:val'))=='bullet':
                        # Literal bullet glyphs may require Symbol/Wingdings.
                        for prop in list(child):
                            if prop.tag!=q('w:rFonts'):child.remove(prop)
                        child.append(elt('w:sz',val=24));child.append(elt('w:szCs',val=24))
                    else:lvl.remove(child)
        self.numbering_root.insert(len(self.numbering_root.findall('w:abstractNum',NS)),a)
        n=deepcopy(old);n.set(q('w:numId'),newnid);n.find('w:abstractNumId',NS).set(q('w:val'),newaid)
        for level,start in enumerate(first_values or []):
            override=n.find('w:lvlOverride[@w:ilvl="'+str(level)+'"]',NS)
            if override is None:override=elt('w:lvlOverride',ilvl=level);n.append(override)
            previous=override.find('w:startOverride',NS)
            if previous is not None:override.remove(previous)
            override.insert(0,elt('w:startOverride',val=start))
        self.numbering_root.append(n)
        mapping[nid]=newnid;self.numbering_changed=True
        if not any(r.get('Type','').endswith('/numbering') for r in self.rels):
            relation=E.Element('{'+REL+'}Relationship',Type=NS['r']+'/numbering',Target='numbering.xml');self.add_rel(relation)
        if not any(t.get('PartName')=='/word/numbering.xml' for t in self.types):
            E.SubElement(self.types,'{'+CT+'}Override',PartName='/word/numbering.xml',ContentType='application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml')
        return newnid

    def plan_manual_lists(self,rows,pred,source_hash,body):
        byid={r['block_id']:r for r in rows};groups={g['id']:g for g in pred.get('number_groups',[])};plans={};roots={}
        for row in rows:
            if row['role']!='list_item' or not row.get('list_edit',{}).get('allowed'):continue
            auth=validate_list_edit(row,source_hash,ordinary_text(body[row['source_locator']['body_child_index']]))
            group=groups.get(row.get('number_group_id'));level=0;seen=set();root=group
            while root and root.get('owner_candidate_id') in byid:
                owner=byid[root['owner_candidate_id']]
                if owner['role']!='list_item':break
                if not (root.get('closed_parent_interval') or root.get('explicit_path_parent')):
                    root=None;break
                parent=groups.get(owner.get('number_group_id'))
                if parent is None or parent['id'] in seen:break
                seen.add(parent['id']);root=parent;level+=1
            if not group or not root or level>8:
                self.exceptions.append(dict(kind='manual_list_hierarchy_retained',block_id=row['block_id'],message='列表父层级未确定，保留原号'));continue
            if auth['form']=='decimal_path' and len(auth['numbers'])!=level+1:
                self.exceptions.append(dict(kind='manual_list_hierarchy_retained',block_id=row['block_id'],message='列表多级路径未能唯一绑定父项，保留原号'));continue
            config=roots.setdefault(root['id'],dict(forms={},starts={}))
            if level in config['forms'] and config['forms'][level]!=auth['form']:
                self.exceptions.append(dict(kind='manual_list_marker_conflict_retained',block_id=row['block_id'],message='同一列表层级存在不同编号形式，保留局部原号'));continue
            config['forms'][level]=auth['form'];config['starts'].setdefault(level,auth['numbers'][-1]);plans[row['block_id']]=(root['id'],level,auth)
        for key,value in roots.items():
            value['num_id']=manual_list_definition(self.numbering_root,value['forms'],value['starts']);self.numbering_changed=True
        return {bid:dict(num_id=roots[key]['num_id'],level=level,authorization=auth) for bid,(key,level,auth) in plans.items()}

    def add_group(self,label,prefix):
        """Explicit organization node, bound to the checked forms plan."""
        level=len(prefix)
        if not self.master or not 1<=level<=8 or 'Heading '+str(level) not in self.styles:
            raise ValueError('Organization heading exceeds available template levels')
        p=elt('w:p');run=elt('w:r');t=elt('w:t');t.text=label;run.append(t);p.append(run)
        pp=style_paragraph(p,self.styles['Heading '+str(level)],False)
        if self.chapter_num_id:number_properties(pp,self.chapter_num_id,level-1)
        else:text_edit(p,0,0,'.'.join(map(str,prefix))+' ')
        index=len(self.body)-1;self.body.insert(index,p)
        self.extra.append(dict(kind='organization_group',label=label,prefix=prefix,
                               output_index=index,text=ordinary_text(p)))

    def add_source(self,source,candidate,sid,source_number=1,section_id=None,source_title=None,source_prefix=None,display_title=None):
        outline_prefix=(source_prefix or [source_number]) if self.master else []
        source_depth=len(outline_prefix)
        group_style='Heading '+str(source_depth) if self.master else 'Title'
        source=Path(source);pred=json.loads(Path(candidate).read_text(encoding='utf8'))
        if pred['source_sha256']!=sha(source):raise ValueError('Source/candidate hash mismatch')
        if pred.get('numbering_contract')!=CONTRACT:raise ValueError('Old decisions need a fresh scan and contextual recognition before number edits')
        if any(r['status']=='review' for r in pred['blocks']) and not self.allow_pending:raise ValueError('Explicit local resolution required; whole-document draft may preserve pending blocks')
        with ZipFile(source) as z:sp={n:z.read(n) for n in z.namelist()}
        sr=xml(sp['word/document.xml']);sb=sr.find(q('w:body'));rels={x.get('Id'):x for x in xml(sp['word/_rels/document.xml.rels'])};s_types=xml(sp['[Content_Types].xml'])
        source_styles={s.get(q('w:styleId')):s for s in xml(sp['word/styles.xml']).findall('w:style',NS)}
        rows=deepcopy(pred['blocks']);num=[0]*5;target={};caption_maps={'图':{},'表':{}};relmap={};bmids={};bmnames={};nummap={};retained_numbers={}
        missing_repairs={};legacy_ids={}
        for e in sb.iter():
            if e.tag in (q('v:shape'),q('v:shapetype')) and e.get('id'):legacy_ids[e.get('id')]=sid+'_'+e.get('id')
            if e.tag==q('o:OLEObject') and e.get('ObjectID'):legacy_ids[e.get('ObjectID')]=sid+'_'+e.get('ObjectID')
        front=pred.get('source_frontmatter',{});excluded_indices=set()
        for item in front.get('exclusions',[]):
            index=item['source_index'];node=sb[index]
            if item['source_sha256']!=pred['source_sha256'] or item['xml_sha256']!=hashlib.sha256(E.tostring(node,encoding='utf8',with_tail=False)).hexdigest():raise ValueError('Source front matter binding changed')
            if item['kind'] not in ('source_cover','source_toc'):raise ValueError('Unsupported source exclusion')
            excluded_indices.add(index);self.exclusions.append(dict(item,source_id=sid,source_text=ordinary_text(node)))
        declared={item['block_id'] for item in front.get('exclusions',[])}
        if any(r['role'] in ('source_cover','source_toc') and r['block_id'] not in declared for r in rows):raise ValueError('Undeclared source exclusion')
        rows=[r for r in rows if r['block_id'] not in declared]
        self.exceptions.extend(dict(w,source_id=sid) for w in front.get('warnings',[]))
        title_info=front.get('title')
        if not self.master and title_info and not any(r['role']=='document_title' for r in rows):
            p=elt('w:p');run=elt('w:r');t=elt('w:t');t.text=title_info['text'];run.append(t);p.append(run)
            pp=style_paragraph(p,self.styles['Title'],False);self.cancel_title_rule(pp);number_properties(pp,0)
            index=len(self.body)-1;self.body.insert(index,p)
            self.extra.append(dict(kind='source_document_title',source_id=sid,text=title_info['text'],output_index=index))
        nf=NumberingFacts(sp);actual_numbering={sr.getroottree().getpath(p):nf.inspect(p) for p in sb.iter(q('w:p'))}
        if section_id:
            selected=set(select_section(pred,section_id));rows=[r for r in rows if r['block_id'] in selected]
        retained_paragraphs=[p for r in rows for p in sb[r['source_locator']['body_child_index']].iter(q('w:p'))]
        unknown_ids={actual_numbering[sr.getroottree().getpath(p)]['num_id'] for p in retained_paragraphs if actual_numbering[sr.getroottree().getpath(p)].get('status')=='unresolved' and actual_numbering[sr.getroottree().getpath(p)].get('num_id')}
        safe_unknown=set();numbering_notices={}
        if unknown_ids:
            numbering=xml(sp.get('word/numbering.xml',b'<w:numbering xmlns:w="'+NS['w'].encode()+b'"/>'))
            row_by_index={r['source_locator']['body_child_index']:r for r in rows}
            for nid in unknown_ids:
                definition=next((n for n in numbering.findall('w:num',NS) if n.get(q('w:numId'))==nid),None)
                aid=definition.find('w:abstractNumId',NS) if definition is not None else None
                abstract=next((n for n in numbering.findall('w:abstractNum',NS) if n.get(q('w:abstractNumId'))==aid.get(q('w:val'))),None) if aid is not None else None
                reason=('missing_definition' if abstract is None else 'dependent_definition' if any(n.tag in (q('w:numStyleLink'),q('w:styleLink'),q('w:pStyle'),q('w:lvlPicBulletId')) or any(k.startswith('{'+NS['r']+'}') for k in n.attrib) for base in (abstract,definition) for n in base.iter()) else None)
                for part,data in sp.items():
                    if re.fullmatch(r'word/(?:header[^/]*|footer[^/]*|footnotes|endnotes|comments)\.xml',part):
                        if any(nf.inherited(p,'w:numPr/w:numId')==nid for p in xml(data).iter(q('w:p'))):reason=reason or 'external_story_instance'
                for p in sb.iter(q('w:p')):
                    auto=actual_numbering[sr.getroottree().getpath(p)]
                    if auto.get('num_id')!=nid:continue
                    owner=p
                    while owner.getparent() is not sb:owner=owner.getparent()
                    if sb.index(owner) not in row_by_index:reason=reason or 'incomplete_instance';continue
                    if auto.get('ilvl') is None or not 0<=auto['ilvl']<=8:reason='invalid_level'
                    item=row_by_index[sb.index(owner)]
                    if item['role']=='heading':
                        item.update(retained_role='heading',retained_level=item.get('level'),status='review')
                if reason:numbering_notices[nid]=reason
                else:safe_unknown.add(nid)
        # Physical objects receive numbers before optional titles are resolved.
        # A pending/orphan title never freezes another object's numbering.
        row_by_id={r['block_id']:r for r in rows};held_caption_roles=set();pairs={}
        source_caption_keys=Counter()
        for r in rows:
            if r['role'] in ('figure_caption','table_caption') and not r.get('subcaptions'):
                prefix=caption_prefix(ordinary_text(sb[r['source_locator']['body_child_index']]))
                if prefix:source_caption_keys[prefix['label'],prefix['number_key']]+=1
        objects=physical_objects(rows,sb)
        top_objects={(o['block_id'],o['label']):o for o in objects if o['source_path']=='.'}
        for obj in objects:
            self.caption_count[obj['label']]+=1;obj['number']=str(self.caption_count[obj['label']])
            identity=sid+':'+obj['block_id']+':'+obj['source_path']+':'+obj['label']
            obj.update(source_id=sid,source_sha256=pred['source_sha256'],
                       object_id='OBJ_'+hashlib.sha256(identity.encode()).hexdigest()[:24],
                       bookmark_name='WPSN_'+hashlib.sha256(identity.encode()).hexdigest()[:24],caption_block_id=None)
        for r in rows:
            role=r['role']
            if role not in ('figure_caption','table_caption'):continue
            related=r.get('relates_to');target_row=row_by_id.get(related)
            if related:
                if target_row is None or related==r['block_id']:raise ValueError('Caption target outside selected source or self-related')
                physical=sb[target_row['source_locator']['body_child_index']]
                valid=(role=='table_caption' and physical.tag==q('w:tbl') and not target_row.get('figure_layout')) or (role=='figure_caption' and bool(pictures(physical)) and (physical.tag==q('w:p') or (physical.tag==q('w:tbl') and target_row.get('figure_layout',{}).get('kind')=='table')))
                if not valid:raise ValueError('Caption target physical type mismatch')
            p=sb[r['source_locator']['body_child_index']]
            supported=not (p.findall('.//w:fldSimple',NS) or p.findall('.//w:fldChar',NS)) or caption_seq(p) is not None
            label='图' if role=='figure_caption' else '表'
            if related and r['status'] not in ('review','retained_unpaired') and supported and not r.get('subcaptions'):
                pairs.setdefault((related,label),[]).append(r)
        for key,linked in pairs.items():
            if len(linked)==1 and key in top_objects:top_objects[key]['caption_block_id']=linked[0]['block_id']
        chosen={o['caption_block_id']:o for o in objects if o['caption_block_id']}
        for r in rows:
            if r['role'] not in ('figure_caption','table_caption') or r['block_id'] in chosen or r.get('subcaptions'):continue
            self.exceptions.append(dict(source_id=sid,block_id=r['block_id'],kind='caption_association_unresolved',message='题名关联未确定，原文按正文保留；实际对象已独立编号，可在后续评审中绑定题名',candidates=r.get('association_candidates',[])))
            r.update(retained_caption_role=r['role'],role='body',level=None,caption_retained=True)
        for r in rows:
            if r['role'] in ('figure_caption','table_caption') and r['status']=='review':r['status']='retained_unpaired'
            if (r.get('number_candidate') or {}).get('conflict')=='automatic_and_manual':
                self.exceptions.append(dict(source_id=sid,block_id=r['block_id'],kind='automatic_and_manual_numbering_preserved',message='自动编号与手打编号并存，保留两者；角色确认不等于授权删除其中一套'))
            if r['status']=='review':
                self.exceptions.append({'source_id':sid,'block_id':r['block_id'],'kind':'unresolved_structure_preserved','evidence':r.get('evidence')})
                r['role']='heading_retained' if r.get('retained_role')=='heading' else 'body'
                r['level']=r.get('retained_level') if r['role']=='heading_retained' else None
            if r['role']=='heading':
                idx=r['source_locator']['body_child_index'];validate_edit(r,pred['source_sha256'],ordinary_text(sb[idx]),actual_numbering.get(sr.getroottree().getpath(sb[idx])))
        manual_lists=self.plan_manual_lists(rows,pred,pred['source_sha256'],sb)
        group_title=None
        if self.master and not any(r['role']=='document_title' for r in rows):
            label=source_title or (title_info or {}).get('text') or source.stem
            if not isinstance(label,str) or not label.strip() or len(label)>200:
                raise ValueError('Source group title must contain 1..200 characters')
            text=label if self.chapter_num_id else '.'.join(map(str,outline_prefix))+' '+label
            p=elt('w:p');run=elt('w:r');run.append(elt('w:t'));run[-1].text=text;p.append(run)
            pp=style_paragraph(p,self.styles[group_style],False)
            if self.chapter_num_id:number_properties(pp,self.chapter_num_id,source_depth-1)
            for flag,value in [('pageBreakBefore',0),('keepNext',1),('keepLines',1)]:pp.append(elt('w:'+flag,val=value))
            index=len(self.body)-1;self.body.insert(index,p)
            group_title={'source_id':sid,'source_number':source_number,'label':label,
                         'origin':'explicit_source_label_or_filename','text':text,'native_number':'.'.join(map(str,outline_prefix)) if self.chapter_num_id else None,'source_prefix':outline_prefix}
            self.extra.append(dict(group_title,kind='source_group_title',output_index=index))
        retained_nodes=[sb[r['source_locator']['body_child_index']] for r in rows]
        mark_groups={};kept_events={}
        for n in sb.iter():
            if n.tag in (q('w:bookmarkStart'),q('w:bookmarkEnd')):mark_groups.setdefault(n.get(q('w:id')),[]).append(n)
        for root in retained_nodes:
            for n in root.iter():
                if n.tag in (q('w:bookmarkStart'),q('w:bookmarkEnd')):kept_events.setdefault(n.get(q('w:id')),[]).append(n)
        detached_bookmarks=[];invalid_bookmarks=set()
        kept_name_counts=Counter(n.get(q('w:name')) for nodes in kept_events.values() for n in nodes if n.tag==q('w:bookmarkStart'))
        for ident,nodes in kept_events.items():
            starts=[n for n in nodes if n.tag==q('w:bookmarkStart')];ends=[n for n in nodes if n.tag==q('w:bookmarkEnd')]
            original=mark_groups.get(ident,[])
            valid=(ident is not None and len(starts)==len(ends)==1 and len(original)==2 and
                   original[0].tag==q('w:bookmarkStart') and original[1].tag==q('w:bookmarkEnd') and
                   starts[0].get(q('w:name')) and kept_name_counts[starts[0].get(q('w:name'))]==1)
            if not valid:
                invalid_bookmarks.add(ident)
                detached_bookmarks.append(dict(bookmark_id=ident,source_starts=sum(n.tag==q('w:bookmarkStart') for n in original),
                    source_ends=sum(n.tag==q('w:bookmarkEnd') for n in original),kept_starts=len(starts),kept_ends=len(ends)))
        kept_marks=[m for node in retained_nodes for m in node.iter(q('w:bookmarkStart')) if m.get(q('w:id')) not in invalid_bookmarks]
        removed_bookmarks={m.get(q('w:name')) for m in sb.findall('.//w:bookmarkStart',NS) if m.get(q('w:name'))}-{m.get(q('w:name')) for m in kept_marks}
        for m in kept_marks:
            old=m.get(q('w:id'));self.mark_id+=1;bmids[old]=str(self.mark_id);bmnames[m.get(q('w:name'))]=f'{sid}_B{self.mark_id:05}'
        # Decide numbering before importing any block or touching references.
        for r in rows:
            idx=r['source_locator']['body_child_index'];text=ordinary_text(sb[idx]);role=r['role']
            if role=='heading':
                level=r['level'];num[level-1]+=1
                for k in range(level,5):num[k]=0
                if any(x==0 for x in num[:level]):raise ValueError('Heading has a missing parent')
                number='.'.join(map(str,num[:level]));target[r['block_id']]='.'.join(map(str,outline_prefix))+'.'+number if self.master else number
            elif r['block_id'] in chosen:
                label='图' if role=='figure_caption' else '表';prefix=caption_prefix(text)
                if prefix or not (sb[idx].findall('.//w:fldSimple',NS) or sb[idx].findall('.//w:fldChar',NS)):
                    p=sb[idx]
                    if (p.findall('.//w:fldSimple',NS) or p.findall('.//w:fldChar',NS)) and caption_seq(p) is None:
                        self.exceptions.append(dict(source_id=sid,block_id=r['block_id'],kind='caption_field_preserved',message='题名域不在已支持的SEQ范围，保留原显示及指令'))
                        continue
                    new=chosen[r['block_id']]['number'];old=prefix['number_key'] if prefix else None
                    if old is not None:
                        caption_maps[label][old]=new if source_caption_keys[label,old]==1 else None
                    target[r['block_id']]=label+' '+new+'　'
        field_targets={};heading_refs={};caption_info={};caption_by_old={'图':{},'表':{}}
        for r in rows:
            p=sb[r['source_locator']['body_child_index']]
            if r['role'] in ('figure_caption','table_caption') and r['block_id'] in target:
                f=caption_seq(p)
                label=target[r['block_id']][0];prefix=caption_prefix(ordinary_text(p));old=prefix['number_key'] if prefix else None
                names=number_bookmarks(p,prefix['number'],prefix['number_span']) if prefix else []
                if f:names=list(dict.fromkeys(names+number_bookmarks(p,f['cache'],f['text_span'])))
                names=[name for name in names if name in bmnames]
                name=bmnames[names[0]] if names else chosen[r['block_id']]['bookmark_name']
                chosen[r['block_id']]['bookmark_name']=name
                info=dict(block_id=r['block_id'],number=re.search(r'\d+',target[r['block_id']])[0],old=old,bookmark_name=name,create_bookmark=not names)
                caption_info[r['block_id']]=info
                if old is not None and caption_maps[label].get(old) is not None:caption_by_old[label][old]=info
                for name in names:field_targets[name]=info
            if r['role']=='heading':
                o=r.get('number_candidate') or {};values=o.get('numbers')
                if values:
                    old='.'.join(map(str,values));new=target[r['block_id']]
                    heading_refs[old]=new if old not in heading_refs else None
        equation_refs=[];equation_descriptions={};native_equation_numbers=set()
        for r in rows:
            node=sb[r['source_locator']['body_child_index']];value=ordinary_text(node)
            if r['role']=='equation' and node.findall('.//m:oMath',NS) and re.fullmatch(r'\s*[（(]\s*\d+\s*[）)]\s*',value):native_equation_numbers.add(re.search(r'\d+',value)[0])
        for r in rows:
            if r['role']=='equation_description':
                text=ordinary_text(sb[r['source_locator']['body_child_index']]);match=re.match(r'^\s*式\s*(\d+)\s+',text)
                if match:equation_descriptions[match.group(1)]=r['source_locator']['body_child_index']
        for r in rows:
            if r['role'] not in ('equation_description','equation'):
                equation_refs.extend(equation_references(ordinary_text(sb[r['source_locator']['body_child_index']])))
        table_caps={r['relates_to']:r for r in rows if r['role']=='table_caption' and r.get('relates_to')}
        ordered=[]
        for r in rows:
            if r['role']=='table_caption' and r.get('relates_to'):continue
            if r['role']=='table' and r['block_id'] in table_caps:ordered.append(table_caps[r['block_id']])
            ordered.append(r)
        for r in ordered:
            idx=r['source_locator']['body_child_index'];src=sb[idx];role=r['role'];node=deepcopy(src);before=ordinary_text(src);changes=[]
            for ident in invalid_bookmarks:
                marks=[n for n in node.iter() if n.tag in (q('w:bookmarkStart'),q('w:bookmarkEnd')) and n.get(q('w:id'))==ident]
                if not marks:continue
                root_marker=node in marks
                for mark in marks:
                    if mark is node:node.tag=q('w:p');node.attrib.clear()
                    else:
                        parent=mark.getparent();at=parent.index(mark)
                        for child in list(mark):parent.insert(at,child);at+=1
                        parent.remove(mark)
                changes.append(dict(operation='detach_unmatched_bookmark',bookmark_id=ident,count=len(marks),root_marker=root_marker))
                self.exceptions.append(dict(source_id=sid,block_id=r['block_id'],kind='bookmark_metadata_detached',resolved_by_policy=True,
                    message='不完整或跨前置范围的书签标记已局部解除，文字和对象保留',bookmark_id=ident))
            if role=='excluded_toc':
                raise ValueError('Legacy fixture exclusion is not authorized for general sources; rescan and preserve content')
            old_ppr=deepcopy(node.find(q('w:pPr'))) if node.tag==q('w:p') else None
            def edit(p,a,b,new,reason):
                pi=0 if node.tag==q('w:p') else node.findall('.//w:p',NS).index(p)
                old=ordinary_text(p)[a:b];text_edit(p,a,b,new);changes.append({'operation':reason,'paragraph_index':pi,'start':a,'end':b,'old':old,'new':new})
            # Ordinary reference-looking prose stays literal. A term ending in
            # 格式/公式 is never permission to rewrite its following number.
            # Preserve native SEQ/REF identity and only update bound instructions/caches.
            seq=caption_seq(node) if role in ('figure_caption','table_caption') and r['block_id'] in target else None
            converted_reference_nodes=set()
            for fi,f in reversed(list(enumerate(field_views(node)))):
                instr=f['instruction'];old=f['cache'];ref=REF_FIELD.fullmatch(instr)
                removed_ref=re.match(r'\s*(?:REF|PAGEREF|NOTEREF)\s+(\w+)',instr,re.I)
                if removed_ref and removed_ref[1] in removed_bookmarks:
                    frozen=freeze_field(f)
                    converted_reference_nodes.update(t for run in frozen for t in run.findall('.//w:t',NS))
                    changes.append(dict(operation='freeze_excluded_frontmatter_reference',field_index=fi,instruction=instr,old=old,new=old,bookmark=removed_ref[1]))
                    self.exceptions.append(dict(source_id=sid,block_id=r['block_id'],kind='reference_to_excluded_frontmatter_retained',message='引用指向已排除的来源封面/目录，保留显示文字，待确定新指向'))
                elif SEQ_FIELD.fullmatch(instr) and (r.get('caption_retained') or (re.match(r'\s*SEQ\s+[图表](?=\s|$)',instr,re.I) and (seq is None or f['start'] is not seq['start']))):
                    # An orphan's old SEQ must not advance the real object series.
                    freeze_field(f)
                    changes.append(dict(operation='freeze_unbound_caption_seq',field_index=fi,instruction=instr,old=old,new=old))
                    self.exceptions.append(dict(source_id=sid,block_id=r['block_id'],kind='unbound_caption_seq_display_retained',message='未绑定实际对象的旧图表SEQ保留显示文字，不推进总稿图表计数'))
                elif seq is not None and any(f['start'] is c['start'] for c in seq['chapter_fields']):
                    freeze_field(f)
                    changes.append(dict(operation='freeze_caption_chapter_field',field_index=fi,instruction=instr,old=old,new=old))
                elif seq is not None and f['start'] is seq['start']:
                    new=re.search(r'\d+',target[r['block_id']])[0]
                    new_instr=' SEQ '+target[r['block_id']][0]+' \\* ARABIC '
                    changes.append(dict(operation='native_caption_seq',field_index=fi,instruction=instr,new_instruction=new_instr,old=old,new=new,target_block_id=r['block_id']))
                    update_field(f,new_instr,new)
                elif ref and ref[1] in field_targets:
                    info=field_targets[ref[1]];new_instr=' REF '+bmnames[ref[1]]+' '+ref[2]
                    changes.append(dict(operation='native_caption_ref',field_index=fi,instruction=instr,new_instruction=new_instr,old=old,new=info['number'],target_block_id=info['block_id'],bookmark=ref[1]))
                    converted_reference_nodes.update(f['cache_nodes']);update_field(f,new_instr,info['number'])
                elif re.match(r'\s*(?:REF|PAGEREF|NOTEREF)\b',instr,re.I):
                    frozen=freeze_field(f);converted_reference_nodes.update(t for run in frozen for t in run.findall('.//w:t',NS))
                    changes.append(dict(operation='freeze_unresolved_reference',field_index=fi,instruction=instr,old=old,new=old))
                    self.exceptions.append(dict(source_id=sid,block_id=r['block_id'],kind='reference_display_frozen',resolved_by_policy=True,message='未可靠迁移的引用已解除域，原显示内容及对象保留'))
            # Visible figure/table references are changed only with a unique
            # source-local known target. No edits inside URL targets or fields.
            for p in ([node] if node.tag==q('w:p') else node.findall('.//w:p',NS)):
                if role in ('figure_caption','table_caption','heading','document_title') or r.get('figure_layout'):continue
                text=ordinary_text(p)
                matches=list(CAPTION_REFERENCE.finditer(text))
                for match in reversed(matches):
                    if r.get('caption_retained'):
                        own=caption_prefix(text)
                        if own and match.start()<own['span'][1]:
                            self.exceptions.append(dict(source_id=sid,block_id=r['block_id'],kind='unmapped_object_reference_preserved',text=match.group(0),
                                message='未确定题名中的旧图表号按原文保留，未宣称已迁移到总稿对象号'))
                            continue
                    label=match['label'];old=caption_number_key(match['number']);new=caption_maps[label].get(old);info=caption_by_old[label].get(old)
                    if protected_reference_span(p,match.start(2),match.end(2),set(field_targets)|removed_bookmarks,converted_reference_nodes):continue
                    if new is not None and info:
                        edit(p,match.start(2),match.end(2),new,'mapped_object_reference')
                        changes[-1]['target_block_id']=info['block_id']
                        instruction=' REF '+info['bookmark_name']+' \\h '
                        try:promote_text_field(p,match.start(2),match.start(2)+len(new),instruction,new)
                        except ValueError:
                            self.exceptions.append(dict(kind='plain_reference_native_conversion_retained',source_id=sid,block_id=r['block_id'],message='文字引用已按唯一对象改号，但其复杂字符布局未转为REF，后续编辑需核对'))
                        else:changes.append(dict(operation='promote_caption_reference',paragraph_index=0 if node.tag==q('w:p') else node.findall('.//w:p',NS).index(p),start=match.start(2),end=match.start(2)+len(new),instruction=instruction,new=new,bookmark_name=info['bookmark_name'],target_block_id=info['block_id']))
                    elif new is None:self.exceptions.append({'source_id':sid,'source_index':idx,'kind':'unmapped_object_reference_preserved','text':match.group(0)})
            if node.tag==q('w:p') and role not in ('heading','figure_caption','table_caption','document_title'):
                for match in reversed(list(re.finditer(r'(?:第|参见|详见|见)\s*(\d+(?:\.\d+){0,4})(?=\s*(?:小节|节|章))',ordinary_text(node)))):
                    new=heading_refs.get(match[1])
                    if new and new!=match[1] and not protected_reference_span(node,match.start(1),match.end(1),set(field_targets)|removed_bookmarks,converted_reference_nodes):edit(node,match.start(1),match.end(1),new,'mapped_heading_reference')
                    elif new is None:self.exceptions.append(dict(source_id=sid,block_id=r['block_id'],kind='unmapped_heading_reference_preserved',text=match[0]))
            if node.tag==q('w:tbl'):
                if r.get('figure_layout'):style_figure_table(node,self.styles,self.width,source_styles)
                else:style_table(node,self.styles,self.width,source_styles)
            elif node.tag!=q('w:p'):
                # Keep opaque containers and normalize their nested prose.
                for table in node.findall('.//w:tbl',NS):style_table(table,self.styles,self.width,source_styles)
                for p in node.findall('.//w:p',NS):
                    if p.xpath('ancestor::w:tc',namespaces=NS):continue
                    style_paragraph(p,self.styles['图形段' if pictures(p) else 'Normal'],source_styles=source_styles)
            elif node.tag==q('w:p'):
                style=ROLE_STYLE.get(role,'Normal');level=r.get('level')
                if role=='document_title' and title_info and r['block_id']==title_info['block_id'] and ordinary_text(node)!=title_info['text']:
                    edit(node,0,len(ordinary_text(node)),title_info['text'],'source_title_presentation')
                if role=='document_title' and self.master and display_title and ordinary_text(node)!=display_title:
                    edit(node,0,len(ordinary_text(node)),display_title,'source_display_title')
                if role=='heading_retained':
                    style='Heading '+str(level+source_depth) if type(level) is int and 1<=level<=5 else 'Normal'
                    if style not in self.styles:
                        self.exceptions.append(dict(source_id=sid,block_id=r['block_id'],kind='retained_heading_depth_unavailable',message='保留原号的标题超出目标样式级数，按正文保留，未删除原号或内容'))
                        style='Normal'
                if role=='heading':
                    style='Heading '+str(level+source_depth)
                    if style not in self.styles:raise ValueError('Target heading level has no approved style')
                    permission=r['number_edit'];lo,hi=permission['span']
                    edit(node,lo,hi,'' if self.chapter_num_id else target[r['block_id']]+' ','heading_number')
                    changes[-1]['authorization']=permission
                    if self.chapter_num_id:changes[-1]['native_number']=target[r['block_id']]
                    if permission['kind']=='automatic':changes.append({'operation':'detach_automatic_heading_numbering','source_num_id':permission['automatic']['num_id'],'deleted_text_characters':0})
                elif role=='document_title' and self.master:
                    style=group_style;edit(node,0,0,'' if self.chapter_num_id else '.'.join(map(str,outline_prefix))+' ','source_group_number')
                    if self.chapter_num_id:changes[-1]['native_number']='.'.join(map(str,outline_prefix))
                elif role=='list_item' and r['block_id'] in manual_lists:
                    plan=manual_lists[r['block_id']];permission=plan['authorization'];lo,hi=permission['span']
                    edit(node,lo,hi,'','manual_list_prefix');changes[-1]['authorization']=permission
                    changes[-1]['native_list']=dict(num_id=plan['num_id'],ilvl=plan['level'],source_numbers=permission['numbers'],form=permission['form'])
                elif role in ('figure_caption','table_caption') and r['block_id'] in target:
                    if seq is not None:
                        current=ordinary_text(node);prefix=caption_prefix(current);active=caption_seq(node)
                        if prefix is None or active is None:raise ValueError('Caption ordinal lost its bound SEQ')
                        a,b=active['text_span']
                        edit(node,b,prefix['span'][1],'　','caption_number_after_seq')
                        edit(node,0,a,target[r['block_id']][0]+' ','caption_number_before_seq')
                    else:
                        prefix=caption_prefix(before);number=caption_info[r['block_id']]['number']
                        if prefix:
                            a,b=prefix['number_span']
                            edit(node,b,prefix['span'][1],'　','caption_text_separator')
                            edit(node,a,b,number,'caption_text_number')
                            edit(node,0,a,target[r['block_id']][0]+' ','caption_text_label')
                        else:edit(node,0,0,target[r['block_id']],'caption_number')
                        instruction=' SEQ '+target[r['block_id']][0]+' \\* ARABIC '
                        promote_text_field(node,2,2+len(number),instruction,number)
                        changes.append(dict(operation='promote_caption_number',paragraph_index=0,start=2,end=2+len(number),instruction=instruction,new=number,target_block_id=r['block_id']))
                elif role=='equation' and re.fullmatch(r'\s*[（(]\s*\d+\s*[）)]\s*',before):
                    number=re.search(r'\d+',before).group(0)
                    if number in equation_refs:
                        self.exceptions.append(dict(source_id=sid,block_id=r['block_id'],kind='referenced_equation_number_retained',number=number,message='来源公式仍被引用，本处暂保留原尾号及引用；未生成新公式号，待明确迁移后再去号'))
                    else:edit(node,0,len(before),'','remove_equation_number_tail')
                elif role=='equation_description':
                    match=re.match(r'^\s*式\s*\d+\s*',before)
                    if match and match.group(0).strip().lstrip('式').strip() not in equation_refs:edit(node,0,match.end(),'','remove_unused_equation_label')
                elif role=='list_item' and actual_numbering.get(sr.getroottree().getpath(src),{}).get('status') in (None,'none') and not re.match(r'^\s*(?:[•·●]|\d+[、.)）])',before):
                    source_style=src.find('w:pPr/w:pStyle',NS)
                    if source_style is not None and 'bullet' in source_style.get(q('w:val'),'').lower():edit(node,0,0,'• ','materialize_inherited_bullet')
                pp=style_paragraph(node,self.styles[style],preserve_emphasis=role not in ('heading','heading_retained','document_title','front_heading','figure_caption','table_caption'),source_styles=source_styles)
                if self.chapter_num_id and (role=='heading' or (role=='document_title' and self.master)):
                    number_properties(pp,self.chapter_num_id,(level-1+source_depth) if role=='heading' else source_depth-1)
                elif self.template_style_numbers.get(style) not in (None,'0'):number_properties(pp,0)
                if role=='list_item' and r['block_id'] in manual_lists:
                    plan=manual_lists[r['block_id']];number_properties(pp,plan['num_id'],plan['level'])
                actual_auto=actual_numbering.get(sr.getroottree().getpath(src),{})
                if role=='list_item' and r['block_id'] not in manual_lists and actual_auto.get('status') in (None,'none') and not any(x.get('block_id')==r['block_id'] and 'retained' in x.get('kind','') for x in self.exceptions):
                    self.exceptions.append(dict(kind='list_native_conversion_retained',source_id=sid,block_id=r['block_id'],message='列表身份保留，编号前缀或分组尚未取得迁号授权，保留原文字'))
                if role!='heading' and actual_auto.get('status') not in (None,'none') and actual_auto.get('num_id') not in safe_unknown|set(numbering_notices):
                    unresolved=actual_auto.get('num_id') in safe_unknown
                    if actual_auto.get('status')!='resolved' and not unresolved:raise ValueError('Automatic numbering preservation precondition failed')
                    nid=actual_auto['num_id'];values=actual_auto.get('values');counter_key=(nid,actual_auto['ilvl']);previous=retained_numbers.get(counter_key)
                    if unresolved:values=None
                    if not unresolved and previous and values[:-1]==previous[:-1] and values[-1]!=previous[-1]+1:
                        # A removed heading or a section selection may leave a gap.
                        # Start a new instance at the original visible ordinal.
                        nummap.pop(nid,None)
                    new_id=self.import_numbering(sp,nid,nummap,values,preserve_definition=unresolved);retained_numbers[counter_key]=values
                    if unresolved and not any(x['source_id']==sid and x['source_num_id']==nid for x in self.numbering_retention):
                        self.numbering_retention.append(dict(source_id=sid,source_num_id=nid,target_num_id=new_id,mode='complete_instance_definition_preserved'))
                        self.exceptions.append(dict(source_id=sid,block_id=r['block_id'],kind='automatic_numbering_definition_retained',message='自动编号定义和完整序列保留；未解释或按目标章节改号'))
                    number_properties(pp,new_id,actual_auto['ilvl'])
                    pp.append(elt('w:ind',left=0,right=0,firstLine=0,hanging=0,leftChars=0,rightChars=0,firstLineChars=0,hangingChars=0))
                    changes.append({'operation':'preserve_automatic_list_numbering','label_resolved':not unresolved,'source_num_id':actual_auto['num_id'],'target_num_id':new_id,'source_label':actual_auto['label'],'deleted_text_characters':0})
                # Layout belongs to the target contract, not the source report.
                for flag,value in (('pageBreakBefore',0),('keepNext',int(role in ('heading','heading_retained','document_title','front_heading','figure','table_caption'))),('keepLines',int(role in ('heading','heading_retained','document_title','front_heading','figure','figure_caption','table_caption','equation')))):
                    pp.append(elt('w:'+flag,val=value))
                if role=='document_title' and not self.master:self.cancel_title_rule(pp)
                if role in ('front_heading','equation_description','table_text'):pp.append(elt('w:ind',left=0,right=0,firstLine=0,firstLineChars=0,leftChars=0,rightChars=0))
                if role=='equation_description':pp.append(elt('w:spacing',before=0,after=120,line=300,lineRule='auto'))
                if not before.strip() and not node.xpath('.//m:oMath|.//w:drawing|.//w:pict',namespaces=NS):
                    breaks=node.xpath('.//w:br[@w:type="page"]',namespaces=NS)
                    if breaks:
                        changes.append({'operation':'remove_standalone_page_break','count':len(breaks)})
                        for br in breaks:br.getparent().remove(br)
                if role=='equation' and not is_legacy_equation(node) and not node.findall('m:oMathPara',NS):
                    maths=node.findall('m:oMath',NS)
                    if not maths or ordinary_text(node).strip():
                        if not node.findall('.//m:oMath',NS):raise ValueError('Equation candidate has no original math object')
                        self.exceptions.append(dict(source_id=sid,block_id=r['block_id'],kind='equation_layout_retained',
                            reason='nested_native_math_carrier' if not maths else 'retained_visible_content',
                            message='原公式和现有文字保留；本处未强行转换行间容器，不影响其余内容入稿'))
                    else:
                        wrapper=E.Element(q('m:oMathPara'));pr=E.SubElement(wrapper,q('m:oMathParaPr'));jc=E.SubElement(pr,q('m:jc'));jc.set(q('m:val'),'center')
                        index=node.index(maths[0])
                        for math in maths:node.remove(math);wrapper.append(math)
                        node.insert(index,wrapper);changes.append({'operation':'promote_standalone_math','count':len(maths)})
            if r.get('figure_layout',{}).get('kind')=='inline-row':
                node=picture_row_table(node,r['figure_layout']['columns'])
                style_figure_table(node,self.styles,self.width,source_styles)
            nested_styles=normalize_embedded_story_styles(node,self.styles['Normal'],source_styles)
            if nested_styles:changes.append(dict(operation='normalize_embedded_story_styles',paragraphs=nested_styles,target_style_id=self.styles['Normal']))
            # Unknown numbering is handled for every paragraph in the retained
            # block, including table cells. Never drop the whole source for it.
            source_ps=list(src.iter(q('w:p')))
            target_ps=list(node.iter(q('w:p')))
            for pi,(source_p,target_p) in enumerate(zip(source_ps,target_ps)):
                facts=actual_numbering.get(sr.getroottree().getpath(source_p),{});nid=facts.get('num_id')
                if nid not in safe_unknown and nid not in numbering_notices:continue
                pp=target_p.find('w:pPr',NS)
                if pp is None:pp=elt('w:pPr');target_p.insert(0,pp)
                if nid in safe_unknown:
                    new_id=self.import_numbering(sp,nid,nummap,preserve_definition=True)
                    number_properties(pp,new_id,facts['ilvl'])
                    if not any(x['source_id']==sid and x['source_num_id']==nid for x in self.numbering_retention):
                        self.numbering_retention.append(dict(source_id=sid,source_num_id=nid,target_num_id=new_id,mode='complete_instance_definition_preserved'))
                        self.exceptions.append(dict(source_id=sid,block_id=r['block_id'],kind='automatic_numbering_definition_retained',message='自动编号定义及完整序列原样保留，未解释为章节；其他内容继续入稿'))
                    changes.append(dict(operation='preserve_automatic_list_numbering',paragraph_index=pi,label_resolved=False,source_num_id=nid,target_num_id=new_id,source_label=None,deleted_text_characters=0))
                else:
                    number_properties(pp,0)
                    prefix='[原自动编号待核] '
                    text_edit(target_p,0,0,prefix)
                    changes.append(dict(operation='unresolved_numbering_notice',paragraph_index=pi,source_num_id=nid,reason=numbering_notices[nid],prefix=prefix))
                    self.exceptions.append(dict(source_id=sid,block_id=r['block_id'],kind='automatic_numbering_label_unavailable',message='程序未可靠解析本处旧自动编号，以明确提示代替；原文字/对象保留，原编号定义见冻结来源',reason=numbering_notices[nid],source_num_id=nid))
            # Names and IDs are session-independent package identifiers here.
            for link in reversed(node.findall('.//w:hyperlink',NS)):
                anchor=link.get(q('w:anchor'))
                if anchor not in removed_bookmarks:continue
                parent=link.getparent();at=parent.index(link)
                for child in list(link):parent.insert(at,child);at+=1
                parent.remove(link);changes.append(dict(operation='unlink_excluded_frontmatter_anchor',anchor=anchor))
                self.exceptions.append(dict(source_id=sid,block_id=r['block_id'],kind='link_to_excluded_frontmatter_retained',message='来源前置导航链接已解除，显示文字保留'))
            for n in node.iter():
                if n.tag in (q('v:shape'),q('v:shapetype')) and n.get('id') in legacy_ids:n.set('id',legacy_ids[n.get('id')])
                if n.tag==q('v:shape') and n.get('type','').startswith('#') and n.get('type')[1:] in legacy_ids:n.set('type','#'+legacy_ids[n.get('type')[1:]])
                if n.tag==q('o:OLEObject'):
                    for attr in ('ShapeID','ObjectID'):
                        if n.get(attr) in legacy_ids:n.set(attr,legacy_ids[n.get(attr)])
                if n.tag in (q('w:bookmarkStart'),q('w:bookmarkEnd')):
                    old=n.get(q('w:id'))
                    if old not in bmids:raise ValueError('Unmatched source bookmark')
                    n.set(q('w:id'),bmids[old])
                    if n.tag==q('w:bookmarkStart'):n.set(q('w:name'),bmnames[n.get(q('w:name'))])
                if n.tag==q('w:hyperlink') and n.get(q('w:anchor')) in bmnames:n.set(q('w:anchor'),bmnames[n.get(q('w:anchor'))])
                original_attributes=dict(n.attrib)
                for key,value in list(original_attributes.items()):
                    if not key.startswith('{'+NS['r']+'}'):continue
                    relation=rels.get(value);kind=relation.get('Type','').rsplit('/',1)[-1] if relation is not None else ''
                    image=image_attribute(n,key)
                    if n.tag==q('a:blip') and key==q('r:link'):
                        embedded=rels.get(original_attributes.get(q('r:embed')))
                        if embedded is not None and embedded.get('TargetMode')!='External' and part_name(embedded.get('Target','')) in sp:
                            changes.append(dict(operation='drop_image_link',path=relative_path(node,n),relationship_id=value))
                            n.attrib.pop(key,None);continue
                    if value not in relmap:
                        missing=image and (relation is None or relation.get('TargetMode')=='External' or part_name(relation.get('Target','')) not in sp)
                        if missing:
                            data=missing_png();digest=hashlib.sha256(data).hexdigest();dest='word/media/missing_'+digest[:20]+'.png';self.parts[dest]=data
                            original=dict(relation.attrib) if relation is not None else None
                            repaired=E.Element('{'+REL+'}Relationship',Type=NS['r']+'/image',Target=posixpath.relpath(dest,'word'))
                            newrel=self.add_rel(repaired)
                            if not any(t.get('Extension')=='png' for t in self.types):E.SubElement(self.types,'{'+CT+'}Default',Extension='png',ContentType='image/png')
                            missing_repairs[value]=dict(relationship_id=value,original=original,placeholder_sha256=digest)
                        elif kind=='hyperlink':newrel=self.add_rel(relation)
                        elif kind in ('image','oleObject','package') and relation.get('TargetMode')!='External':
                            srcpart=part_name(relation.get('Target',''));data=sp[srcpart];digest=hashlib.sha256(data).hexdigest();ext=Path(srcpart).suffix.lower()
                            dest=('word/media/batch_'+digest[:24]+ext if kind=='image' else 'word/embeddings/'+sid+'_'+hashlib.sha256(value.encode()).hexdigest()[:8]+'_'+digest[:16]+ext);self.parts[dest]=data
                            newrel=self.add_rel(relation);newrel.set('Target',posixpath.relpath(dest,'word'))
                            if not any(t.get('Extension')==ext[1:] for t in self.types):
                                typ=next((t for t in s_types if t.get('Extension')==ext[1:]),None)
                                if typ is not None:self.types.append(deepcopy(typ))
                                else:
                                    override=next((t for t in s_types if t.get('PartName')=='/'+srcpart),None)
                                    if override is None:raise ValueError('Related part content type missing: '+srcpart)
                                    E.SubElement(self.types,'{'+CT+'}Override',PartName='/'+dest,ContentType=override.get('ContentType'))
                        else:raise ValueError('Unsupported relationship '+kind+'; source='+sid+'; block='+r['block_id'])
                        relmap[value]=newrel.get('Id')
                    if image and key==q('r:link') and value in missing_repairs and not original_attributes.get(q('r:embed')):
                        n.set(q('r:embed'),relmap[value]);n.attrib.pop(key,None)
                        changes.append(dict(operation='missing_link_placeholder_as_embed',path=relative_path(node,n),relationship_id=value))
                    else:n.set(key,relmap[value])
                    if value in missing_repairs and not any(c.get('operation')=='missing_media_notice' and c.get('relationship_id')==value for c in changes):
                        owner=n
                        while owner is not node and owner.tag!=q('w:p'):owner=owner.getparent()
                        if owner.tag==q('w:p'):
                            ps=[node] if node.tag==q('w:p') else node.findall('.//w:p',NS);pi=ps.index(owner);at=len(ordinary_text(owner))
                            message=' [图片资料缺失：'+sid+':'+r['block_id']+'；关系'+value+']'
                            text_edit(owner,at,at,message);changes.append(dict(operation='missing_media_notice',paragraph_index=pi,start=at,end=at,old='',new=message,relationship_id=value,source_id=sid,block_id=r['block_id']))
                        self.exceptions.append(dict(source_id=sid,block_id=r['block_id'],kind='missing_image_data',message='源图片部件缺失，仅本处用明确占位，文字及其他对象继续入稿',relationship_id=value))
            for color in node.xpath('.//m:oMath//w:color',namespaces=NS):
                if dict(color.attrib)!={q('w:val'):'000000'}:
                    self.math_colors.append({'source_id':sid,'source_index':idx,'before':dict(color.attrib),'after':'000000'});color.attrib.clear();color.set(q('w:val'),'000000')
            for inline in node.findall('.//wp:inline',NS):
                self.docpr+=1;dp=inline.find(q('wp:docPr'))
                if dp is not None:dp.set('id',str(self.docpr))
                ex=inline.find(q('wp:extent'))
                if ex is None:continue
                cx=int(ex.get('cx'));cy=int(ex.get('cy'));limit=self.width
                if cx<=0 or cy<=0:continue
                cells=inline.xpath('ancestor::w:tc',namespaces=NS)
                if cells:
                    cw=cells[-1].find('w:tcPr/w:tcW',NS)
                    if cw is not None and cw.get(q('w:type'))=='dxa':limit=min(limit,max(200,int(cw.get(q('w:w')))-170))
                scale=min(1.,limit*635/cx)
                if scale<1:
                    for extent in [ex]+inline.xpath('.//a:xfrm/a:ext',namespaces=NS):extent.set('cx',str(round(cx*scale)));extent.set('cy',str(round(cy*scale)))
                self.images.append({'source_id':sid,'source_index':idx,'scale':scale})
            if r['block_id'] in caption_info and caption_info[r['block_id']]['create_bookmark']:
                info=caption_info[r['block_id']];view=sequence_field(node)
                if view is None:raise ValueError('Managed caption lost its sequence field')
                self.mark_id+=1;ident=str(self.mark_id);wrap_number_bookmark(view,ident,info['bookmark_name'])
                changes.append(dict(operation='add_caption_number_bookmark',id=ident,name=info['bookmark_name']))
            oi=len(self.body)-1;self.body.insert(oi,node)
            self.records.append({'source_id':sid,'source_index':idx,'block_id':r['block_id'],'output_index':oi,'role':role,'level':r.get('level'),'decision_status':r['status'],'relates_to':r.get('relates_to'),'source_text':before,'output_text':ordinary_text(node),'changes':changes,'section_boundary':r.get('section_boundary'),'number_group_id':r.get('number_group_id'),'figure_layout':r.get('figure_layout'),'subcaptions':r.get('subcaptions',False),'retained_caption_role':r.get('retained_caption_role')})
        self.materialize_object_numbers(sid,objects)
        self.sources.append({'source_id':sid,'path':str(source.resolve()),'sha256':sha(source),'candidate':str(Path(candidate).resolve()),'candidate_sha256':sha(candidate),'source_number':source_number,'source_prefix':outline_prefix,'bookmark_id_map':bmids,'bookmark_name_map':bmnames,'relationship_map':relmap,'caption_maps':caption_maps,'retained_caption_roles':sorted(held_caption_roles),'heading_reference_map':heading_refs,'caption_field_targets':field_targets,'generated_group_title':group_title,'source_frontmatter':front,'excluded_bookmark_names':sorted(removed_bookmarks),'missing_media_repairs':list(missing_repairs.values()),'legacy_id_map':legacy_ids})
        self.sources[-1]['bookmark_detachments']=detached_bookmarks

    def materialize_object_numbers(self,sid,objects):
        # Keep element identities while inserting/moving caption paragraphs.
        # Numeric body offsets are recomputed only after all moves in this source.
        refs=[(r,self.body[r['output_index']]) for r in self.records+self.extra]
        mapped={r['block_id']:(r,self.body[r['output_index']]) for r in self.records if r['source_id']==sid}
        pointers=[]
        for obj in objects:
            row,root=mapped[obj['block_id']]
            physical=root if obj['source_path']=='.' else root.xpath(obj['source_path'],namespaces=NS)[0]
            pointers.append((obj,row,root,physical))
        captions=[]
        for obj,row,root,physical in pointers:
            caption_id=obj['caption_block_id']
            if caption_id:
                cap=mapped[caption_id][1]
                cap.getparent().remove(cap)
            else:
                cap=elt('w:p');run=elt('w:r');t=elt('w:t');t.text=obj['label']+' ';run.append(t);cap.append(run)
                field=elt('w:fldSimple',instr=' SEQ '+obj['label']+' \\* ARABIC ',dirty=1)
                run=elt('w:r');t=elt('w:t');t.text=obj['number'];run.append(t);field.append(run);cap.append(field)
                self.mark_id+=1;wrap_number_bookmark(sequence_field(cap),str(self.mark_id),obj['bookmark_name'])
                pp=style_paragraph(cap,self.styles['图题' if obj['label']=='图' else '表题'],False)
                number_properties(pp,0)
                for flag,value in [('pageBreakBefore',0),('keepNext',int(obj['label']=='表')),('keepLines',1)]:pp.append(elt('w:'+flag,val=value))
                if physical.getparent() is self.body:
                    generated=dict(kind='object_number_caption',source_id=sid,object_id=obj['object_id'],
                                   label=obj['label'],number=obj['number'],text=ordinary_text(cap),output_index=-1)
                    self.extra.append(generated);refs.append((generated,cap))
            parent=physical.getparent();parent.insert(parent.index(physical)+(obj['label']=='图'),cap)
            captions.append((obj,row,root,physical,cap))
        for item,node in refs:item['output_index']=self.body.index(node)
        for obj,row,root,physical,cap in captions:
            top=cap
            while top.getparent() is not self.body:top=top.getparent()
            obj.update(object_output_index=self.body.index(root),object_path=relative_path(root,physical),
                       caption_output_index=self.body.index(top),caption_path=relative_path(top,cap),
                       generated=obj['caption_block_id'] is None)
            self.object_numbers.append(obj)
        # Nested captions are additions, not changes to technical cell content.
        for row,root in mapped.values():row['output_text']=ordinary_text(root)
    def add_gap(self,source,sid,reason,source_number=1,source_title=None,source_prefix=None):
        gap_prefix=source_prefix or [source_number]
        label=source_title or Path(source).name
        if self.master and self.chapter_num_id:
            p=elt('w:p');run=elt('w:r');t=elt('w:t');t.text=label;run.append(t);p.append(run)
            pp=style_paragraph(p,self.styles['Heading '+str(len(gap_prefix))],False);number_properties(pp,self.chapter_num_id,len(gap_prefix)-1)
            index=len(self.body)-1;self.body.insert(index,p)
            self.extra.append(dict(kind='source_gap_heading',source_id=sid,source_number=source_number,source_prefix=gap_prefix,output_index=index,text=label))
        text=f'【内容缺口：{label}（{sid}）未搬入；请查看交付清单中的原件。】'
        p=elt('w:p');run=elt('w:r');t=elt('w:t');t.text=text;run.append(t);p.append(run)
        style_paragraph(p,self.styles['Normal'],False)
        index=len(self.body)-1;self.body.insert(index,p)
        self.extra.append(dict(kind='source_gap',source_id=sid,source_number=source_number,output_index=index,text=text))
        self.gaps.append(dict(source_id=sid,source_number=source_number,path=str(Path(source).resolve()),sha256=sha(source),reason=reason,output_index=index))

    def save(self,path):
        path=Path(path);path.parent.mkdir(parents=True,exist_ok=True)
        self.parts['word/document.xml']=pack(self.root);self.parts['word/_rels/document.xml.rels']=pack(self.rels);self.parts['[Content_Types].xml']=pack(self.types)
        if self.numbering_changed:self.parts['word/numbering.xml']=pack(self.numbering_root)
        with ZipFile(path,'x',ZIP_DEFLATED) as z:
            for name,data in self.parts.items():z.writestr(name,data)
        report={'version':'0.5.7','target_contract':'continuous-body-v1','template_version':self.template_version,'route':'OOXML_TEMPLATE_ASSEMBLY','wps_status':'NOT_RUN','visual_status':'SKIPPED_BY_USER','template':str(self.template.resolve()),'template_sha256':sha(self.template),'output':str(path.resolve()),'output_sha256':sha(path),'master':self.master,'sources':self.sources,'mapping':self.records,'exclusions':self.exclusions,'generated_blocks':self.extra,'exceptions':self.exceptions,'image_scaling':self.images,'math_color_normalizations':self.math_colors,'preserve_only_parts_changed':[n for n,b in self.original_parts.items() if n not in ('word/document.xml','word/_rels/document.xml.rels','[Content_Types].xml') and self.parts.get(n)!=b],'scope':'Draft output with source-bound transformations; no claim of independently approved structure or WPS acceptance.'}
        if self.template_profile['body_slot']['kind'] in ('marked','bookmark'):report['target_contract']='preserved-frontmatter-v1'
        report['template_body_slot']=self.template_profile['body_slot']
        report['numbering_contract']=CONTRACT;report['numbering_definitions_imported']=self.numbering_changed
        report['numbering_retention']=self.numbering_retention;report['gaps']=self.gaps
        report['native_chapter_num_id']=self.chapter_num_id
        report['object_numbering_contract']='physical-objects-v1';report['object_numbers']=self.object_numbers
        report['caption_role_contract']='source-layout-v1'
        report['status']='DRAFT_WITH_UNRESOLVED_STRUCTURE' if any(x['kind']=='unresolved_structure_preserved' for x in self.exceptions) else 'GENERATED'
        report['preserve_only_parts_changed']=[n for n in report['preserve_only_parts_changed'] if n!='word/numbering.xml' or not self.numbering_changed]
        path.with_suffix('.assembly.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf8');return report
