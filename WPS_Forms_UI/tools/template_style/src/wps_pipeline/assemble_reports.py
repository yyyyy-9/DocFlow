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
from .docx_formatting import NS,q,xml,pack,elt,ordinary_text,style_paragraph,style_table,style_figure_table,picture_row_table,sha
from .resolve_styles import inspect_template
from .numbering_facts import NumberingFacts
from .numbered_candidates import validate_edit,validate_list_edit,CONTRACT,select_section
from .native_output import number_properties,manual_list_definition,field_views,update_field,promote_text_field,sequence_field,number_bookmarks,wrap_number_bookmark

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
        self.records=[];self.exclusions=[];self.sources=[];self.exceptions=[];self.images=[];self.math_colors=[];self.caption_count=Counter();self.mark_id=0;self.docpr=0;self.extra=[]
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
        root=xml(parts['word/numbering.xml']);old=root.find('w:num[@w:numId="'+nid+'"]',NS)
        if old is None:raise ValueError('Missing source numbering definition')
        aid=old.find('w:abstractNumId',NS).get(q('w:val'));abstract=root.find('w:abstractNum[@w:abstractNumId="'+aid+'"]',NS)
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

    def add_source(self,source,candidate,sid,source_number=1,section_id=None,source_title=None):
        source=Path(source);pred=json.loads(Path(candidate).read_text(encoding='utf8'))
        if pred['source_sha256']!=sha(source):raise ValueError('Source/candidate hash mismatch')
        if pred.get('numbering_contract')!=CONTRACT:raise ValueError('Old decisions need a fresh scan and contextual recognition before number edits')
        if any(r['status']=='review' for r in pred['blocks']) and not self.allow_pending:raise ValueError('Explicit local resolution required; whole-document draft may preserve pending blocks')
        with ZipFile(source) as z:sp={n:z.read(n) for n in z.namelist()}
        sr=xml(sp['word/document.xml']);sb=sr.find(q('w:body'));rels={x.get('Id'):x for x in xml(sp['word/_rels/document.xml.rels'])};s_types=xml(sp['[Content_Types].xml'])
        source_styles={s.get(q('w:styleId')):s for s in xml(sp['word/styles.xml']).findall('w:style',NS)}
        rows=deepcopy(pred['blocks']);num=[0]*5;target={};caption_maps={'图':{},'表':{}};relmap={};bmids={};bmnames={};nummap={};retained_numbers={}
        nf=NumberingFacts(sp);actual_numbering={sr.getroottree().getpath(p):nf.inspect(p) for p in sb.iter(q('w:p'))}
        if section_id:
            selected=set(select_section(pred,section_id));rows=[r for r in rows if r['block_id'] in selected]
        unknown_ids={a['num_id'] for a in actual_numbering.values() if a.get('status')=='unresolved' and a.get('num_id')}
        safe_unknown=set()
        if unknown_ids:
            numbering=xml(sp['word/numbering.xml'])
            row_by_index={r['source_locator']['body_child_index']:r for r in rows}
            for nid in unknown_ids:
                definition=numbering.find('w:num[@w:numId="'+nid+'"]',NS)
                if definition is None:raise ValueError('Missing numbering definition; original source retained outside draft')
                aid=definition.find('w:abstractNumId',NS)
                abstract=numbering.find('w:abstractNum[@w:abstractNumId="'+aid.get(q('w:val'))+'"]',NS) if aid is not None else None
                if abstract is None or any(n.tag in (q('w:numStyleLink'),q('w:styleLink'),q('w:pStyle'),q('w:lvlPicBulletId')) or any(k.startswith('{'+NS['r']+'}') for k in n.attrib) for base in (abstract,definition) for n in base.iter()):
                    raise ValueError('Numbering preservation requires a self-contained definition')
                for p in sb.iter(q('w:p')):
                    auto=actual_numbering[sr.getroottree().getpath(p)]
                    if auto.get('num_id')!=nid:continue
                    if p.getparent() is not sb or sb.index(p) not in row_by_index:
                        raise ValueError('Unresolved numbering must retain its complete top-level sequence')
                    item=row_by_index[sb.index(p)]
                    if item['role'] not in ('heading','body','list_item'):
                        raise ValueError('Unresolved numbering sequence intersects a moved or special block')
                    if item['role']=='heading':
                        item.update(retained_role='heading',retained_level=item.get('level'),status='review')
                safe_unknown.add(nid)
        # Pairing is constrained by physical source XML, independently of the
        # model receiver. An unresolved family keeps its old numbers together.
        row_by_id={r['block_id']:r for r in rows};held_caption_roles=set()
        for r in rows:
            role=r['role']
            if role not in ('figure_caption','table_caption'):continue
            related=r.get('relates_to');target_row=row_by_id.get(related)
            if related:
                if target_row is None or related==r['block_id']:raise ValueError('Caption target outside selected source or self-related')
                physical=sb[target_row['source_locator']['body_child_index']]
                valid=(role=='table_caption' and physical.tag==q('w:tbl') and not target_row.get('figure_layout')) or (role=='figure_caption' and bool(physical.findall('.//w:drawing',NS)) and (physical.tag==q('w:p') or (physical.tag==q('w:tbl') and target_row.get('figure_layout',{}).get('kind')=='table')))
                if not valid:raise ValueError('Caption target physical type mismatch')
            if not related or r['status'] in ('review','retained_unpaired'):
                held_caption_roles.add(role)
                self.exceptions.append(dict(source_id=sid,block_id=r['block_id'],kind='caption_association_unresolved',message='题名与对象未确定；保留原文，需确认关联后重新编号',candidates=r.get('association_candidates',[])))
            p=sb[r['source_locator']['body_child_index']]
            if (p.findall('.//w:fldSimple',NS) or p.findall('.//w:fldChar',NS)) and caption_seq(p) is None:held_caption_roles.add(role)
        for role in sorted(held_caption_roles):
            self.exceptions.append(dict(source_id=sid,kind='caption_numbering_group_retained',caption_role=role,block_ids=[r['block_id'] for r in rows if r['role']==role],message='本来源同类题名及其引用整体保留旧号，目标连续编号尚未完成；其他来源可继续整理'))
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
            label=source_title or source.stem
            if not isinstance(label,str) or not label.strip() or len(label)>200:
                raise ValueError('Source group title must contain 1..200 characters')
            text=label if self.chapter_num_id else f'{source_number} {label}'
            p=elt('w:p');run=elt('w:r');run.append(elt('w:t'));run[-1].text=text;p.append(run)
            pp=style_paragraph(p,self.styles['Heading 1'],False)
            if self.chapter_num_id:number_properties(pp,self.chapter_num_id,0)
            for flag,value in [('pageBreakBefore',0),('keepNext',1),('keepLines',1)]:pp.append(elt('w:'+flag,val=value))
            index=len(self.body)-1;self.body.insert(index,p)
            group_title={'source_id':sid,'source_number':source_number,'label':label,
                         'origin':'explicit_source_label_or_filename','text':text,'native_number':str(source_number) if self.chapter_num_id else None}
            self.extra.append(dict(group_title,kind='source_group_title',output_index=index))
        for m in sb.findall('.//w:bookmarkStart',NS):
            old=m.get(q('w:id'));self.mark_id+=1;bmids[old]=str(self.mark_id);bmnames[m.get(q('w:name'))]=f'{sid}_B{self.mark_id:05}'
        # Decide numbering before importing any block or touching references.
        for r in rows:
            idx=r['source_locator']['body_child_index'];text=ordinary_text(sb[idx]);role=r['role']
            if role=='heading':
                level=r['level'];num[level-1]+=1
                for k in range(level,5):num[k]=0
                if any(x==0 for x in num[:level]):raise ValueError('Heading has a missing parent')
                number='.'.join(map(str,num[:level]));target[r['block_id']]=f'{source_number}.{number}' if self.master else number
            elif role in ('figure_caption','table_caption') and r.get('relates_to') and role not in held_caption_roles and not r.get('subcaptions'):
                label='图' if role=='figure_caption' else '表';match=re.match(r'^\s*'+label+r'\s*(\d+(?:[-－.]\d+)?)',text)
                if match or not (sb[idx].findall('.//w:fldSimple',NS) or sb[idx].findall('.//w:fldChar',NS)):
                    p=sb[idx]
                    if (p.findall('.//w:fldSimple',NS) or p.findall('.//w:fldChar',NS)) and caption_seq(p) is None:
                        self.exceptions.append(dict(source_id=sid,block_id=r['block_id'],kind='caption_field_preserved',message='题名域不在已支持的SEQ范围，保留原显示及指令'))
                        continue
                    self.caption_count[label]+=1;new=str(self.caption_count[label]);old=match.group(1) if match else None
                    if old is not None:
                        if old in caption_maps[label]:raise ValueError('Duplicate caption number needs explicit mapping')
                        caption_maps[label][old]=new
                    target[r['block_id']]=label+' '+new+'　'
        field_targets={};heading_refs={};caption_info={};caption_by_old={'图':{},'表':{}}
        for r in rows:
            p=sb[r['source_locator']['body_child_index']]
            if r['role'] in ('figure_caption','table_caption') and r['block_id'] in target:
                f=caption_seq(p)
                label=target[r['block_id']][0];matched=re.match(r'^\s*[图表]\s*(\d+(?:[-－.]\d+)?)',ordinary_text(p));old=matched[1] if matched else None
                names=number_bookmarks(p,old) if old is not None else []
                name=bmnames[names[0]] if names else 'WPSN_'+hashlib.sha256((sid+':'+r['block_id']).encode()).hexdigest()[:24]
                info=dict(block_id=r['block_id'],number=re.search(r'\d+',target[r['block_id']])[0],old=old,bookmark_name=name,create_bookmark=not names)
                caption_info[r['block_id']]=info
                if old is not None:caption_by_old[label][old]=info
                for name in names:field_targets[name]=info
            if r['role']=='heading':
                o=r.get('number_candidate') or {};values=o.get('numbers')
                if values:
                    old='.'.join(map(str,values));new=target[r['block_id']]
                    heading_refs[old]=new if old not in heading_refs else None
        equation_refs=[];equation_descriptions={}
        for r in rows:
            if r['role']=='equation_description':
                text=ordinary_text(sb[r['source_locator']['body_child_index']]);match=re.match(r'^\s*式\s*(\d+)\s+',text)
                if match:equation_descriptions[match.group(1)]=r['source_locator']['body_child_index']
        for r in rows:
            if r['role'] not in ('equation_description','equation'):
                equation_refs.extend(re.findall(r'式\s*(\d+)',ordinary_text(sb[r['source_locator']['body_child_index']])))
        table_caps={r['relates_to']:r for r in rows if r['role']=='table_caption' and r.get('relates_to')}
        ordered=[]
        for r in rows:
            if r['role']=='table_caption' and r.get('relates_to'):continue
            if r['role']=='table' and r['block_id'] in table_caps:ordered.append(table_caps[r['block_id']])
            ordered.append(r)
        for r in ordered:
            idx=r['source_locator']['body_child_index'];src=sb[idx];role=r['role'];node=deepcopy(src);before=ordinary_text(src);changes=[]
            if role=='excluded_toc':
                raise ValueError('Legacy fixture exclusion is not authorized for general sources; rescan and preserve content')
            old_ppr=deepcopy(node.find(q('w:pPr'))) if node.tag==q('w:p') else None
            def edit(p,a,b,new,reason):
                pi=0 if node.tag==q('w:p') else node.findall('.//w:p',NS).index(p)
                old=ordinary_text(p)[a:b];text_edit(p,a,b,new);changes.append({'operation':reason,'paragraph_index':pi,'start':a,'end':b,'old':old,'new':new})
            if role not in ('equation','equation_description'):
                for match in reversed(list(re.finditer(r'式\s*(\d+)',before))):
                    target_index=equation_descriptions.get(match.group(1))
                    if node.tag!=q('w:p') or target_index!=idx-1:
                        raise ValueError('Equation reference requires explicit nonlocal migration: '+sid+':'+str(idx))
                    edit(node,match.start(),match.end(),'上述关系','adjacent_equation_reference')
                    changes[-1]['target_source_index']=target_index
            # Preserve native SEQ/REF identity and only update bound instructions/caches.
            seq=caption_seq(node) if role in ('figure_caption','table_caption') and r['block_id'] in target else None
            converted_reference_nodes=set()
            for fi,f in reversed(list(enumerate(field_views(node)))):
                instr=f['instruction'];old=f['cache'];ref=REF_FIELD.fullmatch(instr)
                if seq is not None and f['start'] is seq['start']:
                    new=re.search(r'\d+',target[r['block_id']])[0]
                    new_instr=' SEQ '+target[r['block_id']][0]+' \\* ARABIC '
                    changes.append(dict(operation='native_caption_seq',field_index=fi,instruction=instr,new_instruction=new_instr,old=old,new=new,target_block_id=r['block_id']))
                    update_field(f,new_instr,new)
                elif ref and ref[1] in field_targets:
                    info=field_targets[ref[1]];new_instr=' REF '+bmnames[ref[1]]+' '+ref[2]
                    changes.append(dict(operation='native_caption_ref',field_index=fi,instruction=instr,new_instruction=new_instr,old=old,new=info['number'],target_block_id=info['block_id'],bookmark=ref[1]))
                    converted_reference_nodes.update(f['cache_nodes']);update_field(f,new_instr,info['number'])
                elif re.match(r'\s*(?:REF|PAGEREF|NOTEREF)\b',instr,re.I):
                    match=re.match(r'(\s*(?:REF|PAGEREF|NOTEREF)\s+)([\w]+)(.*)',instr,re.I|re.S)
                    if match and match[2] in bmnames:
                        new=match[1]+bmnames[match[2]]+match[3];update_field(f,new,old)
                        changes.append(dict(operation='remap_native_reference_field',field_index=fi,instruction=instr,new_instruction=new,old=old,new=old,bookmark=match[2]))
                    self.exceptions.append(dict(source_id=sid,block_id=r['block_id'],kind='reference_field_cache_retained',message='域目标名称已迁移（若存在），显示缓存尚未重算'))
            # Visible figure/table references are changed only with a unique
            # source-local known target. No edits inside URL targets or fields.
            for p in ([node] if node.tag==q('w:p') else node.findall('.//w:p',NS)):
                if role in ('figure_caption','table_caption','heading','document_title') or r.get('figure_layout'):continue
                text=ordinary_text(p)
                matches=list(re.finditer(r'(图|表)\s*(\d+(?:[-－.]\d+)?)',text))
                for match in reversed(matches):
                    label,old=match.groups();new=caption_maps[label].get(old);info=caption_by_old[label].get(old)
                    if protected_reference_span(p,match.start(2),match.end(2),field_targets,converted_reference_nodes):continue
                    if new is not None and info:
                        edit(p,match.start(2),match.end(2),new,'mapped_object_reference')
                        instruction=' REF '+info['bookmark_name']+' \\h '
                        try:promote_text_field(p,match.start(2),match.start(2)+len(new),instruction,new)
                        except ValueError:
                            self.exceptions.append(dict(kind='plain_reference_native_conversion_retained',source_id=sid,block_id=r['block_id'],message='文字引用已按唯一对象改号，但其复杂字符布局未转为REF，后续编辑需核对'))
                        else:changes.append(dict(operation='promote_caption_reference',paragraph_index=0 if node.tag==q('w:p') else node.findall('.//w:p',NS).index(p),start=match.start(2),end=match.start(2)+len(new),instruction=instruction,new=new,bookmark_name=info['bookmark_name'],target_block_id=info['block_id']))
                    elif new is None:self.exceptions.append({'source_id':sid,'source_index':idx,'kind':'unmapped_object_reference_preserved','text':match.group(0)})
            if node.tag==q('w:p') and role not in ('heading','figure_caption','table_caption','document_title'):
                for match in reversed(list(re.finditer(r'(?:第|参见|详见|见)\s*(\d+(?:\.\d+){0,4})(?=\s*(?:小节|节|章))',ordinary_text(node)))):
                    new=heading_refs.get(match[1])
                    if new and new!=match[1] and not protected_reference_span(node,match.start(1),match.end(1),field_targets,converted_reference_nodes):edit(node,match.start(1),match.end(1),new,'mapped_heading_reference')
                    elif new is None:self.exceptions.append(dict(source_id=sid,block_id=r['block_id'],kind='unmapped_heading_reference_preserved',text=match[0]))
            if node.tag==q('w:tbl'):
                if r.get('figure_layout'):style_figure_table(node,self.styles,self.width,source_styles)
                else:style_table(node,self.styles,self.width,source_styles)
            elif node.tag==q('w:p'):
                style=ROLE_STYLE.get(role,'Normal');level=r.get('level')
                if role=='heading_retained':
                    style='Heading '+str(level) if type(level) is int and 1<=level<=5 else 'Normal'
                if role=='heading':
                    style='Heading '+str(level+(1 if self.master else 0))
                    if style not in self.styles:raise ValueError('Target heading level has no approved style')
                    permission=r['number_edit'];lo,hi=permission['span']
                    edit(node,lo,hi,'' if self.chapter_num_id else target[r['block_id']]+' ','heading_number')
                    changes[-1]['authorization']=permission
                    if self.chapter_num_id:changes[-1]['native_number']=target[r['block_id']]
                    if permission['kind']=='automatic':changes.append({'operation':'detach_automatic_heading_numbering','source_num_id':permission['automatic']['num_id'],'deleted_text_characters':0})
                elif role=='document_title' and self.master:
                    style='Heading 1';edit(node,0,0,'' if self.chapter_num_id else str(source_number)+' ','source_group_number')
                    if self.chapter_num_id:changes[-1]['native_number']=str(source_number)
                elif role=='list_item' and r['block_id'] in manual_lists:
                    plan=manual_lists[r['block_id']];permission=plan['authorization'];lo,hi=permission['span']
                    edit(node,lo,hi,'','manual_list_prefix');changes[-1]['authorization']=permission
                    changes[-1]['native_list']=dict(num_id=plan['num_id'],ilvl=plan['level'],source_numbers=permission['numbers'],form=permission['form'])
                elif role in ('figure_caption','table_caption') and r['block_id'] in target:
                    if seq is not None:
                        current=ordinary_text(node);m=re.match(r'^(\s*[图表]\s*)(\d+)([.．、]?\s*)',current)
                        edit(node,m.end(2),m.end(3),'　','caption_separator')
                        edit(node,0,m.start(2),target[r['block_id']][0]+' ','caption_label')
                    else:
                        match=re.match(r'^\s*[图表]\s*\d+(?:[-－.]\d+)?[.．、]?\s*',before)
                        prefix=match.group(0) if match else '';edit(node,0,len(prefix),target[r['block_id']],'caption_number')
                        number=caption_info[r['block_id']]['number'];instruction=' SEQ '+target[r['block_id']][0]+' \\* ARABIC '
                        promote_text_field(node,2,2+len(number),instruction,number)
                        changes.append(dict(operation='promote_caption_number',paragraph_index=0,start=2,end=2+len(number),instruction=instruction,new=number,target_block_id=r['block_id']))
                elif role=='equation' and re.fullmatch(r'\s*[（(]\s*\d+\s*[）)]\s*',before):
                    number=re.search(r'\d+',before).group(0)
                    if number in equation_refs:
                        raise ValueError('Referenced equation tail needs explicit reference migration: '+sid+':'+number)
                    edit(node,0,len(before),'','remove_equation_number_tail')
                elif role=='equation_description':
                    match=re.match(r'^\s*式\s*\d+\s*',before)
                    if match:edit(node,0,match.end(),'','remove_unused_equation_label')
                elif role=='list_item' and actual_numbering.get(sr.getroottree().getpath(src),{}).get('status') in (None,'none') and not re.match(r'^\s*(?:[•·●]|\d+[、.)）])',before):
                    source_style=src.find('w:pPr/w:pStyle',NS)
                    if source_style is not None and 'bullet' in source_style.get(q('w:val'),'').lower():edit(node,0,0,'• ','materialize_inherited_bullet')
                pp=style_paragraph(node,self.styles[style],preserve_emphasis=role not in ('heading','heading_retained','document_title','front_heading','figure_caption','table_caption'),source_styles=source_styles)
                if self.chapter_num_id and (role=='heading' or (role=='document_title' and self.master)):
                    number_properties(pp,self.chapter_num_id,(level if self.master else level-1) if role=='heading' else 0)
                elif self.template_style_numbers.get(style) not in (None,'0'):number_properties(pp,0)
                if role=='list_item' and r['block_id'] in manual_lists:
                    plan=manual_lists[r['block_id']];number_properties(pp,plan['num_id'],plan['level'])
                actual_auto=actual_numbering.get(sr.getroottree().getpath(src),{})
                if role=='list_item' and r['block_id'] not in manual_lists and actual_auto.get('status') in (None,'none') and not any(x.get('block_id')==r['block_id'] and 'retained' in x.get('kind','') for x in self.exceptions):
                    self.exceptions.append(dict(kind='list_native_conversion_retained',source_id=sid,block_id=r['block_id'],message='列表身份保留，编号前缀或分组尚未取得迁号授权，保留原文字'))
                if role!='heading' and actual_auto.get('status') not in (None,'none'):
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
                if role=='equation' and not node.findall('m:oMathPara',NS):
                    maths=node.findall('m:oMath',NS)
                    if not maths or ordinary_text(node).strip():raise ValueError('Unsafe display equation promotion')
                    wrapper=E.Element(q('m:oMathPara'));pr=E.SubElement(wrapper,q('m:oMathParaPr'));jc=E.SubElement(pr,q('m:jc'));jc.set(q('m:val'),'center')
                    index=node.index(maths[0])
                    for math in maths:node.remove(math);wrapper.append(math)
                    node.insert(index,wrapper);changes.append({'operation':'promote_standalone_math','count':len(maths)})
            if r.get('figure_layout',{}).get('kind')=='inline-row':
                node=picture_row_table(node,r['figure_layout']['columns'])
                style_figure_table(node,self.styles,self.width,source_styles)
            # Names and IDs are session-independent package identifiers here.
            for n in node.iter():
                if n.tag in (q('w:bookmarkStart'),q('w:bookmarkEnd')):
                    old=n.get(q('w:id'))
                    if old not in bmids:raise ValueError('Unmatched source bookmark')
                    n.set(q('w:id'),bmids[old])
                    if n.tag==q('w:bookmarkStart'):n.set(q('w:name'),bmnames[n.get(q('w:name'))])
                if n.tag==q('w:hyperlink') and n.get(q('w:anchor')) in bmnames:n.set(q('w:anchor'),bmnames[n.get(q('w:anchor'))])
                for key,value in list(n.attrib.items()):
                    if not key.startswith('{'+NS['r']+'}'):continue
                    if value not in relmap:
                        relation=rels[value];kind=relation.get('Type','').rsplit('/',1)[-1]
                        if kind=='hyperlink':newrel=self.add_rel(relation)
                        elif kind=='image' and relation.get('TargetMode')!='External':
                            srcpart=relation.get('Target');srcpart=srcpart.lstrip('/') if srcpart.startswith('/') else posixpath.normpath(posixpath.join('word',srcpart))
                            data=sp[srcpart];digest=hashlib.sha256(data).hexdigest();ext=Path(srcpart).suffix.lower();dest='word/media/batch_'+digest[:24]+ext;self.parts[dest]=data
                            newrel=self.add_rel(relation);newrel.set('Target',posixpath.relpath(dest,'word'))
                            if not any(t.get('Extension')==ext[1:] for t in self.types):
                                typ=next((t for t in s_types if t.get('Extension')==ext[1:]),None)
                                if typ is None:raise ValueError('Image content type missing')
                                self.types.append(deepcopy(typ))
                        else:raise ValueError('Unsupported relationship '+kind)
                        relmap[value]=newrel.get('Id')
                    n.set(key,relmap[value])
            for color in node.xpath('.//m:oMath//w:color',namespaces=NS):
                if dict(color.attrib)!={q('w:val'):'000000'}:
                    self.math_colors.append({'source_id':sid,'source_index':idx,'before':dict(color.attrib),'after':'000000'});color.attrib.clear();color.set(q('w:val'),'000000')
            for inline in node.findall('.//wp:inline',NS):
                self.docpr+=1;dp=inline.find(q('wp:docPr'))
                if dp is not None:dp.set('id',str(self.docpr))
                ex=inline.find(q('wp:extent'));cx=int(ex.get('cx'));cy=int(ex.get('cy'));limit=self.width
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
            self.records.append({'source_id':sid,'source_index':idx,'block_id':r['block_id'],'output_index':oi,'role':role,'level':r.get('level'),'decision_status':r['status'],'relates_to':r.get('relates_to'),'source_text':before,'output_text':ordinary_text(node),'changes':changes,'section_boundary':r.get('section_boundary'),'number_group_id':r.get('number_group_id'),'figure_layout':r.get('figure_layout'),'subcaptions':r.get('subcaptions',False)})
        self.sources.append({'source_id':sid,'path':str(source.resolve()),'sha256':sha(source),'candidate':str(Path(candidate).resolve()),'candidate_sha256':sha(candidate),'source_number':source_number,'bookmark_id_map':bmids,'bookmark_name_map':bmnames,'relationship_map':relmap,'caption_maps':caption_maps,'retained_caption_roles':sorted(held_caption_roles),'heading_reference_map':heading_refs,'caption_field_targets':field_targets,'generated_group_title':group_title})
    def add_gap(self,source,sid,reason,source_number=1,source_title=None):
        label=source_title or Path(source).name
        if self.master and self.chapter_num_id:
            p=elt('w:p');run=elt('w:r');t=elt('w:t');t.text=label;run.append(t);p.append(run)
            pp=style_paragraph(p,self.styles['Heading 1'],False);number_properties(pp,self.chapter_num_id,0)
            index=len(self.body)-1;self.body.insert(index,p)
            self.extra.append(dict(kind='source_gap_heading',source_id=sid,source_number=source_number,output_index=index,text=label))
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
        report['status']='DRAFT_WITH_UNRESOLVED_STRUCTURE' if any(x['kind']=='unresolved_structure_preserved' for x in self.exceptions) else 'GENERATED'
        report['preserve_only_parts_changed']=[n for n in report['preserve_only_parts_changed'] if n!='word/numbering.xml' or not self.numbering_changed]
        path.with_suffix('.assembly.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf8');return report
