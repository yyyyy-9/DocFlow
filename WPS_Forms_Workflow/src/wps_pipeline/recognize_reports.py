"""Report roles over source facts, with conflicts queued for bounded review."""
import re
from copy import deepcopy
from .recognize_structure import recognize
from .numbered_candidates import rebuild_structure,authorize,list_introduction,VERSION

ABSTRACT={'摘要','执行摘要','abstract'}
REFERENCES={'参考文献','参考资料','参考记录','references','bibliography'}

META=re.compile(r'^(?:部门|作者|编制人?|编写人?|审核人?|批准人?|审定人?|编制单位|编写单位|单位|项目组|报告编号|文件编号|编制日期|完成日期|日期|时间)(?:\s*[:：].*|\s+\S.*)?$')
INSTITUTION=re.compile(r'^(?=.{2,60}$)[^。；;！？!?：:0-9]*(?:研究院|研究所|实验室|技术中心|研发中心|事业部|部门|有限公司)$')
DATE=re.compile(r'^(?:\d{4}\s*年\s*(?:\d{1,2}\s*月\s*(?:\d{1,2}\s*日)?)?|\d{4}[-/.]\d{1,2}(?:[-/.]\d{1,2})?)$')
TOC_NAMES={'目录','目次','contents','tableofcontents','图目录','表目录','插图目录','表格目录','listoffigures','listoftables'}


def title_value(text):
    value=text.strip()
    if value.startswith('【') and value.endswith('】'):value=value[1:-1].strip()
    return value


def toc_entry(p):
    text=p['text'].strip();style=p.get('style_name') or ''
    match=re.match(r'^(.*?)(?:\t+|[.．·…]{2,}| {2,})\s*([0-9]+|[ivxlcdmIVXLCDM]+|[一二三四五六七八九十]+)\s*$',text)
    styled=bool(re.match(r'^(?:toc|目录|目次)\s*[1-9]',style,re.I))
    if not match and styled:match=re.match(r'^(.*?)\s+([0-9]+|[ivxlcdmIVXLCDM]+)\s*$',text)
    if not match and re.match(r'^(?:\d+(?:\.\d+)*|[一二三四五六七八九十百]+[、.．]|第[一二三四五六七八九十百0-9]+[章节篇])',text):
        match=re.match(r'^(.*?)\s+([0-9]+|[ivxlcdmIVXLCDM]+)\s*$',text)
    if match:content=match[1].strip();page=match[2]
    elif styled and text:content=text;page=None
    else:return None
    prefix=re.match(r'^(\d+(?:\.\d+){0,5})[.、]?\s*(?=\S)',content)
    chinese=re.match(r'^(?:第[一二三四五六七八九十百0-9]+[章节篇]|[一二三四五六七八九十百]+[、.．)）])\s*',content)
    number=prefix[1] if prefix else None
    title=content[prefix.end():] if prefix else content[chinese.end():] if chinese else content
    level=len(number.split('.')) if number else 1 if chinese else None
    if styled:
        raw=re.search(r'[1-9]',style);level=int(raw[0]) if raw else level
    if not title or len(title)>200:return None
    return dict(title=title,old_number=number,level_hint=level,page=page,source_text=text,
                field_anchor=p.get('hyperlinks',[]),source_path=p['path'])


def frontmatter_plan(scan):
    """Select source scope before heading inference, using observable facts.

    Metadata matching applies only to the leading front matter. Directory
    entries are corroborating evidence; a TOC instruction alone cannot erase
    a span containing body prose. No page count or template/style ID is fixed.
    """
    facts=scan.get('body_layout_facts',[]);excluded={};entries=[];warnings=[];toc_blocks=set()
    byid={b['block_id']:b for b in scan['blocks']}
    def text(f):return '\n'.join(p['text'].strip() for p in f['paragraphs'] if p['text'].strip())
    def normal(s):return re.sub(r'\s+','',title_value(s)).casefold()
    def body_name(s):
        return normal(re.sub(r'^(?:\d+(?:\.\d+){0,5}[.、]?|第[一二三四五六七八九十百0-9]+[章节篇]|[一二三四五六七八九十百]+[、.．)）])\s*','',s.strip()))
    def exclude(f,kind,evidence):
        excluded[f['block_id']]=dict(block_id=f['block_id'],source_index=f['body_child_index'],
            source_sha256=scan['source_sha256'],xml_sha256=f['xml_sha256'],kind=kind,evidence=evidence)
    native_starts={s['start'] for s in scan.get('document_field_spans',[]) if re.match(r'\s*TOC\b',s['instruction'],re.I)}
    for i,f in enumerate(facts):
        if f['block_id'] in toc_blocks:continue
        paras=f['paragraphs'];header=any(normal(p['text']) in TOC_NAMES for p in paras)
        gallery=any(normal(g) in TOC_NAMES or 'tableofcontents' in normal(g) for g in f['galleries'])
        styled=any(re.match(r'^(?:toc|目录|目次)\s*[1-9]',p.get('style_name') or '',re.I) for p in paras)
        if not (header or gallery or styled or f['body_child_index'] in native_starts):continue
        group=[];found=[]
        for later in facts[i:]:
            if later['drawing_count'] or later['math_count']:break
            local=[];valid=True
            for p in later['paragraphs']:
                value=p['text'].strip()
                if not value or normal(value) in TOC_NAMES:continue
                entry=toc_entry(p)
                if entry:local.append(dict(entry,block_id=later['block_id']))
                elif re.match(r'^(?:错误[!！：:]|Error[!:])',value,re.I):continue
                else:valid=False;break
            if not valid:break
            if later['tag'] not in ('p','sdt','tbl') and not local:break
            group.append(later);found.extend(local)
        following=[p['text'] for g in facts[i+len(group):] for p in g['paragraphs']]
        corroborated=sum(any(body_name(v)==normal(e['title']) for v in following) for e in found)
        if len(found)>=2 or (len(found)==1 and (corroborated or gallery or styled)):
            for g in group:exclude(g,'source_toc',['directory_entry_cluster','body_matches:'+str(corroborated)]);toc_blocks.add(g['block_id'])
            entries.extend(found)
        else:warnings.append(dict(kind='source_toc_boundary_pending',block_id=f['block_id'],message='疑似目录边界或条目证据不足，首稿保留原文'))
    # Institutions are only cover evidence within a leading region containing
    # explicit metadata. A matching word anywhere in the report is not enough.
    prefix=[]
    for f in facts[:80]:
        meaningful=[p for p in f['paragraphs'] if p['text'].strip()]
        if any((p.get('outline') not in (None,'9') and (p.get('style_name') or '').casefold()!='title') or re.match(r'^(?:\d+[、.．)）\s]|[一二三四五六七八九十百]+[、.．)）]|第[一二三四五六七八九十百0-9]+[章节篇])',p['text'].strip()) for p in meaningful) and f['block_id'] not in excluded:break
        prefix.append(f)
    cover_anchor=any(META.fullmatch(p['text'].strip()) or DATE.fullmatch(p['text'].strip()) for f in prefix for p in f['paragraphs'])
    page=scan.get('first_page_facts') or {}
    page_cover=bool(cover_anchor and page.get('status')=='PASS' and page.get('source_sha256')==scan['source_sha256'] and page.get('source_unchanged'))
    title=None;leading=True;previous_label=False
    for ordinal,f in enumerate(facts):
        if f['block_id'] in excluded:continue
        value=text(f);lines=[p for p in f['paragraphs'] if p['text'].strip()]
        if not value:
            cover_context=any(META.fullmatch(p['text'].strip()) for g in facts[:ordinal+5] for p in g['paragraphs'])
            if leading and f['drawing_count'] and not f['math_count'] and cover_context:exclude(f,'source_cover',['leading_cover_graphic'])
            elif leading and f['page_break']:exclude(f,'source_cover',['leading_empty_page_boundary'])
            elif leading and f['tag']=='p' and not f['drawing_count'] and not f['math_count'] and (cover_anchor or toc_blocks):exclude(f,'source_cover',['leading_empty_cover_paragraph'])
            continue
        if not leading:break
        metadata=[p for p in lines if META.fullmatch(p['text'].strip()) or DATE.fullmatch(p['text'].strip())]
        if metadata and len(metadata)==len(lines):
            exclude(f,'source_cover',['leading_cover_metadata']);previous_label=bool(re.fullmatch(r'.*[:：]\s*',value));continue
        if cover_anchor and f['tag']=='p' and len(lines)==1 and INSTITUTION.fullmatch(value) and not f['drawing_count'] and not f['math_count'] and lines[0].get('outline') in (None,'9'):
            exclude(f,'source_cover',['leading_cover_institution_with_metadata']+(['wps_first_page_corroborates_cover'] if page_cover else []));previous_label=False;continue
        if f['tag'] in ('tbl','sdt') and len(metadata)>=2 and not f['math_count']:
            choices=[p for p in lines if p not in metadata and not re.search(r'[。；;！？!?]',p['text']) and 2<=len(p['text'].strip())<=200]
            if title is None and choices:
                p=choices[0];raw=p.get('plain_text',p['text']);title=dict(text=title_value(raw),raw_text=raw,block_id=f['block_id'],source_path=p['path'],origin='leading_cover_container')
            exclude(f,'source_cover',['leading_cover_metadata_container']);previous_label=False;continue
        if previous_label and f['tag']=='p' and len(value)<30 and not f['math_count'] and not f['drawing_count'] and not re.match(r'^\d',value) and not re.search(r'[。；;！？!?]',value):
            exclude(f,'source_cover',['leading_metadata_value']);previous_label=False;continue
        if title is None and f['tag']=='p' and len(lines)==1 and not f['drawing_count'] and not f['math_count']:
            p=lines[0];raw=p.get('plain_text',p['text']);candidate=raw.strip()
            protected=normal(candidate) in TOC_NAMES or candidate.casefold() in ABSTRACT|REFERENCES or list_introduction(candidate)
            numbered=re.match(r'^(?:\d|[一二三四五六七八九十]+、|第[一二三四五六七八九十百0-9]+[章节篇])',candidate)
            heading_style=p.get('outline') not in (None,'9') and (p.get('style_name') or '').casefold()!='title'
            if not protected and not numbered and not heading_style and 2<=len(candidate)<=200 and not re.search(r'[。；;！？!?]',candidate):
                title=dict(text=title_value(candidate),raw_text=raw,block_id=f['block_id'],source_path=p['path'],origin='leading_source_title');previous_label=False;continue
        leading=False
    # Bookmark endpoints are navigation metadata, not a reason to restore
    # author/institution cover lines. The assembler detaches unmatched endpoints
    # within the selected source scope and freezes dependent displayed links.
    hints={}
    for entry in entries:
        matches=[]
        for f in facts:
            if f['block_id'] in excluded or f['tag']!='p':continue
            b=byid[f['block_id']];value=b.get('display_text','').strip()
            if body_name(value)==normal(entry['title']):matches.append(f['block_id'])
        entry['body_matches']=matches
        if len(matches)==1:hints.setdefault(matches[0],[]).append(entry)
    return dict(contract='source-frontmatter-v1',source_sha256=scan['source_sha256'],title=title,
                exclusions=list(excluded.values()),toc_entries=entries,toc_hints=hints,warnings=warnings,
                first_page_evidence=dict(status=page.get('status','NOT_RUN'),used_as_corroboration=page_cover,whole_page_deletion=False))

def predict(scan,overrides=None):
    overrides=overrides or {}
    scope=frontmatter_plan(scan);excluded={e['block_id']:e for e in scope['exclusions']}
    work=deepcopy(scan);work['top_level_ids']=[bid for bid in scan['top_level_ids'] if bid not in excluded]
    if set(overrides)&set(excluded):raise ValueError('Decision addresses excluded source front matter')
    c=recognize(work,overrides=overrides)
    byid={b['block_id']:b for b in scan['blocks']};rows=c['blocks']
    def visible(i):return byid[rows[i]['block_id']].get('display_text','').strip()
    first=next((i for i in range(len(rows)) if visible(i)),None)
    anchors=[i for i,r in enumerate(rows) if (r.get('number_candidate') or {}).get('kind') in ('manual','automatic') and r['role']!='list_item']
    first_number=min(anchors,default=len(rows))
    for i,r in enumerate(rows):
        if r['block_id'] in overrides:continue
        b=byid[r['block_id']];text=visible(i);style=(b.get('style_name') or '').casefold()
        if b['kind']!='paragraph' or r.get('number_candidate') is not None:continue
        after=[visible(j) for j in range(i+1,min(i+4,len(rows))) if visible(j)]
        def role(name,why,level=None,status='auto_rule'):
            r.update(role=name,level=level,title=text if name in ('heading','front_heading','document_title') else None,status=status,evidence=[why],source_number_prefix='')
            r.pop('level_pending',None)
        outline=b.get('outline_level_raw')
        prose=len(text)>90 or (bool(re.search('[。；;！？!?]$',text)) and len(text)>30)
        if outline is not None and 0<=outline<5 and prose:
            # A style is evidence, not semantic truth. Genuine long headings
            # remain reviewable rather than being forced into body text.
            role('heading','effective_outline_conflicts_with_prose',outline+1,'review')
            r.update(possible_roles=['heading','body'],possible_levels=list(range(1,6)))
        elif text.casefold() in ABSTRACT and i<first_number and i<8:
            role('front_heading','opening_abstract_label_with_document_position')
        elif text.casefold() in REFERENCES and any(re.match(r'^\[\d+\]',x) for x in after):
            role('front_heading','reference_heading_before_bibliographic_items')
        elif style=='title':role('document_title','explicit_title_style')
        elif i==first and i<first_number and len(text)<=90 and not re.match(r'^\d',text) and not re.search('[。；;！？!?]',text) and not list_introduction(text) and any(x.casefold() in ABSTRACT for x in after):
            role('document_title','opening_title_before_abstract')
        elif '\t' in text and r['role'] not in ('table','figure','equation'):
            role('table_text','tab_delimited_visible_cells')
        elif 'list' in style and r['role']=='body':role('list_item','source_list_style')
        if r['role'] in ('figure_caption','table_caption','equation_description') and r['status']=='review' and not r.get('association_candidates'):
            r['status']='retained_unpaired';r['evidence'].append('preserve_caption_without_asserting_object_pair')
    for i,r in enumerate(rows):
        if r.get('status')!='auto_rule' or r.get('role')!='list_item':continue
        numbers=(r.get('number_candidate') or {}).get('numbers') or []
        outline=byid[r['block_id']].get('outline_level_raw')
        parent=next((p for p in reversed(rows[:i]) if (p.get('number_candidate') or {}).get('numbers')==numbers[:-1]),None) if len(numbers)>1 else None
        if (outline is not None and 0<=outline<5) or (parent is not None and parent.get('role')=='heading'):
            r.update(status='review',possible_roles=['heading','list_item','body'],possible_levels=list(range(1,6)))
            r['evidence'].append('automatic_list_conflicts_with_outline_or_heading_parent')
    title=scope.get('title')
    if title and title['block_id'] not in excluded and title['block_id'] not in overrides:
        row=next((r for r in rows if r['block_id']==title['block_id']),None)
        if row is not None:
            row.update(role='document_title',level=None,title=title['text'],status='auto_rule')
            row.pop('level_pending',None);row['evidence'].append('leading_source_title_preserved')
    from .numbered_candidates import numeric_expression
    for i,r in enumerate(rows):
        hints=scope['toc_hints'].get(r['block_id'],[])
        if not hints:continue
        r['toc_evidence']=hints;r['evidence'].append('source_directory_text_matched_body')
        levels={h['level_hint'] for h in hints if h['level_hint'] in range(1,6)}
        if r['block_id'] in overrides or r['role'] not in ('body','heading') or len(levels)!=1:continue
        level=next(iter(levels));b=byid[r['block_id']];text=b.get('plain_text','').strip()
        if r['role']=='heading':
            if r.get('level')!=level:
                r.update(status='review',possible_roles=['heading','body'],possible_levels=sorted({level,r.get('level')} - {None}))
                r['evidence'].append('directory_level_conflicts_with_body')
        elif not numeric_expression(text) and not re.search(r'[。；;！？!?]',text) and len(text)<=200:
            following=next((byid[x['block_id']].get('plain_text','').strip() for x in rows[i+1:] if byid[x['block_id']].get('plain_text','').strip()),'')
            if len(following)>40 and not (r.get('number_candidate') or {}).get('conflict'):
                r.update(role='heading',level=level,title=hints[0]['title'],status='auto_rule')
                r.pop('level_pending',None);r['evidence'].append('directory_and_standalone_body_context_agree')
    from .figure_layout import apply as apply_figure_layouts
    apply_figure_layouts(work,c,overrides)
    for bid,e in excluded.items():
        rows.append(dict(block_id=bid,source_locator=byid[bid]['source_locator'],role=e['kind'],level=None,title=None,
                         parent_id=None,relates_to=None,status='excluded_source_frontmatter',evidence=e['evidence']))
    rows.sort(key=lambda r:r['source_locator']['body_child_index'])
    c['source_frontmatter']=scope
    rebuild_structure(rows);authorize(rows,scan)
    c.update(recognizer_version=VERSION,review=[dict(block_id=r['block_id'],text=byid[r['block_id']].get('display_text'),role=r['role'],level=r['level'],evidence=r['evidence']) for r in rows if r['status']=='review'],confirmed_ir=False)
    c['status']='REVIEW_REQUIRED' if c['review'] else 'RULE_CANDIDATES_READY'
    return c
