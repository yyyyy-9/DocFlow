"""Three fixed forms over the existing local document pipeline.

The Agent is an optional caller. This module has no model API or SDK.
"""
from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED
from copy import deepcopy
from uuid import uuid4
import re,json,shutil,hashlib,time
from lxml import etree as E
from .job_support import ROOT,read,write,sha,stamp,contained,fingerprints
from .form_xlsx import write_form,read_form,col
from .docx_formatting import NS,q,elt,pack,ordinary_text
from .forms_organization import tree_plan,prepare_template

SCHEMA='wps-three-forms-1'
FORMS_VERSION='0.1.6'
POLICIES=('人工核查','接受程序建议','Agent辅助核查')
CH=('节点ID','来源/节点','程序建议标题','程序建议处理','处理决定','目标归类','同级顺序','展示标题','备注')
FH=('反馈ID','文档ID / 稳定定位','原文片段','修改类型','填写值','补充说明','作用范围','本条决定','处理结果')
IH=('输入项','填写值','说明')


def binding(state,kind):
    return dict(schema=SCHEMA,run_id=state['run_id'],inventory_sha256=state['inventory_sha256'],kind=kind)


def new_input(path,root='',template='',out='',title='工程报告汇总',policy='人工核查',pages='off',doc_conversion='wps'):
    if policy not in POLICIES:raise ValueError('核查方式无效')
    if pages not in ('off','wps') or doc_conversion not in ('off','wps'):raise ValueError('WPS选项无效')
    rows=[['接入根目录',root,'唯一必填；递归读取 .docx/.doc，原件不改'],['目标模板',template,'留空使用包内模板'],['任务输出目录',out,'必须是新目录；也可命令行 --run 指定'],['总稿名称',title,'可留空'],['核查方式',policy,'人工核查 / 接受程序建议 / Agent辅助核查'],['启用WPS首页检查','启用' if pages=='wps' else '关闭','启用：本机WPS只读副本取物理首页，作为封面/拆分的补充事实；失败继续出稿'],['DOC转DOCX预处理','启用' if doc_conversion=='wps' else '关闭','启用：遇.doc由本机WPS另存DOCX；保留二进制原件及转换回执。失败列缺口，其余继续']]
    return write_form(path,'输入',dict(schema=SCHEMA),rows,IH,{2},'01 输入','填写浅黄色单元格；保存后执行 scan。首页检查和DOC转换分别选择，无需模型服务。',widths=[27,78,76],choices={'B13':list(POLICIES),'B14':['关闭','启用'],'B15':['启用','关闭']},global_label='使用说明',global_value='根目录必填，其余可默认')


def wps_mode(value,default='off'):
    options={'':'off' if default=='off' else 'wps','off':'off','wps':'wps','关闭':'off','启用':'wps'}
    if value not in options:raise ValueError('WPS选项须为启用/关闭（兼容wps/off）')
    return options[value]


def load(run):
    run=Path(run).resolve()
    from .maintenance import assert_open
    assert_open(run)
    s=read(run/'forms_state.json')
    if s.get('schema')!=SCHEMA:raise ValueError('不支持的表单任务版本')
    if sha(run/'inventory.json')!=s['inventory_sha256']:raise ValueError('冻结清单改变，请保留现场并新建任务')
    if fingerprints()!=s['runtime_sha256']:raise ValueError('运行代码已改变；不得用新代码改写旧任务，请显式迁移或从冻结原件新建')
    inv=read(run/'inventory.json')
    for source in inv['sources']:
        if sha(contained(run,source['frozen']))!=source['sha256']:raise ValueError('冻结原件改变: '+source['id'])
        original=source.get('preprocessing_original')
        if original and sha(contained(run,original['file']))!=original['sha256']:raise ValueError('预处理原件改变: '+source['id'])
    if s.get('preprocessing_sha256') and sha(run/'preprocessing.json')!=s['preprocessing_sha256']:raise ValueError('预处理回执改变')
    if sha(run/'source_scan/inputs/template.docx')!=s['frozen_template_sha256']:raise ValueError('冻结模板改变')
    for source in inv['sources']:
        for name in source['evidence']:
            f=run/'source_scan/sources'/source['id']/name
            if sha(f)!=source['evidence'][name]:raise ValueError('扫描依据改变: '+source['id']+'/'+name)
    return run,s,inv


def save(run,s):write(Path(run)/'forms_state.json',s)


def enumerate_sources(root,run,template):
    root=Path(root).resolve();run=Path(run).resolve();template=Path(template).resolve()
    if not root.is_dir():raise ValueError('接入根目录不存在')
    if root==run or root.is_relative_to(run):raise ValueError('输出目录不能包含输入根目录')
    found=[];ignored=[];folders=[];seen=set()
    def walk(folder):
        for p in sorted(folder.iterdir(),key=lambda p:(p.name.casefold(),p.name)):
            rel=p.relative_to(root).as_posix()
            if p.is_symlink() or (getattr(p.stat(),'st_file_attributes',0)&0x400):
                ignored.append(dict(path=rel,reason='symbolic_link_or_reparse_point_not_followed'));continue
            if p.resolve()==run or p.resolve().is_relative_to(run):
                ignored.append(dict(path=rel,reason='current_output_tree'));continue
            if p.is_dir():
                if (p/'forms_state.json').exists() or (p/'state.json').exists():ignored.append(dict(path=rel,reason='existing_run_tree'));continue
                folders.append(rel);walk(p);continue
            if p.name.startswith('~$'):ignored.append(dict(path=rel,reason='office_lock_file'));continue
            if p.resolve().is_relative_to(ROOT/'templates') or p.resolve().is_relative_to(ROOT/'outputs'):
                ignored.append(dict(path=rel,reason='runtime_template_or_output'));continue
            if p.resolve()==template:ignored.append(dict(path=rel,reason='selected_template'));continue
            if p.suffix.lower() not in ('.docx','.doc'):ignored.append(dict(path=rel,reason='unsupported_extension'));continue
            stat=p.stat();identity=(stat.st_dev,stat.st_ino)
            if identity in seen:ignored.append(dict(path=rel,reason='same_file_alias'));continue
            seen.add(identity);found.append((p,rel))
    walk(root)
    if not found:raise ValueError('根目录没有可接入的 .docx/.doc；PDF/WPS等格式不静默当作已处理')
    return found,folders,ignored


def norm(text):
    text=re.sub(r'^\s*(?:\d+(?:[.．]\d+)*[、.．)）]?|[一二三四五六七八九十]+、|[（(]\d+[）)])\s*','',text)
    return re.sub(r'[\s:：。；;、.．]','',text).casefold()


def propose_units(source,scan,candidate,page):
    """Conservative split: opening list + unique confirmed H1 sections.

    A report with a cover/TOC remains atomic. Any unresolved boundary, native
    reference, cross-unit list instance or bookmark keeps it whole.
    """
    front=candidate.get('source_frontmatter',{})
    base_title=(front.get('title') or {}).get('text') or Path(source['name']).stem
    reason='没有可靠首页工作点，整篇保留';rows=candidate['blocks']
    if front.get('exclusions'):return base_title,[], '标准封面/目录报告作为整体'
    first=page.get('text','') if page.get('status')=='PASS' else ''
    method=page.get('method','NOT_AVAILABLE')
    if not first:
        # An explicit page break is a declared opening segment, not evidence
        # of the exact physical first page. Expose that distinction to users.
        with ZipFile(source['path']) as z:body=E.fromstring(z.read('word/document.xml')).find('w:body',NS)
        opening=[];boundary=None
        for i,node in enumerate(body):
            if node.xpath('.//w:br[@w:type="page"]|./w:pPr/w:pageBreakBefore[not(@w:val="0")]',namespaces=NS):boundary=i;break
            opening.append(ordinary_text(node))
        if boundary is None:return base_title,[],reason
        first='\n'.join(opening);method='EXPLICIT_OPENING_SEGMENT_NOT_PHYSICAL_PAGE'
    points=[]
    for line in re.split(r'[\r\n\x07\x0b]+',first):
        m=re.match(r'^\s*(?:\d+[、.)）]|[一二三四五六七八九十]+、|[（(]\d+[）)])\s*(\S.{1,55}?)\s*$',line)
        if m:points.append(norm(m[1]))
    headings=[r for r in rows if r['role']=='heading' and r.get('level')==1 and r['status']!='review' and r.get('number_edit',{}).get('allowed')]
    matches=[r for r in headings if norm(r.get('title') or r.get('text') or '') in points]
    if len(matches)<2 or len(matches)!=len(set(norm(r.get('title') or r.get('text') or '') for r in matches)):
        return base_title,[], '首页工作点与后文章节不能唯一对应，整篇保留'
    with ZipFile(source['path']) as z:body=E.fromstring(z.read('word/document.xml')).find('w:body',NS)
    if body.xpath('.//w:bookmarkStart|.//w:hyperlink[@w:anchor]|.//w:instrText[contains(.,"REF")]|.//w:fldSimple[contains(@w:instr,"REF")]',namespaces=NS):
        return base_title,[], '含书签或内部引用；本版保守地整篇保留，避免拆断关系'
    if re.search(r'(?:见|第)\s*\d+(?:\.\d+)*\s*(?:章|节)',ordinary_text(body)):
        return base_title,[], '含章节文字引用，整篇保留'
    starts=[r['source_locator']['body_child_index'] for r in matches]
    if starts!=sorted(starts) or any(r.get('section_boundary',{}).get('status')=='pending' for r in matches):
        return base_title,[], '章节边界未定，整篇保留'
    # Do not skip unmatched H1 sections or introductory text.
    bounds=[0]+[r['source_locator']['body_child_index'] for r in headings]
    bounds=sorted(set(bounds));end=len(body)-int(body[-1].tag==q('w:sectPr'))
    byindex={r['source_locator']['body_child_index']:r for r in rows};units=[];owners={}
    scanby={r['block_id']:r for r in scan['blocks']}
    for n,start in enumerate(bounds):
        stop=bounds[n+1] if n+1<len(bounds) else end
        root=byindex.get(start);is_heading=root and root['role']=='heading' and root.get('level')==1
        title=(root.get('title') or ordinary_text(body[start])) if is_heading else base_title+' · 项目介绍与工作清单'
        if start==stop:continue
        unit=dict(id=source['id']+'_U'+str(n+1).zfill(3),parent=source['id'],start=start,end=stop,
                  root_block_id=root['block_id'] if is_heading else None,title=title,method=method)
        units.append(unit)
        for row in rows:
            idx=row['source_locator']['body_child_index']
            if start<=idx<stop:
                b=scanby[row['block_id']];auto=b.get('numbering_effective') or {}
                if row['role']!='heading' and auto.get('num_id') not in (None,'0'):
                    owners.setdefault(auto['num_id'],set()).add(unit['id'])
    if any(len(v)>1 for v in owners.values()):return base_title,[], '自动列表实例跨拆分边界，整篇保留'
    return base_title,units,'首页工作点与后文明确章节对应；介绍、清单及未匹配章节也保留。依据：'+method


def scan_input(input_path,run=None,pages=None,doc_conversion=None):
    cells,meta=read_form(input_path,'输入')
    if meta.get('schema')!=SCHEMA:raise ValueError('输入表单版本不支持')
    root=cells.get('B9','').strip()
    if not root:raise ValueError('输入表接入根目录为空（B9）')
    policy=cells.get('B13','').strip() or '人工核查'
    if policy not in POLICIES:raise ValueError('核查方式必须为预设三选一')
    target_value=run or cells.get('B11','').strip()
    target=(Path(target_value).expanduser().resolve() if target_value else
            ROOT/'outputs'/('首稿_'+time.strftime('%Y%m%d_%H%M%S')+'_'+uuid4().hex[:6]))
    if target.exists():raise FileExistsError('任务必须使用全新目录: '+str(target))
    template=Path(cells.get('B10') or ROOT/'templates/target_template.docx').resolve()
    found,folders,ignored=enumerate_sources(root,target,template)
    mode=wps_mode(pages or cells.get('B14',''))
    doc_mode=wps_mode(doc_conversion or cells.get('B15',''),'wps')
    from .workflow import start_job
    tic=time.perf_counter();target.mkdir(parents=True)
    shutil.copy2(input_path,target/'01_输入.xlsx')
    from .form_pages import probe,convert_doc
    prepared=[];preprocessing=[];originals={}
    for i,(path,relative) in enumerate(found,1):
        sid='S'+str(i).zfill(4);copy=target/'preprocessed/originals'/(sid+path.suffix.lower());copy.parent.mkdir(parents=True,exist_ok=True)
        shutil.copy2(path,copy);original=dict(file=copy.relative_to(target).as_posix(),sha256=sha(copy),relative=relative)
        originals[sid]=original;item=dict(id=sid,path=str(copy),name=path.name)
        row=dict(id=sid,original=original,kind='doc_to_docx' if path.suffix.lower()=='.doc' else 'docx_copy',status='NOT_NEEDED')
        if path.suffix.lower()=='.doc':
            dest=target/'preprocessed/converted'/(sid+'.docx')
            result=convert_doc(target/'preprocessed/wps'/sid,copy,dest) if doc_mode=='wps' else dict(status='DISABLED',reason='01表已关闭DOC转换')
            row.update(result)
            if result['status']=='PASS' and result.get('source_unchanged') and dest.is_file():item['path']=str(dest)
            else:item['preprocessing_error']='DOC转DOCX未完成：'+str(result.get('reason') or result.get('status'))
        item['preprocessing_original']=original;prepared.append(item);preprocessing.append(row)
    write(target/'preprocessing.json',dict(mode=doc_mode,documents=preprocessing,semantic_equivalence='NOT_PROVEN_FOR_LEGACY_CONVERSION',model_api_calls=0))
    pp=probe(target,[dict(id=s['id'],path=s['path']) for s in prepared if not s.get('preprocessing_error')],mode)
    write(target/'page_facts.json',pp);pby={p['id']:p for p in pp.get('documents',[])}
    for item in prepared:
        if pby.get(item['id'],{}).get('status')=='PASS':item['page_facts']=pby[item['id']]
    job=dict(template=str(template),title=cells.get('B12') or '工程报告汇总',make_master=True,sources=prepared)
    core=start_job(job,Path(root),target/'source_scan')
    sources=[];nodes=[]
    for src,(_,relative) in zip(core['sources'],found):
        folder=target/'source_scan/sources'/src['id'];c=read(folder/'candidate.json') if (folder/'candidate.json').is_file() else None
        facts=read(folder/'scan/scan.json') if (folder/'scan/scan.json').is_file() else None
        source=dict(src,path=str(target/'source_scan'/src['frozen']),frozen='source_scan/'+src['frozen'],relative=relative,preprocessing_original=originals[src['id']],
                    evidence={n:sha(folder/n) for n in ('candidate.json','scan/scan.json') if (folder/n).is_file()})
        title,units,reason=propose_units(source,facts,c,pby.get(src['id'],{})) if c and facts else (Path(src['name']).stem,[],'来源扫描失败，首稿显式标记整篇缺口；冻结原件保留')
        source.update(true_title=title,units=units,split_reason=reason)
        sources.append(source);group='/'.join(Path(relative).parts[:-1])
        nodes.append(dict(id=src['id'],parent=None,relative=relative,title=title,suggestion='按建议拆分' if units else '整篇保留',group=group,reason=reason,order=len(nodes)+1))
        for u in units:nodes.append(dict(id=u['id'],parent=src['id'],relative=relative+' / '+u['id'],title=u['title'],suggestion='内容单元',group=group,reason=u['method'],order=len(nodes)+1))
    inventory=dict(schema=SCHEMA,root=str(Path(root).resolve()),directories=folders,sources=sources,nodes=nodes,ignored=ignored,
                   preprocessing=preprocessing,rule='physical provenance and target organization are separate; failed conversions are explicit gaps')
    write(target/'inventory.json',inventory)
    state=dict(schema=SCHEMA,version=FORMS_VERSION,run_id=uuid4().hex,created=stamp(),status='WAITING_FOR_CLASSIFICATION',policy=policy,
               title=job['title'],inventory_sha256=sha(target/'inventory.json'),frozen_template_sha256=core['template']['sha256'],
               runtime_sha256=fingerprints(),page_status=pp['status'],preprocessing_sha256=sha(target/'preprocessing.json'),drafts=[],events=[dict(stage='scan',seconds=round(time.perf_counter()-tic,4))],
               usage=dict(program_model_calls=0,agent_tokens=None,agent_tokens_status='UNAVAILABLE'))
    save(target,state);export_classification(target,state,inventory)
    # Always hand over a first master before organization/semantic review.
    # Whole sources are the conservative first-draft organization. Suggestions
    # in the workbook are not treated as already accepted split decisions.
    state['classification_status']='PENDING';save(target,state)
    result=assemble(target,None,origin='program_draft')
    if policy=='接受程序建议':
        auto=target.with_name(target.name+'_分类稿');number=1
        while auto.exists():number+=1;auto=target.with_name(target.name+'_分类稿'+str(number))
        return classify(target,None,auto,origin='program_policy')
    return result


def export_classification(run,state,inv):
    rows=[[n['id'],n['relative'],n['title'],n['suggestion'],'接受建议',n['group'],n['order'],'',n['reason']] for n in inv['nodes']]
    if (run/'approved_plan.json').is_file():
        approved=read(run/'approved_plan.json');current={e['id']:e for e in approved['entries']};excluded={e['id'] for e in approved['excluded']}
        splitparents={e['source_id'] for e in approved['entries'] if e['kind']=='split'}
        for row in rows:
            nid=row[0]
            if nid in current:
                e=current[nid];row[4]='整篇保留' if e['kind']=='whole' else '接受建议';row[5]='/'.join(e['group']);row[6]=e['order'];row[7]=e['title']
            elif nid in splitparents:row[4]='按建议拆分'
            elif nid in excluded:row[4]='不纳入'
    return write_form(run/'02_扫描核查.xlsx','核查',binding(state,'核查'),rows,CH,set(range(5,10)),
        '02 扫描核查：来源与目标组织','只改不合适的行，B6 确认本表。整篇保留会抑制子单元；按建议拆分保留全部非排除单元。归类用 / 分层。',
        widths=[20,58,40,20,20,32,14,40,65],global_label='核查提交',choices={'B6':['未确认','确认本表'],'E9:E1048576':['接受建议','整篇保留','按建议拆分','不纳入']})


def parse_classification(run,s,inv,form,origin):
    if origin in ('program_policy','program_draft'):cells=None
    else:
        cells,b=read_form(form or run/'02_扫描核查.xlsx','核查')
        if b!=binding(s,'核查'):raise ValueError('核查表属于其他任务或冻结清单')
        if cells.get('B6')!='确认本表':raise ValueError('先核查并在 B6 选择“确认本表”')
    parsed={}
    for i,node in enumerate(inv['nodes'],9):
        if cells is not None:
            if [cells.get(c+str(i),'') for c in 'ABCD']!=[node['id'],node['relative'],node['title'],node['suggestion']]:raise ValueError('程序事实列被改动或行被删除/移动: '+node['id'])
            decision=cells.get('E'+str(i),'') or '接受建议';group=cells.get('F'+str(i),'');order=cells.get('G'+str(i),'');title=cells.get('H'+str(i),'').strip() or node['title']
        else:decision='接受建议';group=node['group'];order=node['order'];title=node['title']
        if decision not in ('接受建议','整篇保留','按建议拆分','不纳入'):raise ValueError('非法处理决定: '+node['id'])
        try:
            order=float(order)
            if not order.is_integer() or not 1<=order<=999999:raise ValueError()
            order=int(order)
        except Exception:raise ValueError('同级顺序须为 1..999999 的整数: '+node['id'])
        groups=[g.strip() for g in group.replace('\\','/').split('/') if g.strip()]
        if any(g in ('.','..') or len(g)>100 for g in groups) or len(groups)>7:raise ValueError('归类格式或层级无效: '+node['id'])
        if not title or len(title)>200:raise ValueError('展示标题须为 1..200 字')
        parsed[node['id']]=dict(id=node['id'],decision=decision,group=groups,order=order,title=title)
    if cells:
        last=8+len(inv['nodes'])
        if any(re.fullmatch(r'[A-I](\d+)',k) and int(re.search(r'\d+',k)[0])>last and str(v).strip() for k,v in cells.items()):raise ValueError('不允许追加未知来源节点；请从新根目录重新扫描')
    entries=[];excluded=[]
    for src in inv['sources']:
        p=parsed[src['id']];choice=p['decision']
        if origin=='program_draft':choice='整篇保留'
        elif choice=='接受建议':choice='按建议拆分' if src['units'] else '整篇保留'
        if choice=='不纳入':excluded.append(dict(id=src['id'],reason='explicit_classification_exclusion'));continue
        if choice=='按建议拆分':
            if not src['units']:raise ValueError('没有安全拆分建议，不能臆造边界: '+src['id'])
            for unit in src['units']:
                d=parsed[unit['id']]
                if d['decision']=='不纳入':excluded.append(dict(id=unit['id'],reason='explicit_unit_exclusion'));continue
                if d['decision']=='按建议拆分':raise ValueError('本版不递归猜测子单元边界')
                if d['group']==[g for g in Path(src['relative']).parts[:-1]]:d['group']=p['group']
                entries.append(dict(d,source_id=src['id'],unit=unit,kind='split'))
        else:entries.append(dict(p,source_id=src['id'],unit=None,kind='whole'))
    if not entries:raise ValueError('没有纳入任何来源；不创建空白假总稿')
    return entries,excluded


def derive_unit(run,source,entry,dest):
    """Local selected copy, with a replayable map back to frozen XML.

    Only already-authorized chapter prefixes/styles are rebased. This is
    deterministic organization, never recorded as a model/human judgment.
    """
    with ZipFile(run/source['frozen']) as z:
        parts={n:z.read(n) for n in z.namelist()}
        metadata={n:deepcopy(z.getinfo(n)) for n in z.namelist()}
    root=E.fromstring(parts['word/document.xml']);body=root.find('w:body',NS);original=list(body);u=entry['unit']
    candidate=read(run/'source_scan/sources'/source['id']/'candidate.json');rows={r['source_locator']['body_child_index']:r for r in candidate['blocks']}
    byname={s.find('w:name',NS).get(q('w:val')).casefold():s.get(q('w:styleId')) for s in E.fromstring(parts['word/styles.xml']).findall('w:style',NS) if s.find('w:name',NS) is not None}
    for n in list(body):body.remove(n)
    title=elt('w:p');pp=elt('w:pPr');pp.append(elt('w:pStyle',val=byname.get('title','Normal')));pp.append(elt('w:numPr'));pp[-1].append(elt('w:numId',val=0));title.append(pp)
    # Display names belong to the target organization, not the frozen source
    # unit. A rename must not invalidate decisions about unchanged body blocks.
    r=elt('w:r');t=elt('w:t');t.text=u['title'];r.append(t);title.append(r);body.append(title)
    mapping=[];counters=[0]*5
    for i in range(u['start'],u['end']):
        source_node=original[i];row=rows.get(i);edit=None
        if row and row['block_id']==u.get('root_block_id'):
            mapping.append(dict(source_index=i,unit_index=0,operation='root_to_title',authorization=row['number_edit']));continue
        node=deepcopy(source_node)
        if row and row['role']=='document_title':
            # The intrinsic introductory title remains source data. A target
            # display rename is applied separately by the organization plan.
            mapping.append(dict(source_index=i,unit_index=0,operation='source_title_to_title'));continue
        if row and row['role']=='heading' and u.get('root_block_id'):
            level=row['level']-1 if type(row.get('level')) is int else None
            proposal=list(counters)
            if type(level) is int and 1<=level<=4 and row.get('number_edit',{}).get('allowed'):
                proposal[level-1]+=1;proposal[level:]=[0]*(5-level)
                if not any(v==0 for v in proposal[:level]):
                    auth=row['number_edit'];oldtext=ordinary_text(node);lo,hi=auth['span'];new='.'.join(map(str,proposal[:level]))+' '
                    from .assemble_reports import text_edit
                    text_edit(node,lo,hi,new);counters=proposal
                    pp=node.find('w:pPr',NS)
                    if pp is None:pp=elt('w:pPr');node.insert(0,pp)
                    for n in list(pp):
                        if n.tag in (q('w:pStyle'),q('w:numPr'),q('w:outlineLvl')):pp.remove(n)
                    pp.insert(0,elt('w:pStyle',val=byname['heading '+str(level)]));num=elt('w:numPr');num.append(elt('w:numId',val=0));pp.append(num)
                    edit=dict(authorization=auth,new_prefix=new,level=level,source_text=oldtext)
        record=dict(source_index=i,unit_index=len(body),operation='rebase_heading' if edit else 'copy',edit=edit)
        if row and row['role']=='heading' and u.get('root_block_id') and not edit:record['retained_heading_reason']='子标题缺少安全重建条件，原文原号保留，交后续结构判断；未中止分类'
        mapping.append(record);body.append(node)
    body.append(deepcopy(original[-1]));parts['word/document.xml']=pack(root)
    with ZipFile(dest,'x',ZIP_DEFLATED) as z:
        # Preserve the frozen member metadata, including timestamps and order.
        # Rebuilding the same unit must keep its complete byte hash, so valid
        # structure decisions and local overrides survive later classification.
        for name,data in parts.items():z.writestr(metadata[name],data)
    return dict(id=entry['id'],source_id=source['id'],source=source['frozen'],source_sha256=source['sha256'],file=dest.relative_to(run).as_posix(),sha256=sha(dest),
                range=[u['start'],u['end']],title=u['title'],display_title=entry['title'],mapping=mapping)


def assemble(run,form=None,origin='user',selection=None):
    run,s,inv=load(run)
    if s['drafts']:raise ValueError('已有稿件；请通过反馈建立新目录，不覆盖首稿')
    if origin not in ('user','model','program_policy','program_draft'):raise ValueError('提交来源须为 user/model/program_policy/program_draft')
    entries,excluded=selection or parse_classification(run,s,inv,form,origin)
    if origin=='program_draft':
        for e in entries:
            if len(e['group'])>3:e['group']=[' / '.join(e['group'])]
    plan=tree_plan(entries);byid={e['id']:e for e in entries};sources={r['id']:r for r in inv['sources']}
    # Validate depth before any partial draft is created.
    for row in plan['sources']:
        e=byid[row['source_id']];cp=run/'source_scan/sources'/e['source_id']/'candidate.json';cand=read(cp) if cp.exists() else dict(blocks=[])
        level=max([r.get('level') or 0 for r in cand['blocks'] if r['role']=='heading' and r['status']!='review']+[0])
        if len(row['prefix'])+level>9:raise ValueError('目标超过 Word 九级标题能力；请减少归类层级，不会压平源标题')
    folder=run/'prepared';folder.mkdir();template=folder/'template.docx';tp=prepare_template(run/'source_scan/inputs/template.docx',template)
    derivations=[];inputs=[]
    for row in plan['sources']:
        e=byid[row['source_id']];src=sources[e['source_id']]
        if e['kind']=='whole':path=run/src['frozen']
        else:
            path=folder/(e['id']+'.docx');derivations.append(derive_unit(run,src,e,path))
        inputs.append(dict(id=e['id'],path=str(path),title=e['title'],name=e['title']+'.docx',
                           **{k:src[k] for k in ('page_facts','preprocessing_error') if k in src and e['kind']=='whole'}))
    approved=dict(schema=SCHEMA,run_id=s['run_id'],inventory_sha256=s['inventory_sha256'],origin=origin,entries=entries,excluded=excluded,
                  organization=plan,derivations=derivations,template_preparation=tp)
    write(run/'approved_plan.json',approved)
    from .verify_forms import verify_preparation
    checked=verify_preparation(run,inv,approved);write(run/'preparation_check.json',checked)
    if checked['status']!='PASS':raise ValueError('单元准备独立检查失败；保留现场，不装配')
    # Export a revision's next form only after its accepted plan is saved.
    # The first scan already has its deliberately unreviewed suggestion form.
    if not (run/'02_扫描核查.xlsx').exists():export_classification(run,s,inv)
    from .workflow import start_job,finish,status
    core_run=run/'draft_001';core=start_job(dict(template=str(template),title=s['title'],sources=inputs,make_master=True,organization=plan),run,core_run)
    # Classification changes may keep valid previous local judgments. They
    # are reused only for byte-identical frozen source units, never re-labelled
    # as new human/model decisions.
    if s.get('parent'):
        oldhome=Path(s['parent']);oldforms=read(oldhome/'forms_state.json');oldrun=oldhome/oldforms['drafts'][-1]['run'];oldstate=read(oldrun/'state.json')
        from .model_review import resolve
        from .workflow import save as core_save,refresh_preparation_status
        reusable={x['id']:x for x in oldstate['sources']};unreused=[]
        for src in core['sources']:
            oldsrc=reusable.get(src['id']);dp=oldrun/'sources'/src['id']/'decisions.json'
            if oldsrc and oldsrc['sha256']==src['sha256']:
                if dp.is_file():
                    ds=read(dp);folder=core_run/'sources'/src['id'];c=resolve(read(folder/'scan/scan.json'),ds)
                    write(folder/'decisions.json',ds);write(folder/'candidate.json',c)
                    core['stages']['recognize:'+src['id']].update(evidence_sha256=sha(folder/'candidate.json'),status='REVIEW_REAPPLIED')
            elif dp.is_file():unreused.append(src['id'])
        validids={x['id'] for x in core['sources'] if x['id'] in reusable and x['sha256']==reusable[x['id']]['sha256']}
        core['layout_overrides']=[x for x in oldstate.get('layout_overrides',[]) if x['source_id'] in validids]
        s['classification_unreused_decisions']=unreused;refresh_preparation_status(core_run,core);core_save(core_run,core)
    finished=finish(core_run,core,retain=True);receipt=status(core_run,finished);write(run/'draft_001.status.json',receipt)
    s.update(status=finished['status'],approved_plan_sha256=sha(run/'approved_plan.json'))
    s['drafts'].append(dict(run='draft_001',status=finished['status'],state_sha256=sha(core_run/'state.json')))
    s['events'].append(dict(stage='assemble',time=stamp(),origin=origin,selected_units=len(entries)))
    save(run,s);export_feedback(run,s,core_run,finished);write_summary(run,s,inv)
    return forms_status(run)


def classify(run,form,dest,origin='user'):
    run,s,inv=load(run);dest=Path(dest).resolve()
    if not s['drafts']:raise ValueError('扫描应已交出首稿；请先完成首稿')
    if dest.exists() or dest==run or dest.is_relative_to(run):raise ValueError('分类修订使用任务外全新目录；已有总稿保留')
    entries,excluded=parse_classification(run,s,inv,form,origin)
    # Check all selection and tree constraints before creating a revision.
    tree_plan(entries)
    dest.mkdir(parents=True)
    for name in ('inventory.json','01_输入.xlsx','page_facts.json'):shutil.copy2(run/name,dest/name)
    if (run/'preprocessing.json').exists():shutil.copy2(run/'preprocessing.json',dest/'preprocessing.json');shutil.copytree(run/'preprocessed',dest/'preprocessed')
    shutil.copytree(run/'source_scan',dest/'source_scan')
    shutil.copy2(form or run/'02_扫描核查.xlsx',dest/'submitted_classification.xlsx')
    new=deepcopy(s);new.update(run_id=uuid4().hex,created=stamp(),status='CLASSIFICATION_ACCEPTED',parent=str(run),drafts=[],classification_status='APPLIED',classification_origin=origin)
    save(dest,new)
    return assemble(dest,None,origin=origin,selection=(entries,excluded))


def export_feedback(run,s,core_run,core,path=None):
    rows=[];bindings=[];seen=set()
    history={x['id']:x for x in s.get('feedback_history',[])}
    closed={(x.get('document'),x.get('locator'),x.get('kind')) for x in history.values() if x['status'] in ('closed_in_checked_draft','closed_retained_in_checked_draft')}
    merged={}
    for original in core.get('finish_list',[]):
        issue=deepcopy(original);sid=issue.get('source_id');bid=issue.get('block_id');doc=issue.get('document') or 'master'
        kind='structure_pending' if issue.get('kind') in ('structure_pending','unresolved_structure_preserved') else issue.get('kind')
        key=('source_block',sid,bid,kind) if sid and bid else ('document',doc,sid,kind,issue.get('text',''))
        group=merged.setdefault(key,dict(issue,kind=kind,related_documents=[],related_kinds=[],messages=[]))
        if doc not in group['related_documents']:group['related_documents'].append(doc)
        if original.get('kind') not in group['related_kinds']:group['related_kinds'].append(original.get('kind'))
        message=issue.get('message') or original.get('kind')
        if message not in group['messages']:group['messages'].append(message)
        if doc=='master':group['document']='master'
    for issue in merged.values():
        doc=issue.get('document') or 'master';sid=issue.get('source_id');bid=issue.get('block_id');locator=sid+':'+bid if sid and bid else sid or ''
        key=(doc,locator,issue.get('kind'))
        if key in closed:continue
        if key in seen:continue
        seen.add(key);fid='F'+str(len(rows)+1).zfill(5)
        hit=None;d=next((d for d in core['deliverables'] if d['id']==doc),None)
        if d and locator:hit=next((x for x in read(core_run/d['locator_index']) if x['locator']==locator),None)
        rows.append([fid,doc+(' | '+locator if locator else ''),(hit or {}).get('text',issue.get('text',''))[:500],'','','；'.join(issue['messages']),'仅此处','暂缓','待处理'])
        source=next((x for x in core['sources'] if x['id']==sid),None)
        key_material='|'.join((source['sha256'] if source else '',doc,locator,issue.get('kind','')))
        stable_id='Q-'+hashlib.sha256(key_material.encode('utf8')).hexdigest()[:24]
        bindings.append(dict(id=fid,history_id=stable_id,document=doc,locator=locator,kind=issue.get('kind'),source_id=sid,block_id=bid,display=rows[-1][:3],related_documents=issue['related_documents'],related_kinds=issue['related_kinds']))
    # User-added unresolved requests must survive in the next editable table,
    # not only in an internal ledger. Carry the latest state per stable ID.
    for item in history.values():
        values=item.get('submitted_values')
        if not values or item['status'] in ('closed_in_checked_draft','closed_retained_in_checked_draft'):continue
        if item['status']=='deferred' and item.get('kind')!='user_feedback':continue
        fid='F'+str(len(rows)+1).zfill(5);row=[fid]+list(values[1:8])+['待处理'];row[7]='暂缓'
        row[5]=(row[5]+'；' if row[5] else '')+'上批：'+item['status']
        rows.append(row);bindings.append(dict(id=fid,history_id=item['id'],document=item.get('document','master'),locator=item.get('locator',''),kind=item.get('kind','user_feedback'),source_id=item.get('source_id'),block_id=item.get('block_id'),display=row[:3]))
    # Blank rows are explicitly part of the export; new feedback is located by
    # exact quote or a stable locator before any revision is created.
    for i in range(30):rows.append(['','','','','','','仅此处','暂缓',''])
    meta=dict(binding(s,'反馈'),core_run=str(core_run.resolve()),parent_state_sha256=sha(core_run/'state.json'),exports=bindings,
              outputs={d['id']:d['sha256'] for d in core['deliverables']})
    write(run/'feedback_export.json',meta)
    s['feedback_export_sha256']=sha(run/'feedback_export.json');save(run,s)
    return write_form(path or run/'03_反馈提交.xlsx','反馈',meta,rows,FH,set(range(2,9)),
        '03 反馈提交：批量定位，一次修订','B6 提交本表；默认仅此处。B列填写 文档ID | 来源ID:块ID，或只填文档ID并复制原文到C列。大范围/内容缺口保留为明确事项。',
        widths=[17,50,65,20,35,65,17,19,25],global_label='反馈提交',global_value='未提交',
        choices={'B6':['未提交','提交本表'],'D9:D1048576':['标题级别','改为正文','改为列表','题名关联','字号','对齐','首行缩进','段前','段后','行距','目标归类','单元顺序','内容缺口','其他'],'G9:G1048576':['仅此处','本单元','全文'],'H9:H1048576':['提交修改','保留原样','暂缓']})


def write_summary(run,s,inv):
    lines=['# 三表流程结果','',f'状态：{s["status"]}',f'来源文件：{len(inv["sources"])}；忽略/不支持条目：{len(inv["ignored"])}。',
           '','本次正文处理没有调用模型 API。Agent 可填写同一表单，不是由程序调用模型。',
           '','## 已出稿']
    for d in s['drafts']:lines.append(f'- [{d["run"]}]({d["run"]}/DELIVERY.md)：{d["status"]}')
    lines+=['','## 下一步','- 在 03_反馈提交.xlsx 集中填写；feedback 命令验证整批后，只创建一个新修订。',
            '- 暂缓/未定位/超出本地修订范围的事项保留，不冒充已修改。','- 不覆盖冻结原件、模板、旧稿；不将旧稿作为新来源。',
            '',f'首页读取：{s.get("page_status")}；WPS 成稿验收：NOT_RUN；截图：NOT_REQUESTED。',
            '精确 Agent token：UNAVAILABLE (null)，程序模型调用：0。']
    for item in inv.get('preprocessing',[]):
        if item['kind']=='doc_to_docx':lines.append('- DOC转换 '+item['id']+'：'+item['status']+'；原始二进制保留；语义等价未单独证明。'+str(item.get('reason','')))
    (run/'DELIVERY.md').write_text('\n'.join(lines)+'\n',encoding='utf8')


def forms_status(run):
    run,s,inv=load(run);core=None
    if s['drafts']:
        latest=s['drafts'][-1];cr=contained(run,latest['run']);state=read(cr/'state.json')
        from .workflow import status
        core=status(cr,state)
    return dict(status=s['status'],run=str(run),files=len(inv['sources']),suggested_split_sources=sum(bool(r['units']) for r in inv['sources']),
                ignored=len(inv['ignored']),forms=[str(run/n) for n in ('01_输入.xlsx','02_扫描核查.xlsx','03_反馈提交.xlsx') if (run/n).exists()],
                latest=core,usage=s['usage'],page_status=s['page_status'])


def submit_feedback(run,form,dest,origin='user'):
    """Validate the entire batch before one fresh revision and one finish."""
    run,s,inv=load(run);dest=Path(dest).resolve()
    if not s['drafts']:raise ValueError('请先生成首稿')
    if origin not in ('user','model'):raise ValueError('反馈来源须为 user/model')
    if dest.exists() or dest==run or dest.is_relative_to(run):raise ValueError('反馈输出必须是任务外的全新目录')
    if sha(run/'feedback_export.json')!=s['feedback_export_sha256']:raise ValueError('反馈导出绑定记录改变')
    expected=read(run/'feedback_export.json');cells,b=read_form(form or run/'03_反馈提交.xlsx','反馈')
    if b!=expected:raise ValueError('反馈表不是本稿导出的表单')
    if cells.get('B6')!='提交本表':raise ValueError('在 B6 选择“提交本表”后提交')
    parent=Path(expected['core_run']);old=read(parent/'state.json')
    if sha(parent/'state.json')!=expected['parent_state_sha256']:raise ValueError('反馈对应的旧任务状态已改变')
    from .workflow import status,revise,finish,save as core_save,refresh_preparation_status
    from .feedback import locate,validate_layout
    from .model_review import resolve,validate_relation
    if status(parent,old)['receipt_integrity']!='PASS':raise ValueError('父稿或回执改变；先保留手工修改，不自动覆盖')
    if sha(run/'approved_plan.json')!=s['approved_plan_sha256']:raise ValueError('已核查组织计划改变')
    approved=read(run/'approved_plan.json');exports={r['id']:r for r in expected['exports']}
    rows=sorted({int(re.search(r'\d+',k)[0]) for k in cells if re.fullmatch(r'[A-I]\d+',k) and int(re.search(r'\d+',k)[0])>=9})
    actions=[];ledger=[];seen=set();touched=set();prepared_decisions={};layout=[];entries=deepcopy(approved['entries'])
    for n in rows:
        raw=[cells.get(c+str(n),'') for c in 'ABCDEFGHI']
        v=[value.strip() for value in raw];fid=v[0] or 'NEW-'+s['run_id'][:8]+'-'+str(n)
        if v[0]:
            if fid not in exports or fid in seen:raise ValueError('未知或重复反馈ID: '+fid)
            seen.add(fid)
            # Source quotes can legitimately contain leading/trailing spaces.
            # Compare immutable facts as exported, normalize only user intent.
            if raw[:3]!=exports[fid]['display'] or raw[8]!='待处理':raise ValueError('反馈事实/结果列被改动: '+fid)
        if v[7] in ('','暂缓'):
            if v[0]:ledger.append(dict(id=exports[fid].get('history_id',fid),status='deferred',kind=exports[fid]['kind'],document=exports[fid]['document'],locator=exports[fid]['locator'],submitted_values=v))
            continue
        if v[7] not in ('提交修改','保留原样'):raise ValueError('本条决定非法: '+fid)
        item=dict(id=exports.get(fid,{}).get('history_id',fid),origin=origin,status='awaiting_generation',kind=exports.get(fid,{}).get('kind','user_feedback'),problem=v[5],wanted=v[3]+':'+v[4],submitted_values=v)
        if v[6] not in ('','仅此处'):
            ledger.append(dict(item,status='retained_scope_requires_explicit_plan',requested_scope=v[6]));continue
        doc,sep,locator=v[1].partition('|');doc=doc.strip() or 'master';locator=locator.strip()
        if v[7]=='保留原样':
            if fid not in exports:raise ValueError('空白新行不能凭空关闭程序事项')
            ledger.append(dict(item,document=doc,locator=locator,retain=True));continue
        if not v[3]:raise ValueError('提交修改时需要修改类型: '+fid)
        found=locate(parent,doc,quote=v[2] or None,locator=locator or None,problem=v[5] or v[3],wanted=v[4] or v[3],origin=origin)
        if found['status']!='LOCATED':ledger.append(dict(item,status='retained_'+found['status'].lower(),document=doc,locator=locator));continue
        hit=found['matches'][0];sid,bid=hit['source_id'],hit['block_id'];item.update(document=doc,locator=hit['locator'],source_id=sid,block_id=bid)
        layout_kind=v[3] in ('字号','对齐','首行缩进','段前','段后','行距')
        key=(doc if layout_kind else None,sid,bid,v[3])
        if key in touched:raise ValueError('同一位置同类反馈重复，请合并: '+hit['locator'])
        touched.add(key)
        if v[3] in ('目标归类','单元顺序'):
            entry=next((e for e in entries if e['id']==sid),None)
            if entry is None:raise ValueError('归类对象没有对应内容单元')
            if v[3]=='目标归类':
                group=[g.strip() for g in v[4].replace('\\','/').split('/') if g.strip()]
                if any(g in ('.','..') or len(g)>100 for g in group) or len(group)>7:raise ValueError('归类格式无效')
                entry['group']=group
            else:
                if not re.fullmatch(r'[1-9]\d{0,5}',v[4]):raise ValueError('单元顺序须为 1..999999 整数')
                entry['order']=int(v[4])
            actions.append(dict(item,kind='organization'));ledger.append(item);continue
        scan=read(parent/'sources'/sid/'scan/scan.json');blocks={b['block_id']:b for b in scan['blocks']}
        if blocks[bid]['kind']!='paragraph':
            ledger.append(dict(item,status='retained_nonparagraph_requires_local_plan'));continue
        if v[3] in ('标题级别','改为正文','改为列表','题名关联'):
            role={'标题级别':'heading','改为正文':'body','改为列表':'list_item'}.get(v[3]);level=None
            decision=dict(block_id=bid,role=role,level=None,origin=origin,decision_id=fid,evidence_ids=[bid],reason=v[5] or '表单局部反馈：'+v[3])
            if v[3]=='标题级别':
                if not re.fullmatch('[1-5]',v[4]):raise ValueError('标题级别填写来源级别1..5（总稿归类层级由程序叠加）')
                decision['level']=int(v[4])
            elif v[3]=='题名关联':
                label,sep,target=v[4].partition(':')
                if not sep or label not in ('图','表') or target not in scan['top_level_ids']:raise ValueError('题名关联填写 图:bXXXXXX 或 表:bXXXXXX；必须是同一来源实际对象')
                decision.update(role='figure_caption' if label=='图' else 'table_caption',relates_to=target)
                validate_relation(decision,blocks,set(scan['top_level_ids']))
            existing=prepared_decisions.setdefault(sid,read(parent/'sources'/sid/'decisions.json') if (parent/'sources'/sid/'decisions.json').exists() else [])
            if any(d.get('source_id')==sid and d.get('block_id')==bid for d in actions if d.get('kind')=='structure'):raise ValueError('同一文段提交互斥结构修改')
            existing[:]=[d for d in existing if d['block_id']!=bid]+[decision]
            actions.append(dict(item,kind='structure',decision=decision));ledger.append(item)
        elif v[3] in ('字号','对齐','首行缩进','段前','段后','行距'):
            keys={'字号':'font_pt','对齐':'alignment','首行缩进':'first_line_pt','段前':'space_before_pt','段后':'space_after_pt','行距':'line_multiple'}
            key=keys[v[3]]
            try:value={'左对齐':'left','居中':'center','右对齐':'right','两端对齐':'both'}.get(v[4],v[4]) if key=='alignment' else float(v[4])
            except Exception:raise ValueError('格式填写值无效: '+fid)
            settings=dict(kind='layout',**{key:value});validate_layout(settings)
            layout.append(dict(document=doc,source_id=sid,block_id=bid,feedback_id=fid,origin=origin,settings=settings,reason=v[5] or v[3]))
            actions.append(dict(item,kind='layout'));ledger.append(item)
        else:ledger.append(dict(item,status='retained_requires_development'))
    if set(exports)-seen:raise ValueError('程序反馈行被删除；要保留请使用暂缓/保留原样')
    # Resolve every affected source before creating any destination directory.
    for sid,decisions in prepared_decisions.items():resolve(read(parent/'sources'/sid/'scan/scan.json'),decisions)
    organization=tree_plan(entries)
    # Establish a fresh wrapper. The old run, output and Excel are immutable.
    dest.mkdir(parents=True)
    for name in ('inventory.json','01_输入.xlsx','page_facts.json'):shutil.copy2(run/name,dest/name)
    shutil.copytree(run/'source_scan',dest/'source_scan')
    if (run/'preprocessing.json').exists():shutil.copy2(run/'preprocessing.json',dest/'preprocessing.json');shutil.copytree(run/'preprocessed',dest/'preprocessed')
    if (run/'prepared').exists():shutil.copytree(run/'prepared',dest/'prepared')
    shutil.copy2(form or run/'03_反馈提交.xlsx',dest/'submitted_feedback.xlsx')
    approved.update(entries=entries,organization=organization,parent_plan_sha256=s['approved_plan_sha256']);write(dest/'approved_plan.json',approved)
    new=deepcopy(s);new.update(run_id=uuid4().hex,created=stamp(),status='REVISING',parent=str(run),drafts=[],approved_plan_sha256=sha(dest/'approved_plan.json'))
    core_run=dest/'draft_001';state=revise(parent,core_run)
    state['job']['organization']=organization;order=[r['source_id'] for r in organization['sources']]
    state['sources'].sort(key=lambda r:order.index(r['id']))
    for sid,decisions in prepared_decisions.items():
        folder=core_run/'sources'/sid;candidate=resolve(read(folder/'scan/scan.json'),decisions)
        write(folder/'decisions.json',decisions);write(folder/'candidate.json',candidate)
        state['stages']['recognize:'+sid].update(evidence_sha256=sha(folder/'candidate.json'),status='FEEDBACK_APPLIED')
    state.setdefault('layout_overrides',[]).extend(layout);refresh_preparation_status(core_run,state);core_save(core_run,state)
    state=finish(core_run,state,retain=True);checked=status(core_run,state);write(dest/'draft_001.status.json',checked)
    for item in ledger:
        if item['status']!='awaiting_generation':continue
        item['status']='execution_issue'
        if checked['receipt_integrity']=='PASS' and state['status']!='NEEDS_REPAIR':
            if item.get('retain'):item['status']='closed_retained_in_checked_draft'
            else:
                action=next((a for a in actions if a['id']==item['id']),None)
                applied=bool(action)
                if action and action['kind']=='structure':
                    row=next(r for r in read(core_run/'sources'/item['source_id']/'candidate.json')['blocks'] if r['block_id']==item['block_id'])
                    d=action['decision'];applied=row['status']!='review' and row['role']==d['role'] and (d['role']!='heading' or row.get('level')==d['level'])
                item['status']='closed_in_checked_draft' if applied else 'awaiting_effective_resolution'
    history=deepcopy(s.get('feedback_history',[]));history.extend(dict(x,submitted_parent=str(run)) for x in ledger)
    new.update(status=state['status'],feedback_history=history,drafts=[dict(run='draft_001',status=state['status'],state_sha256=sha(core_run/'state.json'))])
    new['events'].append(dict(stage='feedback_batch',time=stamp(),submitted=len(ledger),applied=len(actions),core_builds=1))
    write(dest/'feedback_ledger.json',ledger);save(dest,new);export_classification(dest,new,inv)
    receipt_rows=[[x['id'],x.get('document','')+' | '+x.get('locator',''),(x.get('submitted_values') or ['']*9)[2],(x.get('submitted_values') or ['']*9)[3],(x.get('submitted_values') or ['']*9)[4],x.get('problem',''),(x.get('submitted_values') or ['']*9)[6],'',x['status']] for x in ledger]
    write_form(dest/'反馈处理回执.xlsx','反馈回执',dict(schema=SCHEMA,run_id=new['run_id']),receipt_rows,FH,set(),
               '本批反馈处理回执','只读结果；下一批请填写新 03_反馈提交.xlsx。本回执不作为提交表。',widths=[24,48,60,20,26,48,16,18,38],global_label='批次状态',global_value=state['status'])
    export_feedback(dest,new,core_run,state);write_summary(dest,new,inv)
    return forms_status(dest)
