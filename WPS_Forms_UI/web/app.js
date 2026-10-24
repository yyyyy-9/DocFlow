'use strict';
const $ = id => document.getElementById(id);
const params = new URLSearchParams(location.hash.slice(1));
let token = params.get('token') || sessionStorage.getItem('wps-ui-token') || '';
if (token) sessionStorage.setItem('wps-ui-token', token);
history.replaceState(null, '', '/');
window.addEventListener('hashchange',()=>{const next=new URLSearchParams(location.hash.slice(1)).get('token');if(next){token=next;sessionStorage.setItem('wps-ui-token',next);history.replaceState(null,'','/');refresh(true).catch(e=>notice(e.message,true));}});
let state = {}, forms = {}, edits = {'02': {}, '03': {}}, currentTab = '01', issueRow = null, context = null;
let logOffset = 0, logLines = [], polling = false, displayedRun = null, contextGeneration = 0;
let previousJob = '';
let wordOpening = false;
let templateOpening = false, previousTemplateState = '';
$('writer').value=sessionStorage.getItem('wps-ui-writer')||'user';
$('writer').onchange=()=>sessionStorage.setItem('wps-ui-writer',$('writer').value);
const types = ['','标题级别','改为正文','改为列表','题名关联','字号','对齐','首行缩进','段前','段后','行距','目标归类','单元顺序','内容缺口','其他'];
const help = {'标题级别':'填来源标题级别 1—5；总稿会叠加来源与分组层级。','题名关联':'填写同来源真实对象，如 图:b000007 或 表:b000009。','字号':'单位磅，5—48。','首行缩进':'单位磅，0—144；不是字符数。','段前':'单位磅，0—144。','段后':'单位磅，0—144。','行距':'倍数，0.5—3；1 为单倍。','对齐':'左对齐 / 居中 / 右对齐 / 两端对齐。','目标归类':'用 / 分层。空白表示移到根层。','单元顺序':'1—999999 的整数。','改为正文':'保留原文；填写值留空。','改为列表':'保留原文与列表身份；填写值留空。','内容缺口':'记录明确待处理事项，不自动补写技术内容。','其他':'记录明确事项，不将说明擅自变为全局规则。'};

function notice(message, error=false) { $('notice').hidden=!message; $('notice').textContent=message||''; $('notice').title=message||''; $('notice').classList.toggle('error',error); }
async function api(path, body, raw=false) {
  const response=await fetch(path,{method:body===undefined?'GET':'POST',headers:{'X-WPS-UI-Token':token,...(body===undefined?{}:{'Content-Type':'application/json'})},...(body===undefined?{}:{body:JSON.stringify(body)})});
  if(raw&&response.ok)return response.blob();
  const result=await response.json();
  if(!response.ok){const error=new Error(result.message||'请求失败');error.detail=result;throw error;}
  return result;
}
function runAction(action){return async()=>{try{await action();}catch(e){notice(e.message,true);}};}
function runName(path){return path?String(path).split(/[\\/]/).pop():'';}
function localKey(kind){return 'wps-ui-edits:'+state.current+':'+kind;}
function hasLocalEdits(kind){return kind?Object.keys(edits[kind]||{}).length>0:hasLocalEdits('02')||hasLocalEdits('03');}
function hasLocalLocator(row=issueRow){return row!==null&&['B','C'].some(col=>Object.hasOwn(edits['03'],col+row));}
function pendingWork(kind){return hasLocalEdits(kind)||!!forms[kind]?.dirty||!!forms[kind]?.external_changed;}
function anyPendingWork(){return pendingWork('02')||pendingWork('03');}
function isRunning(){return state.job?.status==='RUNNING';}
function formsReady(){return !!state.current&&['02','03'].every(kind=>forms[kind]?.run===state.current);}
function hasMaster(){return !!state.master&&!state.task_error;}
function canConfirmFinal(){return !isRunning()&&hasMaster()&&formsReady()&&state.status!=='NEEDS_REPAIR'&&!anyPendingWork();}
function canUseFinal(){return !!state.final&&canConfirmFinal();}
function cell(kind,row,col){const key=col+row.row;return Object.hasOwn(edits[kind],key)?edits[kind][key]:(row.values[col]||'');}
function markEdit(kind,cell,value){edits[kind][cell]=value;localStorage.setItem(localKey(kind),JSON.stringify(edits[kind]));if(kind==='03'&&hasLocalLocator())invalidateLocation();renderStatus(kind);}
function renderStatus(kind){
  if(!forms[kind])return;const f=forms[kind],file=runName(f.path);
  $('status-'+kind).textContent=(hasLocalEdits(kind)?'界面有未同步修改 · ':f.dirty?'已同步，尚未生成新稿 · ':'已与本轮表单一致 · ')+(f.external_changed?'Excel 已在外部改变，请同步 · ':'')+runName(f.run)+' / '+file;
  $('status-'+kind).title=f.path;renderWorkflow();
}
function renderFinal(){
  const valid=canUseFinal(),final=state.final;
  $('final-tab').disabled=!valid;$('training').disabled=!valid;$('delivery-open').disabled=!valid||!state.master||wordOpening;
  if(valid){
    $('final-info').textContent=JSON.stringify(final,null,2);
    $('final-summary').textContent='已确认总稿：'+runName(final.run)+'\n确认来源：'+(final.human_confirmed?'人工':'Agent代填')+'\n本地回执检查：通过；接受保留的待核事项：'+final.pending_accepted+' 项。';
  }else{
    $('final-info').textContent='';$('final-summary').textContent=state.final&&anyPendingWork()?'当前有待处理修改，终稿入口暂不可用。提交生成的新稿需要重新确认；明确撤销修改后，按原稿绑定核对已有确认。':isRunning()?'当前批次仍在执行；完成后按实际交付稿核对确认。':'';$('training-result').textContent='';
  }
}
function setTabState(tab,text,status=''){const button=document.querySelector('[data-tab="'+tab+'"]');$('step-state-'+tab).textContent=text;button.dataset.state=status;}
function renderWorkflow(){
  const running=isRunning(),ready=formsReady(),delivered=hasMaster();
  let batch='从 01 接入文件夹开始';
  if(running)batch=state.job.phase||'正在生成新稿；已有稿件保留。';
  else if(state.task_error)batch='当前任务读取未完成；查看运行日志，已有稿件保留。';
  else if(state.current&&!ready)batch='正在载入当前稿的 02／03 表单。';
  else if(state.current&&!delivered)batch='本轮尚无可用总稿，请查看执行记录与缺口；已有稿件保留。';
  else if(state.current){
    const local=['02','03'].filter(hasLocalEdits),dirty=['02','03'].filter(kind=>forms[kind]?.dirty),external=['02','03'].filter(kind=>forms[kind]?.external_changed);
    batch=local.length?'有未同步修改（'+local.join('、')+'）；同步后才能生成新稿。':external.length?'Excel 有外部修改（'+external.join('、')+'），请先同步。':dirty.length?'修改已同步（'+dirty.join('、')+'），尚未生成新稿。':'当前稿与本批表单一致。';
    if(canUseFinal())batch+=' 终稿已确认。';
  }
  $('batch-state').textContent=batch;
  setTabState('01',delivered?'已有总稿':running?'正在生成':state.current?'待交稿':'从这里开始',delivered?'done':'');
  setTabState('02',!state.current?'等待首稿':!ready?'载入中':hasLocalEdits('02')?'有未同步修改':forms['02']?.external_changed?'Excel 待同步':forms['02']?.dirty?'待生成新稿':'可调整分类',pendingWork('02')?'attention':ready?'done':'');
  setTabState('03',!state.current?'等待首稿':!ready?'载入中':hasLocalEdits('03')?'有未同步修改':forms['03']?.external_changed?'Excel 待同步':forms['03']?.dirty?'待生成新稿':issueRows().length?issueRows().length+' 条待核事项':'可新增意见',pendingWork('03')?'attention':'');
  setTabState('04',canUseFinal()?'已确认':running?'等待本批完成':state.final&&anyPendingWork()?'先处理修改':'待最终确认',canUseFinal()?'done':'');
  $('classify').disabled=running||!ready;
  $('finalize').disabled=!canConfirmFinal()||canUseFinal();
  $('finalize').textContent=canUseFinal()?'当前稿已确认':'确认当前稿为最终版本';
  $('feedback').disabled=running||!ready||(!hasLocalEdits('03')&&!forms['03']?.dirty&&!forms['03']?.external_changed);
  $('finalize-hint').textContent=running?'本批仍在生成，请先等待交付。':!delivered?'实际总稿生成后才可确认完成。':state.status==='NEEDS_REPAIR'?'当前草稿可查看和继续修订；执行或内容缺口尚未修复，暂不能标记为最终完成。':anyPendingWork()?'两页修改先同步并提交生成新稿，或明确撤销；已有总稿可随时查看。':state.pending?'可以继续反馈；如已完成核查，也可明确接受保留事项后确认当前稿。':'如果当前稿已符合要求，可以直接确认完成，无需制造一轮空反馈。';
  const next=$('next-step');next.hidden=true;next.onclick=null;
  if(delivered&&!running&&ready&&!anyPendingWork()){
    const show=(label,action)=>{next.textContent=label;next.hidden=false;next.onclick=action;};
    if(canUseFinal()&&currentTab!=='04')show('查看完成与交付',()=>changeTab('04'));
    else if(currentTab==='01')show('下一步：检查分类与顺序',()=>changeTab('02'));
    else if(currentTab==='02')show('下一步：进入反馈与修订',()=>changeTab('03'));
    else if(currentTab==='03'&&canConfirmFinal())show('下一步：确认当前稿完成',()=>document.querySelector('.final-box').scrollIntoView({behavior:'smooth',block:'center'}));
    else if(currentTab==='04')show('打开最终总稿',runAction(()=>openWords('target',false,null)));
  }
  renderFinal();
}
function changeTab(tab){
  if(tab==='04'&&!canUseFinal()){notice(isRunning()?'本批仍在执行，请先等待交付。':state.final&&anyPendingWork()?'请先处理待提交修改。提交的新稿需重新确认；撤销修改后按原绑定核对。':'请先确认当前稿为最终版本。');return;}
  if((tab==='02'||tab==='03')&&!state.current){notice('先生成首稿或继续已有任务。');return;}
  currentTab=tab;document.querySelectorAll('[data-tab]').forEach(b=>b.classList.toggle('active',b.dataset.tab===tab));for(const t of ['01','02','03','04'])$('tab-'+t).hidden=t!==tab;
  if(tab==='02'&&!forms['02'])loadForm('02').catch(e=>notice(e.message,true));if(tab==='03'&&!forms['03'])loadForm('03').catch(e=>notice(e.message,true));renderWorkflow();
}
async function refresh(force=false){
  const fresh=await api('/api/state');state=fresh;
  $('environment').textContent='界面 '+fresh.ui_version+' / 三表 '+fresh.forms_version+' / 内核协议 '+fresh.kernel_protocol+'\n界面构建 '+(fresh.ui_build||'旧服务，请退出服务后重新启动')+'\n三表构建 '+fresh.core_build+'\n内核目录 '+fresh.core+'\nPython '+fresh.python+'\n本机数据 '+fresh.data;
  const tool=fresh.template_tool,toolKey=tool?(tool.id||'')+':'+tool.status:'';
  $('template-tool').disabled=templateOpening||!tool?.available;
  $('template-tool-status').textContent=!tool?'模板工具需要新版服务，请退出旧服务后重新启动。':!tool.available?tool.message:tool.status==='STARTING'?'正在启动模板样式窗口…':tool.status==='RUNNING'?'模板样式工具 '+tool.version+' 已打开；生成新模板后，在上方选择该文件。':tool.status==='ERROR'?tool.message:'可在独立弹窗中保留封面、目录和页眉，选择正文起点生成新模板；生成后在上方选择该文件。';
  $('template-tool-status').title=tool?.log||'';
  if(toolKey!==previousTemplateState&&tool?.status==='ERROR')notice(tool.message+' 日志：'+tool.log,true);
  previousTemplateState=toolKey;
  $('task-title').textContent=fresh.title||'新建整理任务';
  const running=fresh.job&&fresh.job.status==='RUNNING';
  const statusText={COMPLETED_LOCAL_CHECKS:'已交付 · 本地检查通过',COMPLETED_WITH_RETAINED_ITEMS:'已交付 · 有保留项',NEEDS_REPAIR:'草稿待修补 · 查看缺口与执行记录'};
  $('stage').textContent=running?fresh.job.phase:(fresh.task_error||(fresh.current?runName(fresh.current)+' · '+(statusText[fresh.status]||fresh.status||'已交付'):'等待启动'));
  $('stop').disabled=!running;for(const id of ['start','classify','feedback','finalize','attach'])$(id).disabled=!!running;
  $('open-master').disabled=!fresh.master||wordOpening;$('open-master').textContent=fresh.master?'打开当前总稿 · '+runName(fresh.current):'尚未生成';
  for(const button of document.querySelectorAll('[data-tab="02"],[data-tab="03"]'))button.disabled=!fresh.current;
  $('accept-pending-wrap').hidden=!fresh.pending;if(!fresh.pending)$('accept-pending').checked=false;
  if(fresh.task_error)notice(fresh.task_error,true);
  const jobKey=fresh.job?fresh.job.id+':'+fresh.job.status:'';
  if(jobKey!==previousJob&&fresh.job?.status.startsWith('DELIVERED')){notice('新稿已交付：'+runName(fresh.job.output)+'。界面与新 02／03 已切换到这一轮。');$('notice').title=fresh.job.output||'';}
  else if(jobKey!==previousJob&&fresh.job?.status==='ERROR')notice(fresh.job.message||'本批执行失败，旧稿保留。',true);
  previousJob=jobKey;
  if(fresh.current!==displayedRun||force){
    displayedRun=fresh.current;forms={};context=null;
    for(const kind of ['02','03']){try{edits[kind]=JSON.parse(localStorage.getItem(localKey(kind))||'{}');}catch{edits[kind]={};}}
    if(fresh.current){await loadForm('02');await loadForm('03');}
    if(fresh.current&&fresh.job?.status.startsWith('DELIVERED')){
      if(currentTab==='01')changeTab(fresh.job.operation==='scan'?'02':'03');
      else if(currentTab==='02'&&fresh.job.operation==='classify')changeTab('03');
    }
  }
  if(!canUseFinal()&&currentTab==='04')changeTab('03');
  renderWorkflow();
}
async function loadForm(kind){forms[kind]=await api('/api/form?kind='+kind);renderStatus(kind);if(kind==='02')renderClassification();else renderIssues();}
function inputControl(kind,row,col,options){
  const node=document.createElement(options?'select':'input'),names={E:'处理决定',F:'目标归类',G:'同级顺序',H:'展示题名',I:'备注'};
  node.setAttribute('aria-label',(row.values.B||'第 '+row.row+' 行')+'，'+(names[col]||col));
  if(options)for(const value of options){const o=document.createElement('option');o.value=value;o.textContent=value;node.append(o);}node.value=cell(kind,row,col);node.addEventListener(options?'change':'input',()=>markEdit(kind,col+row.row,node.value));return node;
}
function renderClassification(){
  const body=$('classification').querySelector('tbody');body.replaceChildren();
  for(const row of forms['02'].rows){const tr=document.createElement('tr'),fact=document.createElement('td');fact.textContent=row.values.B;const sub=document.createElement('small');sub.textContent=row.values.C+' · '+row.values.D;fact.append(sub);tr.append(fact);
    for(const col of ['E','F','G','H','I']){const td=document.createElement('td');td.append(inputControl('02',row,col,col==='E'?['接受建议','整篇保留','按建议拆分','不纳入']:null));tr.append(td);}body.append(tr);}
}
function issueRows(){return (forms['03']?.rows||[]).filter(r=>r.values.A||cell('03',r,'B')||cell('03',r,'C')||cell('03',r,'D')||cell('03',r,'F'));}
function currentRow(){return forms['03']?.rows.find(r=>r.row===issueRow);}
function renderIssues(){
  const rows=issueRows(),select=$('issue-select');select.replaceChildren();
  for(const [i,r]of rows.entries()){const o=document.createElement('option');o.value=r.row;o.textContent=(i+1)+' / '+rows.length+' · '+(r.values.C||cell('03',r,'C')||r.values.F||'新增反馈').slice(0,70);select.append(o);}
  if(!rows.some(r=>r.row===issueRow))issueRow=rows[0]?.row??null;select.value=issueRow;
  $('review-nav').hidden=!rows.length;$('previous').disabled=rows.length<2;$('next').disabled=rows.length<2;
  $('empty-review').hidden=!!rows.length;$('issue-panel').hidden=!rows.length;
  if(issueRow!==null)showIssue().catch(e=>notice(e.message,true));renderWorkflow();
}
function drawSnippet(target,part){target.replaceChildren();if(!part){target.textContent='尚无唯一位置；请核对技术定位或新增意见原文。';return;}
  for(const text of part.before||[]){const p=document.createElement('p');p.className='muted';p.textContent=text||'（空段）';target.append(p);}
  const focus=document.createElement('p');focus.className='focus';focus.textContent=part.text||'（图片、表格、公式或无可见普通文字的块；打开Word核对）';target.append(focus);
  for(const text of part.after||[]){const p=document.createElement('p');p.className='muted';p.textContent=text||'（空段）';target.append(p);}
}
function invalidateLocation(){
  ++contextGeneration;context=null;
  $('location-status').className='status-chip unlocated';$('location-status').textContent='定位输入待同步';
  $('source-file').textContent='已修改定位输入，请先同步并定位。';$('source-chapter').textContent='—';$('target-chapter').textContent='—';$('technical').textContent='';
  drawSnippet($('source-snippet'),null);drawSnippet($('target-snippet'),null);
}
async function showIssue(){
  const row=currentRow();if(!row)return;const generation=++contextGeneration,location=$('location-status');
  $('issue-select').value=issueRow;$('issue-title').textContent=row.values.C||cell('03',row,'C')||'新增局部意见';$('issue-message').textContent=cell('03',row,'F');
  for(const [id,col]of [['edit-type','D'],['edit-value','E'],['edit-problem','F'],['edit-scope','G'],['edit-decision','H'],['new-document','B'],['new-quote','C']])$(id).value=cell('03',row,col)||(col==='G'?'仅此处':col==='H'?'暂缓':'');
  $('new-locator').hidden=!!row.values.A;[...$('edit-decision').options].find(o=>o.value==='保留原样').disabled=!row.values.A;
  location.className='status-chip';location.textContent='正在核对位置';
  valueHelp();context=null;drawSnippet($('source-snippet'),null);drawSnippet($('target-snippet'),null);
  if(hasLocalLocator(row.row)){invalidateLocation();return;}
  try{
    const c=await api('/api/context?row='+issueRow);if(generation!==contextGeneration)return;context=c;
    const located=c.status==='LOCATED';location.textContent=located?'已唯一定位':'未唯一定位';location.classList.add(located?'located':'unlocated');
    $('source-file').textContent=c.source?.relative||c.message||'未定位';$('source-chapter').textContent=c.source?.chapter||'—';$('target-chapter').textContent=c.target?.chapter||'—';
    $('technical').textContent=JSON.stringify({run:state.current,locator:c.locator,source:c.source?.path,source_sha256:c.source?.sha256,target:c.target?.path,target_sha256:c.target?.sha256,status:c.status},null,2);
    $('target-caption').textContent='当前总稿 · '+runName(state.current);drawSnippet($('source-snippet'),c.source);drawSnippet($('target-snippet'),c.target?.body_index===undefined?null:c.target);
  }catch(e){if(generation===contextGeneration){location.textContent='定位失败';location.classList.add('unlocated');$('source-file').textContent=e.message;notice(e.message,true);}}
}
function valueHelp(){const type=$('edit-type').value;$('value-help').textContent=help[type]||'默认仅此处；本单元/全文会保留为需明确计划的反馈。';const options=$('value-options');options.replaceChildren();for(const v of type==='对齐'?['左对齐','居中','右对齐','两端对齐']:type==='标题级别'?['1','2','3','4','5']:[]){const o=document.createElement('option');o.value=v;options.append(o);}}
async function resolveConflicts(items){
  const list=$('conflict-list');list.replaceChildren();
  for(const [i,item]of items.entries()){const box=document.createElement('div');box.className='conflict-item';const strong=document.createElement('strong');strong.textContent=item.cell+' · 原值：'+item.base;box.append(strong);for(const side of ['local','remote']){const label=document.createElement('label'),radio=document.createElement('input');radio.type='radio';radio.name='conflict-'+i;radio.value=side;label.append(radio,document.createTextNode((side==='local'?'界面：':'Excel：')+item[side]));box.append(label);}list.append(box);}
  $('conflict-dialog').showModal();
  return new Promise(resolve=>{
    $('cancel-conflict').onclick=()=>{$('conflict-dialog').close();resolve(null);};
    $('resolve-conflict').onclick=()=>{const result={};for(const[i,item]of items.entries()){const radio=list.querySelector('input[name="conflict-'+i+'"]:checked');if(!radio){notice('请为每个冲突选择保留值。',true);return;}result[item.cell]=radio.value;}$('conflict-dialog').close();resolve(result);};
    $('conflict-dialog').oncancel=()=>resolve(null);
  });
}
async function synchronize(kind,importPath=null){
  if(!forms[kind])await loadForm(kind);
  const payload={kind,version:forms[kind].version,changes:edits[kind],origin:$('writer').value,...(importPath?{import_path:importPath}:{})};
  let result;
  try{result=await api('/api/sync',payload);}catch(e){
    if(e.detail?.conflicts){forms[kind].version=e.detail.version;const resolutions=await resolveConflicts(e.detail.conflicts);if(!resolutions)throw new Error('合并暂缓；界面修改已保存，Excel保持原样。');result=await api('/api/sync',{...payload,version:e.detail.version,resolutions});}
    else throw e;
  }
  edits[kind]={};localStorage.removeItem(localKey(kind));forms[kind]=result;renderStatus(kind);if(kind==='02')renderClassification();else renderIssues();
  notice('已同步 '+kind+'：界面、Excel和后续提交使用同一份记录。');return result;
}
async function submit(operation){const kind=operation==='classify'?'02':'03';const form=await synchronize(kind);await api('/api/submit',{operation,version:form.version,origin:$('writer').value});notice('已冻结本批输入。程序执行完成后会自动切换到新稿。');await refresh();}
async function openWords(side='both',plain=false,row=issueRow){
  if(row!==null&&hasLocalLocator(row))throw new Error('定位输入尚未同步，请先点击“同步并定位新增意见”。');
  if(wordOpening)return;wordOpening=true;
  const ids=['open-both','open-source','open-target','plain-both','open-master','delivery-open'];
  for(const id of ids)$(id).disabled=true;
  notice('正在打开'+(plain?'只读查看副本':'WPS并定位两端内容')+'…');
  try{const result=await api('/api/open-word',{...(row===null?{}:{row}),side,plain});const details=result.documents.map(d=>(d.status||result.status)+(d.selected_text?' · '+d.selected_text.slice(0,70):'')).join('\n');notice('Word 对照：'+details);}
  finally{wordOpening=false;$('open-master').disabled=!state.master;for(const id of ids.filter(id=>id!=='open-master'&&id!=='delivery-open'))$(id).disabled=false;renderWorkflow();}
}
function renderLogs(){const filter=$('log-filter').value;const rows=logLines.filter(x=>filter==='all'||(filter==='error'?x.stage==='错误'||/ERROR|FAILED|FAIL/.test(x.message):x.stage!=='程序输出'));$('log-output').textContent=rows.slice(-200).map(x=>x.time.slice(11,19)+' ['+x.stage+'] '+x.message).join('\n');$('log-count').textContent=logLines.length+' 条实际输出 · 本栏显示末200条，完整记录留在日志文件';if($('follow').checked)$('log-output').scrollTop=$('log-output').scrollHeight;}
async function poll(){if(polling)return;polling=true;try{await refresh();const logs=await api('/api/logs?after='+logOffset);logOffset=logs.count;if(logs.lines.length){logLines.push(...logs.lines);renderLogs();}}catch(e){notice('连接未完成：'+e.message+'。已有稿件和工作记录保留。',true);}finally{polling=false;}}
function clearRootError(){$('root').removeAttribute('aria-invalid');$('root-error').hidden=true;$('root-error').textContent='';}
$('root').addEventListener('input',clearRootError);
for(const button of document.querySelectorAll('[data-tab]'))button.onclick=()=>changeTab(button.dataset.tab);
for(const button of document.querySelectorAll('[data-choose]'))button.onclick=runAction(async()=>{const result=await api('/api/choose',{kind:button.dataset.kind});if(result.path){$(button.dataset.choose).value=result.path;if(button.dataset.choose==='root')clearRootError();}});
for(const button of document.querySelectorAll('[data-sync]'))button.onclick=runAction(()=>synchronize(button.dataset.sync));
for(const button of document.querySelectorAll('[data-open-form]'))button.onclick=runAction(()=>api('/api/open-form',{kind:button.dataset.openForm}));
for(const button of document.querySelectorAll('[data-import]'))button.onclick=runAction(async()=>{const result=await api('/api/choose',{kind:'xlsx'});if(result.path)await synchronize(button.dataset.import,result.path);});
for(const button of document.querySelectorAll('[data-export]'))button.onclick=runAction(async()=>{const kind=button.dataset.export;await synchronize(kind);const blob=await api('/api/export-form',{kind},true);const url=URL.createObjectURL(blob),a=document.createElement('a');a.href=url;a.download=kind+'_表单副本.xlsx';a.click();setTimeout(()=>URL.revokeObjectURL(url),30000);});
for(const button of document.querySelectorAll('[data-reset]'))button.onclick=runAction(async()=>{const kind=button.dataset.reset;if(!confirm('明确撤销本页所有尚未提交的修改？已交总稿不改变。'))return;forms[kind]=await api('/api/reset-form',{kind,version:forms[kind].version});edits[kind]={};localStorage.removeItem(localKey(kind));if(kind==='02')renderClassification();else renderIssues();renderStatus(kind);});
for(const type of types){const option=document.createElement('option');option.value=type;option.textContent=type||'选择修改类型';$('edit-type').append(option);}
for(const[id,col]of [['edit-type','D'],['edit-value','E'],['edit-problem','F'],['edit-scope','G'],['edit-decision','H'],['new-document','B'],['new-quote','C']])$(id).addEventListener(id==='edit-problem'||id==='edit-value'||id.startsWith('new-')?'input':'change',()=>{if(issueRow!==null)markEdit('03',col+issueRow,$(id).value);if(id==='edit-type')valueHelp();});
$('start').onclick=runAction(async()=>{const root=$('root').value.trim();clearRootError();if(!root){$('root-error').textContent='请先选择或填写要整理的文件夹。';$('root-error').hidden=false;$('root').setAttribute('aria-invalid','true');$('root').focus();return;}await api('/api/submit',{operation:'scan',root,title:$('title').value.trim(),out:$('out').value.trim(),template:$('template').value.trim(),pages:$('pages').checked,doc_conversion:$('doc-conversion').checked});notice('正在执行首稿链；进度显示在右侧日志。');await refresh();});
$('template-tool').onclick=runAction(async()=>{templateOpening=true;$('template-tool').disabled=true;try{const result=await api('/api/open-template-tool',{});notice(result.message||'正在启动独立模板样式窗口；可从任务栏切换。生成后在01选择新模板。');}finally{templateOpening=false;await refresh();}});
$('attach').onclick=runAction(async()=>{await api('/api/attach',{run:$('resume').value.trim()});await refresh(true);changeTab('03');});
$('classify').onclick=runAction(()=>submit('classify'));$('feedback').onclick=runAction(()=>submit('feedback'));
$('issue-select').onchange=()=>{issueRow=Number($('issue-select').value);showIssue().catch(e=>notice(e.message,true));};
function moveIssue(delta){const rows=issueRows(),i=rows.findIndex(r=>r.row===issueRow);const row=rows[i+delta];if(row){issueRow=row.row;showIssue().catch(e=>notice(e.message,true));}}
$('previous').onclick=()=>moveIssue(-1);$('next').onclick=()=>moveIssue(1);
function addIssue(){const row=forms['03']?.rows.find(r=>!r.values.A&&!cell('03',r,'B')&&!cell('03',r,'C')&&!cell('03',r,'D')&&!cell('03',r,'F'));if(!row)return notice('本轮新增空位已用完，请提交这一批后继续。',true);issueRow=row.row;markEdit('03','B'+row.row,'master');renderIssues();}
$('add-issue').onclick=addIssue;$('empty-add-issue').onclick=addIssue;
$('locate-new').onclick=runAction(async()=>{await synchronize('03');await showIssue();});
$('open-both').onclick=runAction(()=>openWords());$('plain-both').onclick=runAction(()=>openWords('both',true));$('open-source').onclick=runAction(()=>openWords('source'));$('open-target').onclick=runAction(()=>openWords('target'));$('open-master').onclick=runAction(()=>openWords('target',false,null));$('delivery-open').onclick=runAction(()=>openWords('target',false,null));
function setLogHidden(hidden){document.body.classList.toggle('hide-log',hidden);$('toggle-log').textContent=hidden?'展开日志':'收起日志';$('toggle-log').setAttribute('aria-expanded',String(!hidden));}
setLogHidden(window.matchMedia('(max-width:1100px)').matches);
$('toggle-log').onclick=()=>setLogHidden(!document.body.classList.contains('hide-log'));
$('log-filter').onchange=renderLogs;$('follow').onchange=renderLogs;
$('stop').onclick=runAction(async()=>{if(confirm('请求停止当前批次？旧稿保留，未完成的新目录留作现场，不自动重试。'))await api('/api/stop',{});});
$('reload').onclick=runAction(()=>refresh(true));
$('finalize').onclick=runAction(async()=>{await synchronize('02');await synchronize('03');await api('/api/finalize',{accept_pending:$('accept-pending').checked,origin:$('writer').value});await refresh();changeTab('04');notice('终稿确认完成。现在可以打开总稿，或可选保存本轮修正数据。');});
$('training').onclick=runAction(async()=>{if(!canUseFinal())throw new Error('当前有未处理修改，请重新确认终稿后再保存训练数据。');const result=await api('/api/training',{});$('training-result').textContent='已保存：'+runName(result.path)+'\n人工明确修正 '+result.explicit_corrections+' 条；模型候选 '+result.model_candidates+' 条；非标签记录 '+result.not_labels+' 条。';$('training-result').title=result.path;});
$('history').onclick=()=>{const list=$('history-list');list.replaceChildren();for(const item of state.history||[]){const box=document.createElement('div');box.className='history-item';const strong=document.createElement('strong');strong.textContent=item.title+' · '+item.run.split(/[\\/]/).pop();const path=document.createElement('pre');path.textContent=item.run;const button=document.createElement('button');button.textContent=item.run===state.current?'当前任务':'切换到此轮';button.disabled=item.run===state.current;button.onclick=runAction(async()=>{await api('/api/attach',{run:item.run});$('history-dialog').close();await refresh(true);changeTab('03');});box.append(strong,path,button);list.append(box);}$('history-dialog').showModal();};
$('close-history').onclick=()=>$('history-dialog').close();
$('quit').onclick=runAction(async()=>{if(confirm('退出本地服务？Word窗口和已交稿件保留。')){await api('/api/shutdown',{});notice('服务已退出。可以关闭此网页。');clearInterval(timer);}});
window.addEventListener('beforeunload',event=>{if(hasLocalEdits()){event.preventDefault();event.returnValue='';}});
renderWorkflow();poll();const timer=setInterval(poll,1500);
