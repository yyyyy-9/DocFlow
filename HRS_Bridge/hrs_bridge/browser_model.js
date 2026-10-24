/* Canonical Bridge model adapter. HRS ships an exact generated copy; both the
   existing CLI browser and the open editor consume the same HRS lexer/profile. */
(function(root){
'use strict';
const SCHEMA='hrs-word-browser-model-v1';
function build(req){
 const profilePlan=ReportProfile.plan(req.markdown,'engineering');
 const notes=ReportNotes.prepare(marked,marked.lexer(SourceBlocks.normalized(req.markdown).text));
 if(notes.diagnostics.length)throw Error(notes.diagnostics.join('；'));
 const rich=(source,inline,base=null)=>{
  if(inline)source=ReportTables.compactDisplayText(source);
  const abstractIndices=new Set();
  if(!inline&&base!==null){
   const original=marked.lexer(SourceBlocks.normalized(source).text),records=SourceMap.create(source,original).blocks;
   for(const r of records)if(profilePlan.abstracts.some(a=>a.start===base+r.start))abstractIndices.add(original.indexOf(r.token));
  }
  const math=[],prefix='HRSMATH'+crypto.randomUUID().replaceAll('-','')+'X';
  const protectedSource=ReportTableContent.tokenize(source).map(t=>t.kind==='math'?(math.push(t),prefix+(math.length-1)+'X'):t.text).join('');
  const expand=tokens=>tokens.flatMap(token=>{
   if(token.type==='hrsStyledList')token.tokens=rich(token.body,false);
   if(token.tokens)token.tokens=expand(token.tokens);
   if(token.items)for(const item of token.items)item.tokens=expand(item.tokens);
   if(token.type==='text'&&!token.tokens){
    return token.text.split(new RegExp('('+prefix+'[0-9]+X)','g')).filter(Boolean).map(text=>{
     if(text.startsWith(prefix))return {type:'hrsMath',...math[Number(text.slice(prefix.length,-1))]};
     return {...token,text};
    });
   }
   return [token];
  });
  const parsed=inline?marked.Lexer.lexInline(protectedSource):marked.lexer(protectedSource);
  for(const index of abstractIndices)if(parsed[index]?.type==='blockquote')parsed[index].type='hrsAbstract';
  return ReportProfile.wordTokens(expand(parsed));
 };
 const result={schema:SCHEMA,numbering_policy:'document-v2',render_options:req.render_options,request_sha256:req.request_sha256,output_options:req.output_options,
  tokens:req.sections.map((source,i)=>rich(source,false,req.sectionStarts[i])),
  inline:req.inline.map(source=>rich(source,true)),profile:profilePlan,diagrams:{},svgAssets:{}};
 result.footnotes=Object.fromEntries([...notes.definitions].map(([key,body])=>[key,rich(body,false)]));
 const check=(value,inNote=false)=>{
  if(Array.isArray(value)){value.forEach(v=>check(v,inNote));return;}
  if(!value||typeof value!=='object')return;
  if(value.type==='hrsFootnoteRef'){
   if(inNote)throw Error('Word 不支持脚注中嵌套脚注：'+value.key);
   if(!Object.hasOwn(result.footnotes,value.key))throw Error('脚注未定义：'+value.key);
  }
  if(inNote&&value.type==='image')throw Error('Word 脚注暂不支持图片；请将图片放入正文。');
  for(const [key,child] of Object.entries(value))if(!['raw','text'].includes(key))check(child,inNote);
 };
 check(result.tokens);check(result.inline);Object.values(result.footnotes).forEach(v=>check(v,true));
 return result;
}
async function png(svg){
 const parsed=new DOMParser().parseFromString(svg,'image/svg+xml'),node=parsed.documentElement;
 if(node.localName!=='svg'||parsed.querySelector('parsererror'))throw Error('Word 图形不是有效 SVG');
 // Browser image decoding cannot execute SVG scripts. Freeze the bounded image
 // dimensions explicitly; Mermaid's percent/max-width attributes are for DOM.
 const box=(node.getAttribute('viewBox')||'').trim().split(/[ ,]+/).map(Number);
 let width=Number.parseFloat(node.getAttribute('width')),height=Number.parseFloat(node.getAttribute('height'));
 if(box.length===4&&box.every(Number.isFinite)&&box[2]>0&&box[3]>0){width=box[2];height=box[3];}
 if(!(width>0&&height>0&&Number.isFinite(width+height))){width=800;height=600;}
 const scale=Math.min(2,4000/width,4000/height,Math.sqrt(16000000/(width*height)));
 width=Math.max(1,Math.ceil(width*scale));height=Math.max(1,Math.ceil(height*scale));
 node.setAttribute('width',String(width));node.setAttribute('height',String(height));node.style.maxWidth='none';
 const blob=new Blob([new XMLSerializer().serializeToString(node)],{type:'image/svg+xml'});
 const url=URL.createObjectURL(blob),img=new Image();
 try{
  await new Promise((resolve,reject)=>{img.onload=resolve;img.onerror=()=>reject(Error('Word SVG 预览无法解码'));img.src=url;});
  const canvas=document.createElement('canvas');canvas.width=width;canvas.height=height;
  const context=canvas.getContext('2d');if(!context)throw Error('浏览器不能生成 Word 图形');
  context.fillStyle='#fff';context.fillRect(0,0,width,height);context.drawImage(img,0,0,width,height);
  const result=canvas.toDataURL('image/png').split(',')[1];
  if(result.length>16*1024*1024)throw Error('Word 图形预览超过单图上限');
  return result;
 }finally{URL.revokeObjectURL(url);}
}
async function render(req){
 const result=build(req);
 for(const [id,source] of Object.entries(req.diagram_sources||{}))result.diagrams[id]=await png(await ReportMermaid.render(source));
 for(const [name,encoded] of Object.entries(req.svg_sources||{})){
  const bytes=Uint8Array.from(atob(encoded),c=>c.charCodeAt(0));
  result.svgAssets[name]=await png(new TextDecoder().decode(bytes));
 }
 return result;
}
root.ReportBridgeModel={SCHEMA,build,render};
})(window);
