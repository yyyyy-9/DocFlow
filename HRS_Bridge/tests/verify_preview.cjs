// Verify an actual offline preview against scanner text/media, without mocked APIs.
'use strict';
const fs=require('node:fs'),path=require('node:path'),{pathToFileURL}=require('node:url');
const {chromium}=require(path.resolve(path.dirname(process.execPath),'../node_modules/playwright-core'));
(async()=>{
 const [html,scanFile,out,originFile]=process.argv.slice(2);fs.mkdirSync(out,{recursive:false});
 const removed=originFile?JSON.parse(fs.readFileSync(originFile,'utf8')).word_styles.filter(v=>v.projection==='heading-list'&&v.source_number_prefix).map(v=>v.source_number_prefix):[];
 const scan=JSON.parse(fs.readFileSync(scanFile,'utf8'));
 const browser=await chromium.launch({executablePath:process.env.CHROMIUM_PATH,headless:true});
 try{
  const page=await browser.newPage({viewport:{width:1400,height:1000}}),errors=[];
  page.on('pageerror',e=>errors.push(String(e)));
  await page.goto(pathToFileURL(path.resolve(html)).href,{waitUntil:'load'});
  await page.evaluate(async()=>{await document.fonts.ready;await Promise.all([...document.images].map(i=>i.decode().catch(()=>{})));});
  const inspection=await page.evaluate(()=>({
   images:[...document.querySelectorAll('main img')].map(i=>({ok:i.complete&&i.naturalWidth>0})),
   h1:document.querySelectorAll('main h1').length,
   math:document.querySelectorAll('main .markdown-math svg').length,
   text:document.querySelector('main').innerText
  }));
  const expectedImages=Object.values(scan.objects).filter(v=>v.kind==='drawing').length;
  const normalize=t=>t.replace(/\s/g,''),body=normalize(inspection.text),missing=[],normalizations=[];
  for(const b of scan.blocks){
   if(b.kind!=='paragraph')continue;
   if(b.segments.some(s=>s.kind==='math_inline'||s.kind==='math_display')){
    for(const s of b.segments)if(s.kind==='text'&&normalize(s.text)&&!body.includes(normalize(s.text)))missing.push(b.block_id);
   }else{
    let text=b.plain_text||'';
    const prefix=b.outline_level_raw===5?removed.find(p=>text.startsWith(p)):null;
    if(prefix){text=text.slice(prefix.length);normalizations.push({block:b.block_id,removed_prefix:prefix});}
    if(normalize(text)&&!body.includes(normalize(text)))missing.push(b.block_id);
   }
  }
  if(inspection.images.length!==expectedImages||inspection.images.some(i=>!i.ok)||inspection.h1!==1||inspection.math!==2||errors.length||missing.length)
   throw Error(JSON.stringify({...inspection,text:undefined,expectedImages,errors,missing}));
  await page.screenshot({path:path.join(out,'preview.png'),fullPage:true});delete inspection.text;
  const result={status:'PASS',...inspection,errors,source_text_coverage:'PASS',normalizations};
  fs.writeFileSync(path.join(out,'CHECK.json'),JSON.stringify(result,null,2));console.log(JSON.stringify(result));
 }finally{await browser.close();}
})().catch(e=>{console.error(e);process.exitCode=1;});
