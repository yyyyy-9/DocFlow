'use strict';
const path=require('node:path');
const {chromium}=require(require.resolve('playwright-core',{paths:[
 process.env.HRS_BRIDGE_NODE_MODULES||path.resolve(__dirname,'..'),path.resolve(path.dirname(process.execPath),'../node_modules')
]}));
(async()=>{
 const chunks=[];for await(const part of process.stdin)chunks.push(part);
 const req=JSON.parse(Buffer.concat(chunks).toString('utf8'));
 const browser=await chromium.launch({executablePath:req.browser,headless:true});
 try{
  const page=await browser.newPage({viewport:{width:1366,height:900},deviceScaleFactor:2});
  await page.goto(req.base+'/web/index.html?programmatic=1',{waitUntil:'load'});
  await page.waitForFunction(()=>window.ReportExportArchive);
  let result;
  if(req.operation==='tokens'){
   result=await page.evaluate(source=>marked.lexer(source),req.markdown);
  }else if(req.operation==='model'){
   // The archive library loads before the async application initialization.
   // Wait for the owner API instead of treating that interval as a failure.
   await page.waitForFunction(()=>!!window.ReportStudioProgrammatic,{},{timeout:30000});
   await page.addScriptTag({path:path.join(__dirname,'browser_model.js')});
   result=await page.evaluate(async req=>{
    await ReportStudioProgrammatic.ready();await ReportStudioProgrammatic.open('legacy');
    // The original renderer supplies graphical renditions only. Semantic types,
    // identities, references and source text still come from the Python IR.
    const diagrams={},svgAssets={};
    for(const fig of document.querySelectorAll('#preview-content figure[data-figure-id]')){
     const svg=fig.querySelector('.report-figure-graph svg');
     if(svg)diagrams[fig.dataset.figureId]=true;
    }
    for(const img of document.querySelectorAll('#preview-content img')){
     const u=new URL(img.src,location.href);if(/\.svg$/i.test(u.pathname))svgAssets[decodeURIComponent(u.pathname.split('/').pop())]=true;
    }
    return {...ReportBridgeModel.build(req),diagrams,svgAssets};
   },req);
   for(const id of Object.keys(result.diagrams))result.diagrams[id]=(await page.locator('#preview-content figure[data-figure-id="'+id+'"] .report-figure-graph svg').screenshot()).toString('base64');
   // Rasterize only actual SVG image elements, using the same browser rendition.
   for(const name of Object.keys(result.svgAssets)){
    const imgs=page.locator('#preview-content img');let found=false;
    for(let i=0;i<await imgs.count();i++)if(await imgs.nth(i).evaluate((img,name)=>decodeURIComponent(new URL(img.src).pathname.split('/').pop())===name,name)){
     result.svgAssets[name]=(await imgs.nth(i).screenshot()).toString('base64');found=true;break;
    }
    if(!found)throw Error('SVG rendition not found: '+name);
   }
  }else if(req.operation==='serialize'){
   result=await page.evaluate(items=>items.map(item=>{
    if(item.type==='image-group')return ReportLayout.groupSnippet(item.data)+'\n';
    if(item.type==='gfm')return ReportPipeTables.stringify(item.data,{id:item.id,withCaption:!!item.data.caption})+'\n';
    if(item.type==='equation')return ReportEquations.snippet(item.tex,item.id)+'\n';
    throw Error('Unknown source serializer');
   }),req.items);
  }else throw Error('Unknown browser operation');
  process.stdout.write(JSON.stringify(result));
 }finally{await browser.close();}
})().catch(error=>{process.stderr.write(error.stack||String(error));process.exitCode=1;});
