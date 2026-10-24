"""Synthetic native-numbering write/edit fixture through the production workflow.

Explicit fixture decisions exercise review with source hashes; they do not
measure unknown-document recognition or call a model.
"""
from pathlib import Path
from zipfile import ZipFile
import argparse,hashlib,json,sys,subprocess
from docx import Document
from docx.oxml import OxmlElement as O
from docx.oxml.ns import qn
from docx.shared import Inches
from PIL import Image

ROOT=Path(__file__).resolve().parents[1];sys.path.insert(0,str(ROOT/'src'))
from wps_pipeline import workflow,model_review
from wps_pipeline.verify_report_build import Package,verify
from wps_pipeline.verify_native_numbering import labels
from wps_pipeline.job_support import write,read


def e(tag,**attrs):
    node=O('w:'+tag)
    for k,v in attrs.items():node.set(qn('w:'+k),str(v))
    return node


def field(p,instruction,value,complex=False,bookmark=None):
    if bookmark:p._p.append(e('bookmarkStart',id=bookmark[0],name=bookmark[1]))
    if complex:
        for tag,content in (('begin',None),('code',instruction),('separate',None),('result',value),('end',None)):
            r=e('r')
            if tag=='code':t=e('instrText');t.text=content;r.append(t)
            elif tag=='result':t=e('t');t.text=content;r.append(t)
            else:r.append(e('fldChar',fldCharType=tag))
            p._p.append(r)
    else:
        f=e('fldSimple',instr=instruction);r=e('r');t=e('t');t.text=value;r.append(t);f.append(r);p._p.append(f)
    if bookmark:p._p.append(e('bookmarkEnd',id=bookmark[0]))


def source(folder):
    doc=Document();blueprint={};numbering=doc.part.numbering_part.element
    abstract=e('abstractNum',abstractNumId=200);abstract.append(e('multiLevelType',val='multilevel'))
    for level,fmt,text in ((0,'decimal','%1、'),(1,'lowerLetter','(%2)')):
        lvl=e('lvl',ilvl=level);lvl.extend([e('start',val=1),e('numFmt',val=fmt),e('lvlText',val=text),e('suff',val='space')]);abstract.append(lvl)
    numbering.append(abstract)
    for nid,start in ((200,3),(201,1)):
        n=e('num',numId=nid);n.append(e('abstractNumId',val=200));ov=e('lvlOverride',ilvl=0);ov.append(e('startOverride',val=start));n.append(ov);numbering.append(n)
    def p(text,style=None,role=None):
        para=doc.add_paragraph(text,style)
        if role:blueprint[len(doc.element.body)-2]=role
        return para
    p('原生编号演示','Heading 1');p('所有数值与图表均为合成示例，测量值2.3 dB和1.2毫米必须保留。')
    p('列表及重启','Heading 2');p('下面按两个操作步骤执行：')
    for text in ('1、安装设备','（1）检查接口','（2）记录状态','2、启动设备','（1）记录输出','3、关闭设备'):p(text,'List Paragraph','list_item')
    p('下面按新的步骤操作：')
    for text in ('1、连接电源','2、检查指示灯'):p(text,'List Paragraph','list_item')
    p('下面按编号步骤执行：')
    for text,nid,level in [('自动第三项',200,0),('自动子项甲',200,1),('自动子项乙',200,1),('自动第四项',200,0),('重启第一项',201,0),('原组继续第五项',200,0)]:
        para=p(text,'List Paragraph','list_item');np=e('numPr');np.extend([e('ilvl',val=level),e('numId',val=nid)]);para._p.get_or_add_pPr().append(np)
    p('图及引用','Heading 2');image=folder/'image.png';Image.new('RGB',(300,100),(45,100,155)).save(image)
    doc.add_picture(str(image),width=Inches(2.0));cap=p('图 ');field(cap,' SEQ Figure \\r 9 \\* ARABIC ','9',bookmark=(70,'FigNine'));cap.add_run(' 第一张合成图')
    ref=p('图');field(ref,' REF FigNine \\h ','9');ref.add_run('用于验证原生引用；普通文字引用见图9。')
    p('另一张图','Heading 2');doc.add_picture(str(image),width=Inches(2.0));cap=p('图 ');field(cap,' SEQ Figure \\r 12 \\* ARABIC ','12',True,(71,'FigTwelve'));cap.add_run(' 第二张合成图')
    ref=p('图');field(ref,' REF FigTwelve \\h ','12',True);ref.add_run('保留复杂域形式。')
    p('表格及深层标题','Heading 2');p('表 7 合成条件');table=doc.add_table(rows=2,cols=2)
    for c,text in zip([c for r in table.rows for c in r.cells],['参数','示例值','温度','25.0']):c.text=text
    p('参见表7，技术量值2.3 dB保持。')
    for level,text in [(3,'三级结构'),(4,'四级结构'),(5,'五级结构')]:p(text,'Heading '+str(level));p('此处保留正文并检查标题层级。')
    source=folder/'source.docx';doc.save(source);write(folder/'fixture_blueprint.json',{'source':str(source),'explicit_list_paragraph_indices':blueprint,'scope':'Authored synthetic write fixture; not recognition truth passed to scanner.'});return source,blueprint


def run(folder,source,blueprint,template):
    run=folder/'run';state=workflow.start_job({'sources':[{'id':'src','path':str(source)}],'make_master':True,'template':str(template)},folder,run)
    candidate=read(run/'sources/src/candidate.json');scan=read(run/'sources/src/scan/scan.json');facts={b['block_id']:b for b in scan['blocks']};decisions=[]
    for row in candidate['blocks']:
        idx=row['source_locator']['body_child_index'];desired=blueprint.get(idx)
        if desired and (row['role']!=desired or row['status']=='review'):
            decisions.append(dict(block_id=row['block_id'],role=desired,level=None,origin='model',decision_id='synthetic_'+row['block_id'],reason='Explicit authored list fixture; scripted test decision, no model invocation.',evidence_ids=[row['block_id']]))
        elif row['status']=='review' and row['role'] in ('figure_caption','table_caption'):
            before=[r for r in candidate['blocks'] if r['source_locator']['body_child_index']<idx and r['role'] in ('figure','table')]
            target=before[-1]['block_id'] if before else None
            if target:decisions.append(dict(block_id=row['block_id'],role=row['role'],level=None,relates_to=target,origin='model',decision_id='synthetic_'+row['block_id'],reason='Explicit synthetic physical object association.',evidence_ids=[row['block_id'],target]))
    if decisions:
        s=read(run/'review_requests.json')['sources'][0];response=folder/'response.json';write(response,{'sources':[dict(source_id='src',source_sha256=s['source_sha256'],candidate_sha256=s['candidate_sha256'],decisions=decisions)]})
        cp=subprocess.run([sys.executable,str(ROOT/'report_tool.py'),'review',str(run),str(response)],capture_output=True,text=True)
        if cp.returncode:raise RuntimeError(cp.stdout+cp.stderr)
        state=read(run/'state.json')
    state=workflow.finish(run,state,retain=True);status=workflow.status(run,state)
    docs=[]
    for item in state['deliverables']:
        p=run/item['file'];package=Package(p);numbered=labels(package)
        docs.append(dict(id=item['id'],file=str(p),checks=item['checks'],labels=[{'text':''.join(v.xpath('.//w:t/text()',namespaces={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})),**f} for v,f in numbered.items() if f['status']!='none']))
    return dict(status=status,documents=docs,scripted_decisions=len(decisions))


def main():
    parser=argparse.ArgumentParser();parser.add_argument('--out',type=Path,required=True);args=parser.parse_args();out=args.out.resolve();out.mkdir(parents=True,exist_ok=False)
    src,blueprint=source(out);results={}
    for name,template in [('default',ROOT/'templates/target_template.docx'),('cover',ROOT/'templates/cover_template_v0311.docx')]:
        folder=out/name;folder.mkdir();results[name]=run(folder,src,blueprint,template)
    write(out/'RESULTS.json',results)
    print(json.dumps({k:{'status':v['status']['status'],'documents':len(v['documents']),'review':v['scripted_decisions'],'remaining':v['status']['remaining']} for k,v in results.items()},ensure_ascii=False))
    return int(any(v['status']['status'] not in ('COMPLETED_LOCAL_CHECKS','COMPLETED_WITH_RETAINED_ITEMS') or any(d['checks']['assembly']!='PASS' or d['checks']['content']!='PASS' for d in v['documents']) for v in results.values()))

if __name__=='__main__':raise SystemExit(main())
