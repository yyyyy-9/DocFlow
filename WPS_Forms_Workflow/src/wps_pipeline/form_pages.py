"""Optional WPS first-page facts on disposable frozen copies only.

Run in a child process with a deadline. A missing/failed page probe falls back
to a whole source, not to invented first-page boundaries. No model or OCR.
"""
from pathlib import Path
from zipfile import ZipFile
import json,sys,subprocess,os
from .job_support import write,read,sha


def _worker(folder,sources,operation='pages',timeout=90):
    folder=Path(folder)
    folder.mkdir(parents=True,exist_ok=False)
    result=dict(status='NOT_RUN',operation=operation,documents=[],model_api_calls=0,screenshots='NOT_REQUESTED')
    if os.name!='nt':return dict(result,status='UNAVAILABLE',reason='Windows WPS is unavailable')
    request=folder/'request.json';response=folder/'result.json'
    write(request,dict(sources=sources,response=str(response),operation=operation))
    env=os.environ.copy();env['PYTHONPATH']=str(Path(__file__).resolve().parents[1])
    try:
        p=subprocess.run([sys.executable,'-B','-m','wps_pipeline.form_pages',str(request)],env=env,capture_output=True,timeout=timeout)
        (folder/'stdout.log').write_bytes(p.stdout);(folder/'stderr.log').write_bytes(p.stderr)
        if response.exists():return read(response)
        return dict(result,status='FAILED',reason=p.stderr.decode('utf8',errors='replace')[-1500:])
    except subprocess.TimeoutExpired:
        partial=read(response) if response.exists() else result
        return dict(partial,status='TIMEOUT',reason='WPS deadline exceeded; source retained. No global WPS processes killed.')


def probe(folder,sources,mode='off',timeout=90):
    result=dict(status='NOT_RUN',mode=mode,documents=[],model_api_calls=0,screenshots='NOT_REQUESTED')
    if mode=='off':return result
    # One file has one deadline. A slow source cannot erase facts already read
    # from other documents, and no fallback attaches to a user's WPS session.
    for source in sources:
        value=_worker(Path(folder)/'wps_pages'/source['id'],[source],timeout=timeout)
        row=next(iter(value.get('documents',[])),dict(id=source['id'],status=value['status'],reason=value.get('reason')))
        if row.get('status')=='RUNNING':row.update(status=value['status'],reason=value.get('reason'))
        row['worker_status']=value['status'];row['wps_binary_version']=value.get('wps_binary_version');result['documents'].append(row)
    result['status']='PASS' if result['documents'] and all(x['status']=='PASS' and x.get('source_unchanged') for x in result['documents']) else 'PARTIAL'
    return result


def convert_doc(folder,source,dest,timeout=90):
    """Only a disposable .doc copy is opened; the original binary is retained."""
    dest=Path(dest).resolve()
    if dest.exists():raise FileExistsError('转换不得覆盖: '+str(dest))
    dest.parent.mkdir(parents=True,exist_ok=True)
    value=_worker(folder,[dict(id='conversion',path=str(Path(source).resolve()),output=str(dest))],'convert_doc',timeout)
    row=next(iter(value.get('documents',[])),dict(status=value['status'],reason=value.get('reason')))
    if row.get('status')=='RUNNING':row.update(status=value['status'],reason=value.get('reason'))
    row.pop('id',None)
    if row.get('status')=='PASS' and not row.get('source_unchanged'):row.update(status='FAILED',reason='转换输入副本字节改变，未接入转换结果')
    row.update(worker_status=value['status'],method='WPS_SaveAs_DOCX',semantic_equivalence='NOT_PROVEN',wps_binary_version=value.get('wps_binary_version'))
    return row


def child(request):
    request=read(request);result=dict(status='RUNNING',documents=[],model_api_calls=0,screenshots='NOT_REQUESTED',source_saves=0)
    app=None;owned=False
    try:
        import pythoncom,win32com.client,win32process,win32api
        pythoncom.CoInitialize();before=set(win32process.EnumProcesses())
        app=win32com.client.DispatchEx('KWPS.Application')
        try:
            version=win32api.GetFileVersionInfo(str(Path(app.Path)/'wps.exe'),'\\')
            result['wps_binary_version']='.'.join(map(str,(version['FileVersionMS']>>16,version['FileVersionMS']&65535,version['FileVersionLS']>>16,version['FileVersionLS']&65535)))
        except Exception:result['wps_binary_version']='UNAVAILABLE'
        info=app._oleobj_.GetTypeInfo();lib,_=info.GetContainingTypeLib();constants={}
        for i in range(lib.GetTypeInfoCount()):
            if lib.GetTypeInfoType(i)!=pythoncom.TKIND_ENUM:continue
            en=lib.GetTypeInfo(i)
            if en.GetDocumentation(-1)[0].lower() not in ('wdgotoitem','wdgotodirection','wdstatistic','wdsaveformat'):continue
            for j in range(256):
                try:d=en.GetVarDesc(j)
                except Exception:break
                constants[en.GetDocumentation(d[0])[0]]=d[1]
        operation=request.get('operation','pages')
        required=('wdFormatXMLDocument',) if operation=='convert_doc' else ('wdGoToPage','wdGoToAbsolute','wdStatisticPages')
        result['constants']={k:constants[k] for k in required}
        for source in request['sources']:
            doc=None;path=Path(source['path']);digest=sha(path);attributes=win32api.GetFileAttributes(str(path))
            row=dict(id=source['id'],source_sha256=digest,status='RUNNING');result['documents'].append(row)
            try:
                win32api.SetFileAttributes(str(path),attributes|1)
                doc=app.Documents.Open(str(path),False,True,False)
                pid=win32process.GetWindowThreadProcessId(int(doc.ActiveWindow.Hwnd))[1]
                if pid in before:raise ValueError('DispatchEx did not prove an isolated process; document left untouched')
                owned=True;app.Visible=False;app.DisplayAlerts=0;result['owned_pid']=pid
                if not bool(doc.ReadOnly):raise ValueError('WPS did not honor ReadOnly')
                if operation=='convert_doc':
                    destination=Path(source['output'])
                    if destination.exists():raise FileExistsError('转换目标已存在')
                    before_text=str(doc.Content.Text)
                    before_counts={k:int(getattr(doc,k).Count) for k in ('Tables','InlineShapes','Shapes','OMaths')}
                    doc.SaveAs(str(destination),constants['wdFormatXMLDocument'])
                    with ZipFile(destination) as z:
                        for name in ('[Content_Types].xml','word/document.xml'):z.read(name)
                    after_counts={k:int(getattr(doc,k).Count) for k in before_counts}
                    if before_text!=str(doc.Content.Text) or before_counts!=after_counts:
                        raise ValueError('WPS conversion changed observable text/object counts; converted copy retained for review')
                    row.update(status='PASS',output=str(destination),output_sha256=sha(destination),before_counts=before_counts,after_counts=after_counts,
                               text_unchanged_in_open_document=True,method='WPS_SaveAs_DOCX',read_only_input=True)
                else:
                    doc.Repaginate();pages=int(doc.ComputeStatistics(constants['wdStatisticPages']))
                    start=int(doc.GoTo(constants['wdGoToPage'],constants['wdGoToAbsolute'],1).Start)
                    end=int(doc.GoTo(constants['wdGoToPage'],constants['wdGoToAbsolute'],2).Start) if pages>1 else int(doc.Content.End)
                    row.update(status='PASS',page_count=pages,start=start,end=end,text=str(doc.Range(start,end).Text),method='WPS_PHYSICAL_PAGE',read_only=True)
            except Exception as exc:row.update(status='FAILED',reason=str(exc))
            finally:
                if doc is not None and owned:doc.Close(False)
                win32api.SetFileAttributes(str(path),attributes)
                row['source_unchanged']=sha(path)==digest
            write(request['response'],result)
        result['status']='PASS' if all(d['status']=='PASS' and d['source_unchanged'] for d in result['documents']) else 'PARTIAL'
    except Exception as exc:result.update(status='UNAVAILABLE',reason=str(exc))
    finally:
        if app is not None and owned:
            try:app.Quit();result['owned_application_quit']=True
            except Exception as exc:result['quit_error']=str(exc)
        write(request['response'],result)
        try:pythoncom.CoUninitialize()
        except Exception:pass


if __name__=='__main__':child(sys.argv[1])
