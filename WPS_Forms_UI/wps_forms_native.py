"""Owned WPS viewer process. Only opens prebuilt, read-only review copies."""
import json
import os
import sys
import traceback
from pathlib import Path


def main():
    if hasattr(sys.stdin, 'reconfigure'):
        sys.stdin.reconfigure(encoding='utf-8')
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
    app = None
    docs = {}
    pythoncom = None
    for line in sys.stdin:
        try:
            request = json.loads(line)
            if os.name != 'nt':
                raise RuntimeError('原生Word查看需要Windows/WPS')
            if app is None:
                import pythoncom
                import win32com.client
                pythoncom.CoInitialize()
                app = win32com.client.DispatchEx('KWPS.Application')
                app.Visible = True
                try: app.AutomationSecurity = 3
                except Exception: pass
            results = []
            for item in request['documents']:
                path = str(Path(item['path']).resolve())
                # No caller-supplied executable, macro, field update or save.
                if not Path(path).is_file() or Path(path).suffix.lower() != '.docx':
                    raise ValueError('查看副本不存在：' + path)
                key = os.path.normcase(path)
                doc = docs.get(key)
                if doc is not None:
                    try: str(doc.FullName)
                    except Exception: doc = None
                if doc is None:
                    doc = app.Documents.Open(FileName=path, ReadOnly=True, AddToRecentFiles=False)
                    docs[key] = doc
                if not bool(doc.ReadOnly):
                    raise ValueError('WPS未以只读方式打开查看副本；未继续定位或修改。请保存自己的其他工作后核对本机接口。')
                doc.Activate()
                result = dict(path=path, status='OPENED', readonly=bool(doc.ReadOnly))
                bookmark = item.get('bookmark')
                if bookmark:
                    if not doc.Bookmarks.Exists(bookmark):
                        result.update(status='OPENED_WITHOUT_POSITION', message='查看书签未被WPS保留')
                    else:
                        region = doc.Bookmarks.Item(bookmark).Range
                        region.Select()
                        try: app.ActiveWindow.ScrollIntoView(region, True)
                        except Exception: pass
                        result.update(status='POSITIONED', bookmark=bookmark,
                                      selected_text=str(region.Text)[:500],
                                      start=int(region.Start), end=int(region.End))
                results.append(result)
            print(json.dumps(dict(status='DONE', wps_version=str(app.Version), documents=results), ensure_ascii=False), flush=True)
        except Exception as exc:
            print(json.dumps(dict(status='ERROR', message=str(exc), exception=type(exc).__name__,
                                  traceback=traceback.format_exc()), ensure_ascii=False), flush=True)
    # Do not Quit WPS: the user may still be reading its windows.
    docs.clear(); app = None
    if pythoncom is not None:
        pythoncom.CoUninitialize()


if __name__ == '__main__':
    main()
