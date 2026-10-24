"""Optional actual-WPS QA in an isolated COM instance; no user window is selected."""
import argparse
import json
from pathlib import Path


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('source', type=Path)
    parser.add_argument('--pdf', type=Path, required=True)
    parser.add_argument('--refresh-out', type=Path)
    parser.add_argument('--edit-cell',nargs=4,metavar=('TABLE','ROW','COLUMN','TEXT'))
    args = parser.parse_args()
    if args.edit_cell and not args.refresh_out:parser.error('--edit-cell requires --refresh-out; the input is never saved')
    for target in (args.pdf, args.refresh_out):
        if target:
            if target.exists(): raise ValueError('Output exists: ' + str(target))
            target.parent.mkdir(parents=True, exist_ok=True)
    import win32com.client
    app = doc = None
    try:
        app = win32com.client.DispatchEx('KWPS.Application')
        app.Visible = False
        app.DisplayAlerts = 0
        doc = app.Documents.Open(str(args.source.resolve()), ReadOnly=True)
        result = {'wps_version':str(app.Version), 'fields':int(doc.Fields.Count),
                  'bookmarks':int(doc.Bookmarks.Count), 'source':str(args.source.resolve())}
        if args.edit_cell:
            table,row,col=map(int,args.edit_cell[:3]);text=args.edit_cell[3]
            cell=doc.Tables.Item(table).Cell(row,col)
            cell.Range.Text=text
            if str(cell.Range.Text).rstrip('\r\x07')!=text:raise RuntimeError('WPS native cell edit did not match requested text')
            result['edit_cell']={'table':table,'row':row,'column':col,'text':text}
        if args.refresh_out:
            result['fields_update_result'] = doc.Fields.Update()
            doc.SaveAs(str(args.refresh_out.resolve()), 12)
            result['refreshed_docx'] = str(args.refresh_out.resolve())
        result['field_results']=[{'code':str(doc.Fields.Item(i).Code.Text),'result':str(doc.Fields.Item(i).Result.Text)}
                                 for i in range(1,int(doc.Fields.Count)+1)]
        doc.ExportAsFixedFormat(str(args.pdf.resolve()), 17)
        if not args.pdf.is_file() or args.pdf.stat().st_size < 1000:
            raise RuntimeError('WPS produced no readable PDF')
        result.update(status='PASS', pdf=str(args.pdf.resolve()), bytes=args.pdf.stat().st_size)
        print(json.dumps(result, ensure_ascii=False))
    finally:
        try:
            if doc is not None: doc.Close(False)
        finally:
            if app is not None: app.Quit()


if __name__ == '__main__': main()
