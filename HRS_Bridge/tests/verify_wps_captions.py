"""CAP-01: real WPS editing; run explicitly with a Python that has pywin32.

No HRS renderer runs here. InsertCaption and native field refresh must continue
the exported sequences and update references after an insertion/deletion.
"""
import argparse
import hashlib
import json
from pathlib import Path


def main():
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument('source', type=Path)
    p.add_argument('--out', type=Path, required=True)
    args = p.parse_args()
    source = args.source.resolve()
    out = args.out.resolve()
    out.mkdir(parents=True, exist_ok=False)
    before = hashlib.sha256(source.read_bytes()).hexdigest()
    report = {'issue': 'CAP-01', 'status': 'FAIL', 'source': str(source), 'steps': []}
    import win32com.client
    app = doc = None

    def values(label):
        return [str(f.Result.Text).strip() for f in doc.Fields
                if str(f.Code.Text).strip().startswith('SEQ ' + label + ' ')]

    def refresh():
        # A forward REF may precede the changed SEQ. Two ordinary F9 passes
        # refresh the sequence and then any earlier references to that result.
        assert doc.Fields.Update() == 0
        assert doc.Fields.Update() == 0

    def append_caption(label):
        pos = int(doc.Content.End) - 1
        doc.Range(pos, pos).InsertAfter('\r')
        pos = int(doc.Content.End) - 1
        doc.Range(pos, pos).InsertCaption(Label=label, Title=' CAP-01 temporary test', Position=1)
        refresh()
        f = doc.Fields.Item(int(doc.Fields.Count))
        return str(f.Code.Text), str(f.Result.Text).strip()

    try:
        app = win32com.client.DispatchEx('KWPS.Application')
        app.Visible = False
        app.DisplayAlerts = 0
        doc = app.Documents.Open(str(source), ReadOnly=True)
        report['wps_version'] = str(app.Version)
        code, value = append_caption('图')
        assert value == '11', ('New figure should continue at 11', code, value)
        report['steps'].append({'name': 'append_figure', 'status': 'PASS', 'value': value})
        code, value = append_caption('表')
        assert value == '5', ('New table should continue at 5', code, value)
        report['steps'].append({'name': 'append_table', 'status': 'PASS', 'value': value})

        first = next(f for f in doc.Fields if str(f.Code.Text).strip().startswith('SEQ 图 '))
        pos = int(first.Code.Paragraphs.Item(1).Range.Start)
        doc.Range(pos, pos).InsertBefore('\r')
        doc.Range(pos, pos).InsertCaption(Label='图', Title=' CAP-01 inserted before first figure', Position=1)
        refresh()
        assert values('图') == list(map(str, range(1, 13))), values('图')
        refs = [str(f.Result.Text).strip() for f in doc.Fields if str(f.Code.Text).strip().startswith('REF HN_')]
        assert refs == ['6'], refs
        report['steps'].append({'name': 'insert_before_existing_and_refresh_ref', 'status': 'PASS', 'reference': refs})

        added = next(f for f in doc.Fields if str(f.Code.Text).strip().startswith('SEQ 图 '))
        added.Code.Paragraphs.Item(1).Range.Delete()
        refresh()
        assert values('图') == list(map(str, range(1, 12))), values('图')
        refs = [str(f.Result.Text).strip() for f in doc.Fields if str(f.Code.Text).strip().startswith('REF HN_')]
        assert refs == ['5'], refs
        report['steps'].append({'name': 'delete_inserted_and_refresh_ref', 'status': 'PASS', 'reference': refs})
        edited = out / 'edited_native.docx'
        doc.SaveAs(str(edited), 12)
        doc.Close(False)
        doc = app.Documents.Open(str(edited), ReadOnly=True)
        refresh()
        assert values('图') == list(map(str, range(1, 12)))
        assert values('表') == list(map(str, range(1, 6)))
        report['steps'].append({'name': 'save_reopen_retains_fields', 'status': 'PASS'})
        report['status'] = 'PASS'
    except Exception as error:
        report['error'] = repr(error)
        raise
    finally:
        if doc is not None:
            doc.Close(False)
        if app is not None:
            app.Quit()
        report['source_unchanged'] = hashlib.sha256(source.read_bytes()).hexdigest() == before
        (out / 'RESULTS.json').write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf8')
        print(json.dumps(report, ensure_ascii=False))


if __name__ == '__main__':
    main()
