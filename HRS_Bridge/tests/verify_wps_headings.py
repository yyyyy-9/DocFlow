"""HEAD-02 native WPS check for test_headings.MIXED_SOURCE only; no full regression."""
import argparse
import hashlib
import json
from pathlib import Path


def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('source',type=Path)
    parser.add_argument('--out',type=Path,required=True)
    args=parser.parse_args();source=args.source.resolve();out=args.out.resolve()
    out.mkdir(parents=True,exist_ok=False)
    before=hashlib.sha256(source.read_bytes()).hexdigest()
    import win32com.client
    app=doc=None
    result={'issue':'HEAD-02','status':'FAIL','scope':'single title/list fixture in actual WPS'}
    try:
        app=win32com.client.DispatchEx('KWPS.Application')
        app.Visible=False;app.DisplayAlerts=0
        doc=app.Documents.Open(str(source),ReadOnly=True)
        paragraphs={str(p.Range.Text).rstrip('\r'):p for p in doc.Paragraphs}
        expected={'标题与列表专项':'','摘要':'1','细节':'1.1','结论':'2',
                  '列表甲':'1.','列表乙':'2.','无序甲':'•','无序乙':'•',
                  '嵌套甲':'1.','嵌套乙':'2.','后续甲':'3.','后续乙':'4.'}
        observed={text:str(paragraphs[text].Range.ListFormat.ListString) for text in expected}
        result['labels']=observed
        if observed!=expected:raise AssertionError(observed)
        for text,level in [('摘要',1),('细节',2),('结论',1)]:
            assert int(paragraphs[text].OutlineLevel)==level
        assert int(paragraphs['标题与列表专项'].OutlineLevel)==10
        for text in ('列表甲','无序甲','嵌套甲','后续甲'):
            assert int(paragraphs[text].OutlineLevel)==10
        # The reviewed sample is saved before adding a temporary continuation.
        doc.SaveAs(str(out/'标题与列表映射验证.docx'),12)
        doc.ExportAsFixedFormat(str(out/'layout.pdf'),17)
        pos=int(doc.Content.End)-1
        doc.Range(pos,pos).InsertAfter('\r后续章节')
        p=doc.Paragraphs.Item(int(doc.Paragraphs.Count))
        # End the inherited list and explicitly choose the existing chapter
        # outline. WPS retains direct list formatting on a style-only change.
        p.Range.ListFormat.RemoveNumbers()
        p.Style=doc.Styles.Item(-2)
        p.Range.ListFormat.ApplyListTemplateWithLevel(
            paragraphs['摘要'].Range.ListFormat.ListTemplate, True, 2, 1, 1)
        assert str(p.Range.ListFormat.ListString)=='3',str(p.Range.ListFormat.ListString)
        result.update(status='PASS',wps_version=str(app.Version),labels=observed,new_chapter='3')
    except Exception as error:
        result['error']=repr(error);raise
    finally:
        if doc is not None:doc.Close(False)
        if app is not None:app.Quit()
        result['source_unchanged']=before==hashlib.sha256(source.read_bytes()).hexdigest()
        (out/'RESULTS.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf8')
        print(json.dumps(result,ensure_ascii=False))


if __name__=='__main__':main()
