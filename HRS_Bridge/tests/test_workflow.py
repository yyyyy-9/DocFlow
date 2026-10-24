"""A job must finish with an honest status; required recognition cannot disappear."""
import json
from pathlib import Path
from types import SimpleNamespace
import tempfile
import unittest
from unittest.mock import patch
from hrs_bridge.workflow import run


class JobWorkflow(unittest.TestCase):
    def test_review_status_survives_later_successful_steps(self):
        """Successful preview export cannot turn unresolved structure into accepted output."""
        with tempfile.TemporaryDirectory() as temp:
            p=Path(temp);(p/'input.docx').write_bytes(b'input')
            job=p/'job.json';job.write_text(json.dumps({'schema':'hrs-bridge-job-v1','inputs':{'docx':'input.docx'},'output':'output/run01'}))
            with patch('hrs_bridge.workflow.from_docx',return_value={'ok':True,'recognition':{'unresolved':1}}) as convert, \
                 patch('hrs_bridge.workflow.validate',return_value={'ok':True}), \
                 patch('hrs_bridge.workflow.export_html',return_value={'ok':True}):
                receipt=run(SimpleNamespace(hrs=p,wps=p),job)
            self.assertEqual(convert.call_count,1)
            self.assertEqual(receipt['status'],'NEEDS_REVIEW')
            self.assertFalse(receipt['accepted'])
            self.assertTrue((p/'output/run01/WORKFLOW.json').is_file())

    def test_failure_writes_terminal_receipt_without_running_export(self):
        """A conversion failure is preserved with its stage, while input bytes stay unchanged."""
        with tempfile.TemporaryDirectory() as temp:
            p=Path(temp);(p/'input.docx').write_bytes(b'input')
            job=p/'job.json';job.write_text(json.dumps({'schema':'hrs-bridge-job-v1','inputs':{'docx':'input.docx'},'output':'out'}))
            with patch('hrs_bridge.workflow.from_docx',side_effect=ValueError('recognition failed')), \
                 patch('hrs_bridge.workflow.export_html') as export:
                receipt=run(SimpleNamespace(hrs=p,wps=p),job)
            self.assertEqual(receipt['status'],'FAIL')
            self.assertEqual(receipt['stages'][-1]['name'],'word.recognize_and_import')
            self.assertEqual(receipt['stages'][-1]['status'],'FAIL')
            self.assertEqual(export.call_count,0)
            self.assertEqual((p/'input.docx').read_bytes(),b'input')


class RuntimeRoots(unittest.TestCase):
    def test_workspace_uses_the_single_integration_declaration(self):
        """RELEASE-REVIEW-01: defaults must follow the current declaration, not old paths."""
        from hrs_bridge.runtime import default_roots,BridgeError
        with tempfile.TemporaryDirectory() as temp:
            root=Path(temp);bridge=root/'development/bridge';bridge.mkdir(parents=True)
            (root/'PROJECT_STATE.md').write_text('Current state','utf8')
            (root/'integration').mkdir();(root/'hrs').mkdir();(root/'wps').mkdir()
            declaration=root/'integration/release.json'
            declaration.write_text(json.dumps({'projects':{'HTML_Report_Studio':'hrs','WPS_Document_Tool':'wps'}}),'utf8')
            with patch('hrs_bridge.runtime.ROOT',bridge):
                self.assertEqual(default_roots(),(root/'hrs',root/'wps'))
                declaration.write_text('{broken','utf8')
                with self.assertRaisesRegex(BridgeError,'Invalid project roots'):
                    default_roots()
                adjacent=bridge.parent/'HTML_Report_Studio';adjacent.mkdir()
                other=bridge.parent/'WPS_Document_Tool';other.mkdir()
                self.assertEqual(default_roots(),(adjacent,other))

    def test_explicit_roots_do_not_consult_workspace_defaults(self):
        """An independent command with explicit roots remains usable without a workspace."""
        from hrs_bridge.__main__ import main
        with patch('hrs_bridge.__main__.default_roots',side_effect=AssertionError('Defaults must not run')) as defaults, \
             patch('hrs_bridge.__main__.Runtime'),patch('hrs_bridge.blocks.read',return_value={'ok':True}):
            self.assertEqual(main(['--hrs-root','hrs','--wps-root','wps','read','source']),0)
            defaults.assert_not_called()


if __name__=='__main__':unittest.main()
