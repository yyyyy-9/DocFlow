"""API-02: HTML command ownership belongs to HRS, not a second bridge adapter."""
from pathlib import Path
from types import SimpleNamespace
import unittest
import tempfile
from unittest.mock import Mock, patch

from hrs_bridge.browser import export, recover


class HTMLAPIOwner(unittest.TestCase):
    def test_bridge_export_delegates_to_the_selected_hrs(self):
        """The bridge passes options to HRS without rebuilding its own HTML workflow."""
        runtime = SimpleNamespace(hrs=Path(__file__).parent, browser='chrome.exe')
        api = SimpleNamespace(__file__=str(runtime.hrs / 'html_cli.py'), export_html=Mock(return_value={'ok': True}))
        with patch('hrs_bridge.browser.call', side_effect=AssertionError('A second HTML driver was used')), \
             patch('hrs_bridge.browser.importlib.import_module', return_value=api):
            self.assertEqual(export(runtime, 'input', 'output', 'original', 'reading'), {'ok': True})
        api.export_html.assert_called_once_with('input', Path('output').resolve(), browser='chrome.exe',
                                                quality='original', source_mode='reading', accept_numbered_headings=False)

    def test_bridge_recovery_delegates_to_the_selected_hrs(self):
        """Archive/hash/image recovery stays in the same HRS-owned entry point."""
        runtime = SimpleNamespace(hrs=Path(__file__).parent, browser='chrome.exe')
        api = SimpleNamespace(__file__=str(runtime.hrs / 'html_cli.py'), recover_html=Mock(return_value={'ok': True}))
        with patch('hrs_bridge.browser.call', side_effect=AssertionError('A second HTML driver was used')), \
             patch('hrs_bridge.browser.importlib.import_module', return_value=api):
            self.assertEqual(recover(runtime, 'input', 'output'), {'ok': True})
        api.recover_html.assert_called_once_with('input', Path('output').resolve(), browser='chrome.exe')

    def test_existing_bridge_output_parent_behavior_is_preserved(self):
        """Delegating must still create a requested new parent and refuse existing output."""
        runtime = SimpleNamespace(hrs=Path(__file__).parent, browser='chrome.exe')
        api = SimpleNamespace(__file__=str(runtime.hrs / 'html_cli.py'), export_html=Mock(return_value={'ok': True}))
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / 'nested/output/report.html'
            with patch('hrs_bridge.browser.importlib.import_module', return_value=api):
                self.assertEqual(export(runtime, 'input', target), {'ok': True})
                self.assertTrue(target.parent.is_dir())
                target.write_bytes(b'existing user output')
                from hrs_bridge.runtime import BridgeError
                with self.assertRaises(BridgeError):
                    export(runtime, 'input', target)
                self.assertEqual(target.read_bytes(), b'existing user output')


if __name__ == '__main__':
    unittest.main()
