"""B1: real engine, real archive, no UI clicking or mocked API."""
import json
import os
from pathlib import Path
import tempfile
import unittest

from hrs_bridge.runtime import Runtime, BridgeError, sha, default_roots
from hrs_bridge.browser import recover, export

ROOT = Path(__file__).resolve().parents[1]


def runtime():
    hrs,wps=default_roots()
    return Runtime(Path(os.environ.get('HRS_ROOT', hrs)),Path(os.environ.get('WPS_ROOT', wps)))


class HTMLRoundtrip(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.temp = tempfile.TemporaryDirectory(prefix='bridge-html-test-')
        cls.folder = Path(cls.temp.name)
        cls.runtime = runtime()
        cls.receipt = recover(cls.runtime, ROOT / 'fixtures/hrs_sample.html', cls.folder / 'source')

    @classmethod
    def tearDownClass(cls):
        cls.temp.cleanup()

    def test_restore_matches_fixed_source_and_dedup(self):
        """B1: all logical paths survive dedup; the original Markdown hash is exact."""
        self.assertEqual(self.receipt['markdown_sha256'], 'c5288ef773b5518a7262d1ab1a61601189e01fd2e2c829e1147bea6a054123a5')
        self.assertEqual(len(self.receipt['assets']), 18)
        self.assertEqual(self.receipt['unique_blobs'], 12)

    def test_programmatic_export_and_restore_are_exact(self):
        """B1: call the original export function; preserve every source byte and resource."""
        result = export(self.runtime, self.folder / 'source', self.folder / 'export.html')
        self.assertEqual(result['engine'], 'hrs-programmatic-v1')
        back = recover(self.runtime, self.folder / 'export.html', self.folder / 'restored')
        self.assertEqual(back['assets'], self.receipt['assets'])
        self.assertEqual(back['markdown_sha256'], self.receipt['markdown_sha256'])

    def test_reading_html_is_not_a_recovery_archive(self):
        """B1: reading-only output remains readable but cannot masquerade as source."""
        export(self.runtime, self.folder / 'source', self.folder / 'reading.html', mode='reading')
        with self.assertRaisesRegex(BridgeError, '纯阅读'):
            recover(self.runtime, self.folder / 'reading.html', self.folder / 'no-source')

    def test_corrupted_source_is_rejected_and_no_partial_project_left(self):
        """B1: the existing archive hash gate must reject tampering before publishing output."""
        original = (ROOT / 'fixtures/hrs_sample.html').read_text('utf8')
        changed = original.replace('c5288ef773b5518a7262d1ab1a61601189e01fd2e2c829e1147bea6a054123a5', '0' * 64)
        self.assertNotEqual(original, changed)
        target = self.folder / 'tampered.html'
        target.write_text(changed, encoding='utf8')
        with self.assertRaisesRegex(BridgeError, 'SHA-256'):
            recover(self.runtime, target, self.folder / 'bad')
        self.assertFalse((self.folder / 'bad').exists())


if __name__ == '__main__':
    unittest.main(verbosity=2)
