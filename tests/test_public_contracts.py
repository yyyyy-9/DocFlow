"""Public distribution ownership, shared roots and runnable entry contracts."""
from pathlib import Path
import json, os, subprocess, sys, tempfile, unittest

ROOT = Path(__file__).resolve().parents[1]


class PublicContracts(unittest.TestCase):
    def test_metadata_privacy_and_component_integrity(self):
        from maintain_manifest import maintain
        from check_publication import check
        self.assertEqual(maintain()['suite'], json.loads((ROOT/'release.json').read_text('utf8'))['version'])
        self.assertEqual(check()['status'], 'PASS')

    def test_cli_and_default_runtime_share_the_same_kernel(self):
        from hrs_bridge.runtime import default_roots, Runtime
        hrs, wps = default_roots()
        self.assertEqual(Path(wps).resolve(), (ROOT/'WPS_Forms_Workflow').resolve())
        self.assertEqual(Path(hrs).resolve(), (ROOT/'HTML_Report_Studio').resolve())
        Runtime(hrs, wps)
        for args, marker in [(['--version'], 'DocFlow'), (['forms', 'doctor'], 'READY')]:
            result = subprocess.run([sys.executable, '-B', str(ROOT/'run.py'), *args], capture_output=True, text=True, encoding='utf8', timeout=30)
            self.assertEqual(result.returncode, 0, result.stderr)
            self.assertIn(marker, result.stdout)

    def test_dispatcher_preserves_callers_relative_paths(self):
        with tempfile.TemporaryDirectory(prefix='docflow-cli-') as name:
            folder = Path(name)
            # A valid relative manifest with an invalid task must reach manifest
            # validation rather than be looked up in a component directory.
            (folder/'job.json').write_text(json.dumps({'schema': 'invalid-public-test', 'output': 'new', 'tasks': []}), 'utf8')
            result = subprocess.run([sys.executable, '-B', str(ROOT/'run.py'), 'bridge', 'run-job', 'job.json'], cwd=folder, capture_output=True, text=True, encoding='utf8', timeout=30)
            self.assertNotEqual(result.returncode, 0)
            self.assertNotIn('No such file', result.stdout + result.stderr)
            self.assertIn('schema', (result.stdout + result.stderr).lower())


if __name__ == '__main__': unittest.main()
