"""B2: source-bound plans use the current parser/A0 and preserve original bytes."""
import json
from pathlib import Path
import tempfile
import unittest
from hrs_bridge.runtime import BridgeError, sha
from hrs_bridge.blocks import plan
from test_html import runtime


def table(ident='tbl-test', caption='Original'):
    return '```report-table\n' + json.dumps(dict(schema='report-table-v1', id=ident,
        caption=caption, rows=[[dict(text='A', rowspan=1, colspan=1)]])) + '\n```\n'


class BlockPlan(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.folder = Path(self.temp.name)
        (self.folder / 'assets').mkdir()
        self.file = self.folder / 'report.md'
        self.raw = ('\ufeff# 标题 😀\r\n\r\n' + table().replace('\n','\r\n') + '\r\n见 @tbl:tbl-test。\r\n').encode('utf8')
        self.file.write_bytes(self.raw)
        self.runtime = runtime()

    def tearDown(self):
        self.assertEqual(self.file.read_bytes(), self.raw)
        self.temp.cleanup()

    def test_replace_preserves_bom_crlf_and_surrounding_source(self):
        """B2: UTF-16 ranges include emoji; replacement leaves all surrounding bytes exact."""
        result, candidate = plan(self.runtime, self.file, sha(self.raw), table(caption='Changed').replace('\n','\r\n'), object_id='tbl-test')
        self.assertTrue(result['ok'])
        self.assertEqual(candidate, self.raw.replace(b'Original', b'Changed'))

    def test_stale_snapshot_rejected(self):
        """B2: no candidate is produced from obsolete source coordinates."""
        with self.assertRaisesRegex(BridgeError, 'STALE_SNAPSHOT'):
            plan(self.runtime, self.file, '0'*64, table(), object_id='tbl-test')

    def test_duplicate_id_is_ambiguous(self):
        """B2: two objects with the same ID cannot be silently narrowed to the first."""
        self.raw += ('\r\n' + table()).encode('utf8'); self.file.write_bytes(self.raw)
        with self.assertRaisesRegex(BridgeError, 'AMBIGUOUS_ID'):
            plan(self.runtime, self.file, sha(self.raw), table(), object_id='tbl-test')

    def test_breaking_existing_reference_rejected(self):
        """B2: removing a referenced identity must fail the existing contract."""
        with self.assertRaisesRegex(BridgeError, 'MISSING_REFERENCE'):
            plan(self.runtime, self.file, sha(self.raw), table('tbl-other'), object_id='tbl-test')

    def test_create_candidate_by_inserting_existing_schema(self):
        """B2: a new block is validated by the existing parser, with no implicit ID creation."""
        _, candidate = plan(self.runtime, self.file, sha(self.raw), '\r\n'+table('tbl-new'),
                            object_id='tbl-test', operation='insert-after')
        ir=self.runtime.analyze(candidate)
        self.assertEqual(len(ir['objects']['entries']), 2)

    def test_anonymous_block_requires_explicit_snapshot(self):
        """B2: snapshot-local indices never become stable IDs."""
        with self.assertRaisesRegex(BridgeError, 'MISSING_SNAPSHOT_PRECONDITION'):
            plan(self.runtime, self.file, None, '# Replacement', index=0)


    def test_structural_gate_does_not_claim_complete_output_preflight(self):
        from agent_readonly import project
        from hrs_bridge.runtime import source_contract_valid
        from hrs_bridge.workflow import validate
        ir=self.runtime.analyze(self.raw,self.folder/'assets')
        self.assertFalse(project(ir)['safeToCompile'])
        self.assertTrue(source_contract_valid(ir))
        self.assertTrue(validate(self.runtime,self.folder)['ok'])

if __name__=='__main__':unittest.main(verbosity=2)
