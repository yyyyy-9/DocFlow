"""Native equations: legacy roman declarations keep their original scope."""
from pathlib import Path
from types import SimpleNamespace
import os
import unittest
from lxml import etree as E
from hrs_bridge.math_adapter import MathAdapter, normalize_legacy_roman
from hrs_bridge.runtime import BridgeError


class NativeRomanMath(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        exe = os.environ.get('PANDOC_PATH')
        if not exe or not Path(exe).is_file():
            raise unittest.SkipTest('Select the pinned Pandoc with PANDOC_PATH')
        cls.adapter = MathAdapter(SimpleNamespace(pandoc=exe))

    def assert_same_math(self, source, canonical, display):
        actual = self.adapter.omml(source, display)
        expected = self.adapter.omml(canonical, display)
        self.assertEqual(E.tostring(actual, method='c14n'),
                         E.tostring(expected, method='c14n'))

    def test_real_report_geometric_amplitude(self):
        self.assert_same_math(r'A_{\rm geo}=\sqrt{\gamma_m/\gamma_i}.',
                              r'A_{\mathrm{geo}}=\sqrt{\gamma_m/\gamma_i}.', True)

    def test_real_report_radian_superscript(self):
        self.assert_same_math(r'\Phi(\rho,\theta)=\sum_{n,m}a_{nm}^{\rm rad}Z_{nm}(\rho,\theta).',
                              r'\Phi(\rho,\theta)=\sum_{n,m}a_{nm}^{\mathrm{rad}}Z_{nm}(\rho,\theta).', True)

    def test_nested_inline_group_preserves_math_scope(self):
        self.assert_same_math(r'x_{\rm n_{i+1}}+y',
                              r'x_{\mathrm{n_{i+1}}}+y', False)

    def test_unknown_math_still_rejected(self):
        with self.assertRaisesRegex(BridgeError, 'Native math conversion rejected'):
            self.adapter.omml(r'x+\hrsUnknownMathCommand{y}', True)

    def test_cached_equation_returns_independent_tree(self):
        source=r'x_{\rm geo}'
        first=self.adapter.omml(source);first.clear()
        self.assertTrue(len(self.adapter.omml(source)))


class RomanScope(unittest.TestCase):
    def test_nested_braces(self):
        self.assertEqual(normalize_legacy_roman(r'A_{\rm n_{i+1}}+z'),
                         r'A_{\mathrm{n_{i+1}}}+z')

    def test_whitespace_and_escaped_braces(self):
        self.assertEqual(normalize_legacy_roman(r'x_{ \rm \{geo\}}'),
                         r'x_{ \mathrm{\{geo\}}}')

    def test_command_boundary(self):
        source=r'x_{\rmsuffix geo}+\\rm+\mathrm{geo}'
        self.assertEqual(normalize_legacy_roman(source),source)

    def test_text_argument_is_not_reinterpreted(self):
        source=r'\text{literal {\rm geo}}+x_{\rm geo}'
        self.assertEqual(normalize_legacy_roman(source),
                         r'\text{literal {\rm geo}}+x_{\mathrm{geo}}')

    def test_comments_do_not_change_scope(self):
        source='% {\\rm untouched}\nx_{\\rm geo}'
        self.assertEqual(normalize_legacy_roman(source),
                         '% {\\rm untouched}\nx_{\\mathrm{geo}}')

    def test_malformed_or_bare_declaration_is_not_guessed(self):
        for source in [r'x_{\rm geo',r'\rm x+y',r'{x+\rm y}']:
            self.assertEqual(normalize_legacy_roman(source),source)


if __name__ == '__main__':
    unittest.main()
