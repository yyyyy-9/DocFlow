# Public tests

All inputs are synthetic, project-owned or published generic style/cover examples. tools/create_public_fixture.py creates a small archive from checked-in public_source Markdown/geometry assets. Its nine figures, grouped images, four tables, formulas and repeated logical resource paths exercise real HRS and Word paths.

Core: `python tools/run_tests.py --group core --out .work/core-check`. This checks suite metadata/privacy, real UI Forms stages, and Bridge Forms binding/transaction boundaries. It does not require WPS or browser automation.

All: install requirements-browser.txt, Node + HRS_Bridge/package.json, choose CHROMIUM_PATH/HRS_BRIDGE_NODE and PANDOC_PATH, then `python tools/run_tests.py --group all --out .work/all-check`. The public archive is generated once when missing. Each run needs a new output path. Logs and RESULT.json remain outside source control.

Native WPS editing, physical pages and GUI window operations remain separately requested checks. A successful OOXML or local-browser test does not establish office rendering. Windows sharing-lock cases skip explicitly outside Windows; missing optional browser/math dependencies must be recorded, not reported as full verification.

HRS_Bridge/tests cover source/archive integrity, native objects and math, templates and front matter, heading/list numbering isolation, footnotes/references, group caption rows, bounded layouts, source-bound edits, rejection cases and cumulative Forms drafts. WPS_Forms_Workflow/tests/run_native_numbering.py is an authored document factory used by native batch tests; no review decisions are inferred from an evaluation truth file.
