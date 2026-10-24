# Third-party components

DocFlow project-owned code is GPL-3.0-or-later. This does not replace dependency licenses.

| Component | How it is used | License/source |
| --- | --- | --- |
| HRS 2.0.0 | Pinned Git submodule; release ZIPs include its app source and notices | GPL-3.0-or-later; https://github.com/yyyyy-9/HRS |
| Template style tool 1.0.3 | Isolated source component with its original manifest/notices | GPL-3.0-or-later; HRS/template-tool |
| lxml | Python dependency | BSD-3-Clause; dependency distribution notices |
| python-docx | Bridge and synthetic fixture generation | MIT; dependency distribution notices |
| Pillow | Image processing and public synthetic fixtures | MIT-CMU; dependency distribution notices |
| typing_extensions | Python dependency | PSF-2.0; dependency distribution notices |
| Playwright / playwright-core | Optional browser checks and HRS CLI | Apache-2.0; dependency distribution notices |
| pywin32 | Optional native WPS interaction | PSF-style license; dependency distribution notices |
| Pandoc | Optional external formula converter process | GPL-2.0-or-later; https://pandoc.org |

The source/Clean ZIPs do not embed Python, Node, browsers, Pandoc or office applications.
Offline distributions retain dependency wheel license/METADATA files and playwright-core notices.
Browser-side HRS dependencies are documented in HTML_Report_Studio/app/LICENSING.md and RIGHTS.json.

WPS and Microsoft Word are separate proprietary applications and are not distributed here.
This is an independent project with no vendor affiliation.
