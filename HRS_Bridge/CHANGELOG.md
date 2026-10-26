# Bridge changes

## 0.4.3

- Import native w:hyperlink labels/formatting and safe destinations using the shared WPS hyperlink_ranges facts. Map unique confirmed heading/object bookmarks to existing HRS IDs; preserve unresolved/unsafe/relative target diagnostics with NEEDS_REVIEW instead of discarding targets or rejecting the whole document.
- Emit native Word anchors for HRS internal links and include anchor/docLocation changes in semantic fingerprints. Coalesce adjacent equally formatted text inside linked paragraphs before Markdown emission.
- Keep source documents, existing group identities, numbering rules, source-bound review and original import limits. Scanner range fidelity and existing number-fact audits run independently; the Forms version/protocol remain 0.1.6/0.5.7 with a new factual build identity.
- Add synthetic pure-data regressions to the existing core/bridge test entry; independently exercise both HRS component and canonical compositions.

## 0.4.2

- Wait for the HRS programmatic API after asynchronous application initialization, with a bounded timeout.
- Use project-authored public fixtures; create the footnote precondition explicitly when testing a cover without an existing notes part.

## 0.4.1

- Build and validate the complete next Forms round before committing current-directory, editable-form, deliverable and history fields.
- Preserve the previous usable round after a temporary output/metadata read failure or a bad new-deliverable hash; record the rejected attempt and permit explicit resubmission.
- Cover real Windows file sharing locks for both classification and feedback, late metadata failures and invalid output digests.

## 0.4.0

- Add source-bound wps-forms stages, explicit forms-submit and article-ordered review planning.
- Use HRS document-v2 numbering, optional equation SEQ/REF and the shared math adapter.
- Share one existing Forms kernel for native report_tool and forms_tool routes.

## Earlier compatibility

Group picture rows and child-caption rows remain separate. Template body bookmarks, native numbering, source-bound block plan/apply, footnotes/references and restricted Word round trips remain supported.
