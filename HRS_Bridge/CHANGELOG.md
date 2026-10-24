# Bridge changes

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
