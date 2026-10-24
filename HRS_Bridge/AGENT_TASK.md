# Bridge execution contract

Use hrs-bridge-job-v2 manifests. A wps-forms task delegates to the original Forms CLI, delivers the first draft immediately, then uses forms-submit for classification and grouped feedback. forms-review-plan groups pending items in article order without cutting an article.

Validate every new deliverable and all next-round metadata before switching the committed current round. A failed attempt retains the old round, editable declarations and history. Explicit resubmission creates another attempt directory; do not retry the entire chain automatically.

Do not change frozen specs, snapshots, runtime identities or output hashes. Source/template/asset/code mismatches are errors. UI and Bridge controllers keep separate timelines. Existing pause/review semantics remain distinct from wps-forms.

Use README and the root AGENTS.md for public development/testing. Components are produced from this canonical source; generated copies are not independently edited. No model API is implemented in the program.
