# Workflow gates

1. **Stable order:** label `first` and `repeat` provisional while notes may move.
2. **Policy authority:** the minimal policy remains compatible, but users should configure `source_type_policies` and required structured fields.
3. **Identity:** require stable `source_id` values; variants and contextual shorthand do not prove identity.
4. **Optional bibliography:** use `-BibliographyDocx` only with explicit `bibliography_document` markers, then review `bibliography-reconciliation.csv`.
5. **Human approval:** treat all shorthand and rewrite candidates as `review_required`.
6. **Rerun:** after tracked Word edits, audit a new copy and compare schema v2 issue counts.

The audit is non-mutating throughout these gates.
