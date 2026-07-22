# Workflow gates

1. Keep copied DOCX inputs read-only and order-dependent results provisional while notes may move.
2. Accept the minimal policy for compatibility, but require configured `source_type_policies` before claiming type-aware completeness.
3. Require stable `source_id` values; variants and shorthand never prove identity.
4. Use `-BibliographyDocx` only with explicit `bibliography_document` markers and review `bibliography-reconciliation.csv`.
5. Leave contextual shorthand and rewrite candidates `review_required` until human approval.
6. After tracked Word edits, audit a new copy and compare schema v2 issues.
