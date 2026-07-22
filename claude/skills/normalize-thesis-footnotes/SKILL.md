---
name: normalize-thesis-footnotes
description: Audit true Microsoft Word footnotes and optional bibliography markers in copied DOCX files without modifying them.
---

# Normalize Thesis Footnotes

Use the repository audit as the evidence source.
The full contract, configuration, commands, and output definitions live in the repository `README.md`.

Require copied DOCX inputs, a new output directory, an approved policy, and a bibliography registry.
Keep manuscripts read-only, minimize disclosure, and never search for sources, invent facts, edit DOCX files, or delete bibliography entries.

Read `issues.csv` first, then `citations.csv`, optional `bibliography-reconciliation.csv`, `report.md`, and `summary.json`.
Treat `citation_classification`, `adjacent_same_source`, and `ibid_rewrite_candidate` separately.
Contextual shorthand and first-use short-form warnings always require human review.

Use [references/workflow.md](references/workflow.md) and [references/report-interpretation.md](references/report-interpretation.md), then ask the researcher to apply approved changes with Word Track Changes and rerun the audit.
Record proposed changes in the repository's `templates/citation-ledger.csv` before human approval.
