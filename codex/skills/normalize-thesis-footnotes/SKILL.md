---
name: normalize-thesis-footnotes
description: Audit true Microsoft Word footnotes and optional bibliography markers in copied DOCX files without modifying them.
---

# Normalize Thesis Footnotes

Use the repository audit as the evidence source.
The full contract, configuration, commands, and output definitions live in the repository `README.md`.

## Boundaries

1. Require separately named DOCX copies and a new output directory.
2. Read the approved policy and bibliography registry.
3. Keep manuscripts read-only and do not use `-Force` without an exact disposable target.
4. Minimize disclosed footnote and bibliography text.
5. Do not search for sources, invent facts, edit DOCX files, or delete bibliography entries.

## Review

Read `issues.csv` first, then `citations.csv`, optional `bibliography-reconciliation.csv`, `report.md`, and `summary.json`.
Treat `citation_classification`, `adjacent_same_source`, and `ibid_rewrite_candidate` as separate fields.
Contextual shorthand and first-use short-form warnings always require human review.

Use [references/workflow.md](references/workflow.md) for gates and [references/report-interpretation.md](references/report-interpretation.md) for field meanings.

Ask the researcher to apply approved changes in Word with Track Changes and rerun the audit.
Record proposed changes in the repository's `templates/citation-ledger.csv` before human approval.
