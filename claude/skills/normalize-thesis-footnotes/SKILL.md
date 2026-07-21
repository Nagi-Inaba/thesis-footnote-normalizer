---
name: normalize-thesis-footnotes
description: Audit true Microsoft Word footnotes in a copied thesis DOCX without modifying it. Use when a researcher needs to inventory footnotes, distinguish first and later citations, identify immediate-repeat candidates, reconcile registered bibliography items, or prepare a human-reviewed footnote consistency pass after the thesis body is stable.
---

# Normalize Thesis Footnotes

Use the repository PowerShell audit as the evidence source.

1. Require a separately named DOCX copy, citation policy, and bibliography registry.
2. Confirm confidentiality and the AI service's data-handling terms before receiving footnote text.
3. Request only the audit rows and bibliography fields needed for the decision.
4. Run `scripts/Invoke-FootnoteAudit.ps1` from the cloned repository.
5. Read `issues.csv` before interpreting `citations.csv`.
6. Keep unmatched, multiple-match, translated-work, and specialist-source questions for human review.
7. Do not search for sources, invent missing facts, edit the DOCX, or delete bibliography entries.
8. Prepare proposed changes from the repository's `templates/citation-ledger.csv`, including their governing rule, confidence, and approval status.
9. Ask the researcher to apply approved changes with Word Track Changes and rerun the audit.

Read [references/workflow.md](references/workflow.md) and [references/report-interpretation.md](references/report-interpretation.md) before making recommendations.

Report that alias matching and occurrence order do not verify quotation accuracy, page numbers, bibliographic truth, or institutional compliance.
