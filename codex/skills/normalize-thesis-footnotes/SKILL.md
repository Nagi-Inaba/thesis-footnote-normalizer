---
name: normalize-thesis-footnotes
description: Audit true Microsoft Word footnotes in a copied thesis DOCX without modifying it. Use when a researcher needs to inventory footnotes, distinguish first and later citations, identify immediate-repeat candidates, reconcile registered bibliography items, or prepare a human-reviewed footnote consistency pass after the thesis body is stable.
---

# Normalize Thesis Footnotes

Use the repository audit script as the evidence source and keep the manuscript read-only.

## Establish the boundary

1. Confirm that the input is a separately named `.docx` copy.
2. Confirm whether structural editing is complete. If it is not, label first/repeat results provisional.
3. Read the approved citation policy and bibliography registry.
4. Confirm the author's confidentiality requirements and the AI service's data-handling terms before sharing any footnote text.
5. Minimize disclosure to the audit rows and bibliography fields needed for the current decision.
6. Treat missing facts, ambiguous identity, and specialist source forms as human decisions.
7. Do not search for or add sources unless the user separately authorizes research.

## Run the audit

Locate the cloned `thesis-footnote-normalizer` repository and run:

```powershell
.\scripts\Invoke-FootnoteAudit.ps1 `
  -InputDocx <copied-docx> `
  -BibliographyCsv <bibliography-csv> `
  -PolicyJson <policy-json> `
  -OutputDirectory <new-output-directory>
```

Do not use `-Force` unless the user has identified the exact disposable output directory.

## Review in this order

1. Read `issues.csv`.
2. Resolve unmatched and multiply matched notes before relying on occurrence classifications.
3. Read `citations.csv` in footnote order.
4. Compare `first`, `repeat`, and `ibid_candidate` with the approved policy.
5. Keep unused bibliography entries as review items, not deletion instructions.
6. Report the before/after input hashes and whether they match.

Read [references/workflow.md](references/workflow.md) for the decision gates and [references/report-interpretation.md](references/report-interpretation.md) before making recommendations.

## Apply changes

Do not edit the DOCX with this skill.

Prepare the proposed-change ledger from the repository's `templates/citation-ledger.csv`. Record footnote number, original text, proposed text, governing rule, confidence, and approval status. Ask the researcher to apply approved changes in Word with Track Changes enabled, then audit the revised copy again.

## Report limitations

State that the audit matches registered aliases and occurrence order. It does not verify quotations, page numbers, bibliographic truth, theological authority, canon-law form, or compliance with an institution's rules.
