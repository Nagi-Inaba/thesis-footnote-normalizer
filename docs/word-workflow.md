# Word review workflow

## Before the audit

1. Finish structural editing or label order-dependent results provisional.
2. Save and close the manuscript.
3. Create a separately named DOCX copy under a Git-ignored input directory.
4. Complete the policy and bibliography registry.
5. If reconciling a Word bibliography, prepare a separate bibliography DOCX and configure `bibliography_document` markers.

## Run without mutation

```powershell
.\scripts\Invoke-FootnoteAudit.ps1 `
  -InputDocx .\input\thesis-review.docx `
  -BibliographyCsv .\input\bibliography.csv `
  -PolicyJson .\input\citation-policy.json `
  -OutputDirectory .\work\audit-001
```

With an optional bibliography document:

```powershell
.\scripts\Invoke-FootnoteAudit.ps1 `
  -InputDocx .\input\thesis-review.docx `
  -BibliographyCsv .\input\bibliography.csv `
  -BibliographyDocx .\input\bibliography-review.docx `
  -PolicyJson .\input\citation-policy.json `
  -OutputDirectory .\work\audit-002
```

The audit reads OOXML and writes reports only.
It does not change Word fields, note references, styles, hyperlinks, tracked changes, or either DOCX.

## Review and edit

1. Resolve `issues.csv` items before relying on classifications.
2. Review `citation_classification`, `adjacent_same_source`, and `ibid_rewrite_candidate` separately.
3. Review every contextual-shorthand and first-use short-form warning.
4. If present, reconcile `bibliography-reconciliation.csv` against the structured registry.
5. Open the copied manuscript in Word and turn on Track Changes.
6. Apply only approved changes, one footnote at a time.
7. Rerun the audit against a new revised copy and compare issue counts.

Footnotes whose OOXML `w:type` marks them as special notes, including separators, are excluded from citation analysis without creating citation or issue rows.
