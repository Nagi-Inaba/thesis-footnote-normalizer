# Troubleshooting

## `word/footnotes.xml` is missing

The document may use endnotes, manually typed notes, or no true Word footnotes.
The audit processes true Word footnotes only.

## A visible footnote is excluded

Inspect its OOXML `w:type`.
Special footnotes such as separators are deliberately excluded from citation analysis; they are not evidence of a missing citation.

## A real citation is unmatched

Add a distinctive alias to the correct bibliography row.
Avoid common words, short initials, and punctuation-only aliases.

## One footnote matches several sources

The note may validly cite several works, or an alias may be too broad.
Review the note and registry manually.

## A shorthand candidate looks wrong

`adjacent_same_source` reports sequence only.
`ibid_rewrite_candidate` is a policy-aware review flag, not permission to rewrite.
Existing `ibid.`, `op. cit.`, 「同上」, and 「前掲」 remain `review_required` and never prove source identity.

## A first citation has a short-form warning

The first matched use appears to use a registered short form.
Check the required first-use fields in the applicable `source_type_policies`; do not treat the match as proof that the full citation is complete.

## Bibliography reconciliation is empty or incomplete

Confirm that the policy's `bibliography_document` block is enabled and its markers match the selected document structure. When `-BibliographyDocx` is omitted, the audit searches the input DOCX itself.
The audit extracts only marked bibliography content and does not guess where the bibliography begins or ends.
`summary.json` records `marker_not_found` when an enabled marker is not found and `not_observed` when bibliography reconciliation is not enabled.

## A minimal policy passes but type checks are absent

This is backward-compatible behavior.
Add `source_type_policies` with required structured registry fields before relying on source-type completeness checks.

## The body changed after the audit

Run the audit again because `first` and `repeat` depend on document order.

## PowerShell reports an execution-policy error

If institutional policy permits it, use process scope only:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

Do not change machine-wide policy solely for this repository.
