# Troubleshooting

## `word/footnotes.xml` is missing

The document may use endnotes, manually typed note text, or no true Word footnotes. Version 1 audits true footnotes only.

## A real citation is reported as unmatched

Add a distinctive literal alias to the correct bibliography row. Avoid aliases that are common words, short initials, or punctuation-only strings.

## One footnote matches several sources

This can be valid when a note cites several works. It can also mean that an alias is too broad. Review the note and aliases manually.

## An `ibid_candidate` looks wrong

The label requires two adjacent footnotes that each match the same single registered source. It does not inspect the accuracy of page numbers and does not prove that the selected style permits `ibid.`.

## The body changed after the audit

Run the audit again. First-use and repeated-use classifications depend on document order.

## PowerShell reports an execution-policy error

Use a process-scoped policy only if the institution permits it:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

Do not change the machine-wide policy solely for this repository.
