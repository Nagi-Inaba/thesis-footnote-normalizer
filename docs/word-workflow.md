# Word review workflow

## Before the audit

1. Finish structural editing or record that the result is provisional.
2. Save and close the manuscript.
3. Create a separately named copy, such as `thesis-footnote-review-01.docx`.
4. Put the copy under a Git-ignored `input/` directory.
5. Complete the citation policy and bibliography registry.

## After the audit

1. Open `issues.csv` and resolve unmatched or multiply matched notes.
2. Review `citations.csv` in footnote order.
3. Confirm the first and later citation form for every source type.
4. Open the copied DOCX in Microsoft Word.
5. Turn on Track Changes.
6. Apply approved changes one footnote at a time.
7. Do not accept changes until the researcher has reviewed them.
8. Run a new audit against the revised copy.
9. Confirm that the intended issue count decreased and that no citation identity changed unexpectedly.

The audit does not change Word fields, note references, styles, hyperlinks, or tracked changes because it never writes to the DOCX.
