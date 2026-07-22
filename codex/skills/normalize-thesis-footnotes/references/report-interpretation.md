# Report interpretation

- `citation_classification`: `first` or `repeat` for a confirmed registered source.
- `adjacent_same_source`: sequence evidence that adjacent citation-bearing footnotes resolve to the same source.
- `ibid_rewrite_candidate`: policy-aware boolean for human review; never a rewrite instruction.
- Contextual shorthand (`ibid.`, `op. cit.`, 「同上」, 「前掲」): always `review_required` and never bibliographic proof.
- First-use short form: warning that the first matched occurrence may omit policy-required information.
- `bibliography-reconciliation.csv`: comparison of registry data with marked text extracted from an optional bibliography DOCX.
- Special `w:type` footnotes: excluded from citation analysis.

Schema v2 issue types distinguish unmatched and multiple matches, unused registry entries, contextual shorthand, first-use short forms, policy or structured-field gaps, bibliography reconciliation findings, and excluded special footnotes.
Read the repository README and CSV headers for the exact names emitted by the installed version.

Every result is screening evidence, not proof of quotation accuracy, page accuracy, bibliographic truth, or institutional compliance.
