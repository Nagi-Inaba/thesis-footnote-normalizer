# Report interpretation

- `citation_classification`: `first` or `repeat` for a confirmed registered source.
- `adjacent_same_source`: sequence evidence only.
- `ibid_rewrite_candidate`: policy-aware review flag, never a rewrite instruction.
- `ibid.`, `op. cit.`, 「同上」, and 「前掲」: always `review_required` and never proof of identity.
- First-use short form: warning to check policy-required first-use fields.
- `bibliography-reconciliation.csv`: review comparison against marked bibliography-DOCX text.
- Special `w:type` footnotes: excluded from citation analysis.

Schema v2 adds issue types for shorthand, first-use short forms, policy or field gaps, reconciliation, and special-footnote exclusions alongside matching issues.
Use the README and emitted headers for exact installed-version names.

All results are screening evidence, not bibliographic proof.
