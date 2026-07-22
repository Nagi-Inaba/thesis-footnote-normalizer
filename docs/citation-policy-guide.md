# Citation policy guide

## Authority order

Use the graduate-school rules, written supervisor decisions, one base style, and documented project exceptions in that order unless the institution says otherwise.

The audit records policy decisions; it does not choose them.

## Minimal and v0.2 policies

The original minimal policy remains accepted for backward compatibility:

```json
{
  "policy_name": "Thesis citation policy",
  "subsequent_citation": "author-short-title-page",
  "consecutive_same_source": "short-form",
  "new_sources": "prohibited_without_separate_authorization"
}
```

This compatibility form does not describe the fields required for every source type.
Configure `source_type_policies` before treating a full audit as policy-complete.

Each source-type rule declares the structured registry fields required for that type, separated by first and later use where the approved style differs.
Use field names from the bibliography registry, not prose descriptions of a desired citation.

```json
{
  "policy_name": "Thesis citation policy",
  "subsequent_citation": "author-short-title-page",
  "consecutive_same_source": "short-form",
  "new_sources": "prohibited_without_separate_authorization",
  "source_type_policies": {
    "book": {
      "first_use_required_fields": ["author", "title", "publisher", "year"],
      "subsequent_use_required_fields": ["author", "short_title"],
      "consecutive_same_source": "ibid"
    }
  }
}
```

Field names and source types must match the implemented example policy and registry headers.
Do not infer missing publication data from a footnote.

## Decisions required before a full audit

- Required fields for the first and later citation of every `source_type`
- Whether `ibid.`, `op. cit.`, 「同上」, or 「前掲」 is permitted
- Page-range and terminal-punctuation rules
- Translated-work rules
- Specialist-source exceptions
- Whether every cited source must occur in the bibliography
- Whether uncited background sources may remain in the bibliography

## Contextual shorthand

`ibid.`, `op. cit.`, 「同上」, and 「前掲」 depend on context.
The audit always marks their candidates `review_required`; their presence never proves a bibliographic identity or a correct rewrite.

When no controlling rule requires contextual shorthand, prefer an explicit short form such as author, short title, and page.
The audit warns when a short form appears on the first matched use because first-use completeness still needs human review.

## Bibliography sources

`bibliography.csv` remains the identity registry.
Generated `citation-variants.csv` records normalized citation variants and comparison evidence without turning a variant into proof of identity.

An optional bibliography DOCX can be supplied with `-BibliographyDocx`.
The audit extracts text only from markers identified by the policy's `bibliography_document` section and writes `bibliography-reconciliation.csv` for review.
Extraction does not silently replace registry data.

The `bibliography_document` object uses `enabled`, `start_marker`, `end_marker`, `include_heading`, and `paragraph_match_mode`.
When it is enabled without `-BibliographyDocx`, the manuscript copy itself is inspected; supplying the parameter while the policy block is absent or disabled is an error.

Translated books, archival materials, scripture, church documents, canon-law materials, and historical editions often need distinct `source_type_policies`.
Create explicit rules instead of forcing them into a general book rule.
