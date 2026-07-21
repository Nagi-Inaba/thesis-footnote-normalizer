# Repository instructions for AI agents

## Purpose

Help a researcher audit thesis footnotes and their bibliography without modifying the manuscript.

## Non-negotiable rules

- Treat the input DOCX as read-only.
- Work from a separately named copy supplied by the user.
- Never invent missing bibliographic facts.
- Never add a new source unless the user separately authorizes research.
- Treat unmatched and ambiguous citations as `review_required`.
- Do not delete an unused bibliography entry automatically.
- Do not describe `ibid.` or `同上` as valid unless the immediately preceding footnote contains exactly the same single identified source.
- Do not commit manuscripts, audit output, personal data, or absolute local paths.

## Required sequence

1. Confirm that the body text is stable enough for an order-sensitive audit.
2. Read the selected policy and bibliography registry.
3. Run `scripts/Invoke-FootnoteAudit.ps1`.
4. Inspect `issues.csv` before interpreting classifications.
5. Ask a human to resolve ambiguous source identity and policy choices.
6. Apply approved changes manually in Word with Track Changes enabled.
7. Run the audit again and compare the reports.

The repository-level README is the user manual. The runtime skill files are thin execution adapters and must not duplicate the full manual.
