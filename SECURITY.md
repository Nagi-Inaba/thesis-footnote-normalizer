# Security and manuscript privacy

This repository is designed for unpublished manuscripts.

- Run the audit against a copy of the manuscript.
- Keep manuscripts under `input/`, generated runs under `work/` or `output/`, and do not commit those directories.
- The PowerShell audit does not upload data or call a network service.
- The audit never edits the input DOCX.
- AI review is optional. Before sending footnote text to an AI service, confirm the author's confidentiality and data-processing requirements.
- Report vulnerabilities privately through [GitHub Private Vulnerability Reporting](https://github.com/Nagi-Inaba/thesis-footnote-normalizer/security/advisories/new). Do not attach a real manuscript or real bibliographic data.

The tool detects formatting and consistency candidates. It does not establish that a quotation, page number, author, title, or legal citation is substantively correct.
