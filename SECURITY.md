# Security Policy

This repository is intended to demonstrate script governance, validation automation, and privacy-conscious engineering practices. It must not contain customer data or unpublished internal materials.

## Data Boundary

Do not commit, upload, attach, or paste:

- Customer workbooks, ledgers, source reports, bottom tables, raw exports, or generated validation outputs.
- Indicator lists, field-definition workbooks, or full calculation-logic extracts.
- Customer names, project Chinese names, unmasked professional-company names, account details, screenshots, or row-level business records.
- Internal documents, presentation materials, submission materials, exported web pages, emails, or archives unless they are explicitly sanitized and approved for publication.

GitHub is used only for script governance, pull request history, issue traceability, and no-data checks. Real data processing must stay in the local Codex environment.

## Issue And Pull Request Rules

- Do not include sensitive source data in public issues, pull requests, comments, screenshots, or logs.
- Use filenames or high-level descriptions only when a local file must be referenced.
- Project-level validation reports may show project code and project-name pinyin only.
- Professional-company validation reports may show pinyin only.
- Do not paste complete indicator-list field descriptions or raw calculation logic into GitHub.

## Automation Rules

GitHub Actions must not download, cache, upload, or process customer workbooks or raw data. Actions are limited to no-data governance checks such as Python syntax checks, sensitive-file checks, and template/workflow rule checks.

## Reporting A Security Or Privacy Problem

If you discover sensitive material in this repository, do not open a public issue with the sensitive content. Contact the repository owner privately with a minimal description of the problem and the affected path or pull request number.

## Incident Response

If sensitive material is accidentally committed:

1. Stop using the affected branch or pull request.
2. Remove the file from the working tree and block it in `.gitignore` and the sensitive-file checker.
3. Assess whether Git history, caches, workflow artifacts, or credentials need cleanup.
4. Rotate any exposed credentials or tokens.
5. Record the remediation using sanitized language only.
