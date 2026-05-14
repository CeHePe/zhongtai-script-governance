# GitHub Governance For Local Codex Work

GitHub is the governance and traceability layer for this project. It is not a data processing platform.

## Boundary

Keep these files local only:

- Real ledgers, source reports, bottom tables, indicator lists, and raw exports.
- Excel, CSV, TSV, archives, and generated validation outputs.
- Any report that contains unmasked project names or professional-company names.

Commit these files to GitHub:

- Reusable validation and maintenance scripts.
- GitHub Issue and PR templates.
- No-data governance checks.
- Non-sensitive workflow documentation.

## Local Task Flow

1. The user gives the task in local Codex App using natural language.
2. Codex creates or reuses a GitHub Issue for traceability.
3. Codex reads local-only workbooks and optimizes existing scripts.
4. Codex runs validation locally and reports the full result in chat.
5. Codex posts a GitHub-safe report to the Issue.
6. If scripts changed, Codex opens a PR with only script and governance changes.

## GitHub-Safe Report Rules

- Project dimension: show project code and project-name pinyin only.
- Professional-company dimension: show pinyin only.
- Region, line, space, and other higher dimensions may show their original names.
- Do not paste full indicator-list field descriptions into GitHub.
- Do not attach or upload source workbooks.

## First Checks Before Any Commit

Run the sensitive-file check after explicit staging:

```powershell
python scripts/governance/check_sensitive_files.py --staged
```

Inspect the staged file list before committing. Never use `git add .` in this repository.
