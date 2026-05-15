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

This workflow is the default for project tasks unless the user explicitly says not to use GitHub, asks for local-only work, or asks a quick read-only question that does not need traceability.

1. The user gives the task in local Codex App using natural language.
2. Codex creates or reuses a GitHub Issue for traceability.
3. Codex reads local-only workbooks and optimizes existing scripts.
4. Codex runs validation locally and reports the full result in chat.
5. Codex posts a GitHub-safe report to the Issue.
6. If scripts changed, Codex opens a PR with only script and governance changes.

## Standard Trigger

The workflow is default. Say `走GitHub` or `上GitHub` in any Codex thread when you want to make that default explicit.

The longer equivalent instruction is:

```text
按本项目 GitHub 治理流程处理：创建/复用 Issue，本地读取敏感 Excel，报告回写 Issue；如需改脚本，用 codex/issue-编号-简述 分支开 PR，敏感文件不得入库。
```

This means:

- Create or reuse a GitHub Issue before treating the task as tracked work.
- Keep all real workbooks and local outputs outside GitHub.
- Report results to the Issue using GitHub-safe summaries.
- Create or reuse a GitHub Issue and use a PR for any persisted script or governance-doc change.
- Use `codex/issue-<issue-number>-<short-topic>` for the PR branch unless the user explicitly requests another branch.
- After the PR is created and required no-data checks pass, automatically merge it unless the user explicitly asks to keep it open, close it without merge, or wait for review.
- Prefer the GitHub connector for remote Issue, comment, branch, file commit, and PR work.
- For remote branch updates and PR creation, use the GitHub connector first instead of local `git push`.
- Use `gh` only as a fallback when the connector cannot perform the operation.

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
