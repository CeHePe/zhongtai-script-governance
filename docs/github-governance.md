# GitHub Governance For Local Codex Work

GitHub is the governance and traceability layer for this project. It is not a data processing platform.

## GitHub Preflight

本项目任务默认先走 GitHub。每个需要留痕的项目任务开始时，Codex 必须先创建或复用 Issue，并在聊天开头写明 `GitHub：Issue #<number>`。

允许跳过 GitHub 的情况只有三类：用户明确说 `本地-only`、用户明确说 `不用 GitHub`、任务是无需留痕的快速只读问题。跳过时必须说明原因。

`走GitHub` 和 `上GitHub` 是显式强调词，不是唯一触发词；默认规则仍然是走 GitHub。

## Issue Rules

Issue 标题使用中文，labels 使用英文稳定枚举。

- 指标验证标题：`【指标验证】报表 / 指标 / 期间 / 维度`，默认 labels：`codex-task`、`metric-validation`、`needs-local-data`。
- 脚本治理标题：`【脚本治理】主题`，默认 labels：`codex-task`、`script-change`。
- 流程演练标题：`【演练】主题`，默认 labels：`codex-task`、`script-change`。
- 缺少必要底表或台账时，加 `blocked-missing-ledger`。
- 本地验证完成且 GitHub-safe 报告已回写时，加 `validated`。
- 需要人工 review、暂缓自动合并或等待确认时，加 `needs-review`。

Issue workflow 会自动创建缺失 labels，并按中文标题前缀补齐最低必需 labels。标题无法分类时，workflow 会评论提醒补全任务类型。

## PR Rules

有脚本、模板、workflow 或治理文档变更时，必须开 PR，不直接改 `main`。

- 分支名使用 `codex/issue-<issue-number>-<short-topic>`。
- PR 必须关联 Issue，通常使用 `Closes #<issue-number>`。
- PR 必须写明本地验证结论和敏感数据检查结论。
- PR 检查通过后默认自动 squash merge。
- 以下情况不自动合并：draft PR、PR 标题含 `WIP`、PR 带 `needs-review` label、用户明确要求保持打开或等待 review。

## Post-Merge Local Cleanup

PR 合并后，Codex 必须把本地工作区收回 `main`，避免 App 继续停留在已合并或远端已删除的 `codex/...` 分支。

默认步骤：

```powershell
git -c core.quotepath=false -c http.sslBackend=schannel fetch origin main
git -c core.quotepath=false switch main
git -c core.quotepath=false merge --ff-only origin/main
git -c core.quotepath=false branch -d <merged-codex-branch>
```

如果切回 `main` 会覆盖本地未提交改动，先用 stash 临时保存并在聊天中说明；只使用 `git branch -d` 删除已合并本地分支，不使用强制删除。

## Data Boundary

Keep these files local only:

- Real ledgers, source reports, bottom tables, indicator lists, and raw exports.
- Excel, CSV, TSV, archives, and generated validation outputs.
- Any report that contains unmasked project names or professional-company names.

Commit these files to GitHub:

- Reusable validation and maintenance scripts.
- GitHub Issue and PR templates.
- No-data governance checks.
- Non-sensitive workflow documentation.

## GitHub-Safe Report Rules

- Project dimension: show project code and project-name pinyin only.
- Professional-company dimension: show pinyin only.
- Region, line, space, and other higher dimensions may show their original names.
- Do not paste full indicator-list field descriptions into GitHub.
- Do not attach or upload source workbooks.

## Required Checks

Run the sensitive-file check after explicit staging:

```powershell
python scripts/governance/check_sensitive_files.py --staged
```

Run the workflow-rule check before publishing governance changes:

```powershell
python scripts/governance/check_github_workflow_rules.py --repo-root .
```

Inspect the staged file list before committing. Never use `git add .` in this repository.
