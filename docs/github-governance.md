# GitHub Governance For Local Codex Work

GitHub is the governance and traceability layer for this project. It is not a data processing platform.

## GitHub Preflight

默认按最低可行层级执行，优先省 token，不主动升级流程。

- L0 问答/判断/计算正确值：不建 Issue，不读全仓，不跑检查，只直接回答。适用于“是什么、为什么、能不能、怎么做、给方案”。
- L1 本地小改：只改文件、跑最小检查，不上 GitHub。适用于少量真实指标验证、README/AGENTS 小修、注释、局部脚本微调。
- L2 正式脚本/指标任务/治理变更：建 Issue，本地验证，必要时 PR。适用于全量真实指标验证、脚本改动、需要留痕的结论，以及改 workflow、模板、强控规则。

`走GitHub` 和 `上GitHub` 是显式升级到 L2 的触发词；用户明确说 `本地-only` 或 `不用 GitHub` 时降级为本地流程。

## Tool Decision Order

| Layer | Operation scope | First choice | Fallback | Do not use |
|---|---|---|---|---|
| A | Local workspace: status, diff, branch, stash, add, commit | `git` | none | connector / `gh` |
| B | Remote Git objects: fetch, prune, push branch, delete remote branch | `git` | `gh api` | connector unless creating a simple remote file commit |
| C | GitHub governance objects: Issue, comment, labels, PR create/query | GitHub connector | `gh` | manual web edits |
| D | CI / Actions status | GitHub connector | `gh run` / `gh pr checks` | `git` |
| E | Automatic PR merge and branch deletion | GitHub Actions `gh pr merge --squash --delete-branch` | manual `gh pr merge` | connector unless full merge/delete support is confirmed |
| F | Sensitive workbook processing | local Python/scripts | none | GitHub Actions / connector / `gh` |

Decision rule: `git` manages local and remote Git objects, the GitHub connector manages traceability objects, `gh` handles GitHub automation and fallback paths, and real Excel data is processed only by local scripts.

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

L2 有脚本、模板、workflow 或治理文档变更且需要持久化时，必须开 PR，不直接改 `main`。

- 分支名使用 `codex/issue-<issue-number>-<short-topic>`。
- PR 必须关联 Issue，通常使用 `Closes #<issue-number>`。
- PR 必须写明本地验证结论和敏感数据检查结论。
- PR 检查通过后默认自动 squash merge。
- 以下情况不自动合并：draft PR、PR 标题含 `WIP`、PR 带 `needs-review` label、用户明确要求保持打开或等待 review。

## Post-Merge Local Cleanup

PR 合并后，Codex 必须把本地工作区收回 `main`，避免 App 继续停留在已合并或远端已删除的 `codex/...` 分支。

默认步骤：

```powershell
git -c core.quotepath=false -c http.sslBackend=schannel -c http.version=HTTP/1.1 fetch origin main
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
