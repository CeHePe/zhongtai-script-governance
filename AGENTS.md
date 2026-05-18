# AGENTS.md

Applies to `C:\BMW\03 jks\02 zhongtai`.

## GitHub Preflight (Mandatory)

- 默认按最低可行层级执行，优先省 token，不主动升级流程。
- L0 问答/判断/计算正确值：不建 Issue，不读全仓，不跑检查，只直接回答。适用于“是什么、为什么、能不能、怎么做、给方案”。
- L1 本地小改：只改文件、跑最小检查，不上 GitHub。适用于少量真实指标验证、README/AGENTS 小修、注释、局部脚本微调。
- L2 正式脚本/指标任务/治理变更：建 Issue，本地验证，必要时 PR。适用于全量真实指标验证、脚本改动、需要留痕的结论，以及改 workflow、模板、强控规则。
- `走GitHub`、`上GitHub` 是显式升级到 L2 的触发词；用户明确说 `本地-only`、`不用 GitHub` 时降级为本地流程。
- L2 Issue 标题使用中文：`【指标验证】报表 / 指标 / 期间 / 维度`、`【脚本治理】主题`、`【演练】主题`。
- L2 Issue labels 使用英文固定枚举：指标验证默认 `codex-task`、`metric-validation`、`needs-local-data`；脚本或治理变更加 `script-change`；缺底表加 `blocked-missing-ledger`；验证完成加 `validated`；需要人工 review 加 `needs-review`。
- L2 如需持久化脚本、模板、workflow 或治理文档变更，使用 `codex/issue-<issue-number>-<short-topic>` 分支开 PR。PR 检查通过后默认自动合并，除非用户明确要求保持打开、等待 review、或关闭不合并。
- PR 合并后必须做本地收尾：同步远端 `main`、切回 `main`、快进到 `origin/main`，并用 `git branch -d` 清理已合并的本地 `codex/...` 分支；若本地有未提交改动，先 stash 并说明。

## Core Rule

Use the local skill at [SKILL.md](C:/BMW/03%20jks/02%20zhongtai/.codex/skills/zhongtai-metric-validation/SKILL.md) whenever the task involves:
- testing or retesting report metrics
- checking missing ledgers
- explaining why a metric matches or mismatches
- rolling project results up to region, line, space, or professional-company dimensions

## Required Workflow

1. Read `README.txt` first.
2. Read the exact metric row in `JKS_数据中台二期_指标清单.xlsx` before computing anything.
3. Confirm, in this order:
   - required ledgers and source reports
   - monthly vs cumulative logic
   - D-exit exclusion
   - non-assessment exclusion
   - extra high-dimension adjustments
   - whether high dimensions may aggregate from the project report
4. Identify missing ledgers before giving a pass/fail result.
5. Validate project dimension before higher dimensions unless the user explicitly requests an isolated higher-dimension retest. In that case, state the dependency clearly.
6. Keep reusable Python validation scripts in this folder and iterate them instead of rebuilding from scratch.

## Tooling Rules

- Do not rely on PowerShell to directly touch Chinese filenames.
- Use Python to enumerate files, match filenames, and read Excel workbooks.
- Prefer `pandas` and `openpyxl` for all workbook checks.
- When scripting Chinese enum values, use Unicode-safe Python literals and avoid shell-encoding assumptions.
- Do not use `git add .` in this repository. Stage only explicit allowlisted governance, documentation, and script paths.
- Keep real ledgers, source reports, bottom tables, indicator lists, raw exports, and generated validation outputs local-only.
- GitHub is for script governance, Issue/PR traceability, and review history; it is not a data processing platform.

## GitHub Workflow

Use the project GitHub governance workflow only for L2 tasks by default. L0 answers and L1 local small changes should avoid GitHub unless the user explicitly says `走GitHub` or `上GitHub`.

Treat short phrases such as `走GitHub` or `上GitHub` as explicit triggers. The longer equivalent instruction is:

`按本项目 GitHub 治理流程处理：创建/复用 Issue，本地读取敏感 Excel，报告回写 Issue；如需改脚本，用 codex/issue-编号-简述 分支开 PR，敏感文件不得入库。`

Follow these rules:

- Create or reuse a GitHub Issue for task traceability.
- Read sensitive workbooks locally only; never upload Excel, CSV, ledgers, source reports, indicator lists, raw exports, or local outputs.
- Write the local validation result back to the Issue in GitHub-safe form.
- By default, create or reuse a GitHub Issue and use a PR for any persisted script or governance-doc change.
- Use `codex/issue-<issue-number>-<short-topic>` for the PR branch unless the user explicitly requests another branch.
- After the PR is created and required no-data checks pass, automatically merge it unless the user explicitly asks to keep it open, close it without merge, or wait for review.
- After merge, run local cleanup: fetch `origin/main`, switch to `main`, fast-forward to `origin/main`, and delete the merged local `codex/...` branch with `git branch -d`. Use `http.sslBackend=schannel` and `http.version=HTTP/1.1` for fetch/push on Windows if TLS is unstable.
- Prefer the GitHub connector for Issue, comment, branch, remote file commit, and PR operations. For remote branch updates and PR creation, use the GitHub connector first instead of local `git push`; use `gh` or Git HTTPS only as a fallback because sandbox-local GitHub CLI may not read the Windows keyring.
- Do not use `git add .`; stage explicit allowlisted files only and run `python scripts/governance/check_sensitive_files.py --staged` before any local commit.

## Data Rules

- `项目查询.xlsx` is the authoritative helper ledger for project code, level, status, exit date, and region mapping.
- `非考核项目台账.xlsx` is membership-based: present means non-assessment.
- Do not infer business rules from report differences. Use only:
  - the indicator list
  - explicit user clarification
  - validated project-specific exceptions documented in the local skill references

## Known Project Lessons

- Region-level `综管折让金额` and most `截止性收支` checks in `202512` aligned only after excluding all projects with `项目等级=D` and `项目状态=已撤场`, even if the exit date was in `2026`. Treat this as a validated project-specific case, not a universal default.
- `华中` regional `截止性收支` also included `类型=区域` rows from `截止性收支调整台账.xlsx`.
- `安保` minority-interest logic depends on `1.3-S经营收支表-合资公司账面口径` / indicator-library data, not the `金令金匠金颐韵涵` manual ledger.
- `金令金匠` interest must include both `金令` and `金匠` legal entities.
- `安保` interest can match through `威震保安` entities even when `法人与组织关系.xlsx` lacks an explicit `安保` row; report the mapping gap separately.

## Amortization Testing Lesson

- For `带资摊销` / `智能化整改摊销` / `质效提升` testing, first use `1 摊销模板-新 - 2025年四季度v1.5(去重版) - 导入整理 V2` or a similarly named workbook.
- In that workbook, the `B-按项目统计` sheet is the main project-level base table.
- Treat each project's `还原金额` as the default authoritative reconciliation value for the three amortization metrics.
- Prefer validating `发生数 - 计划数 = 还原金额` instead of independently forcing `计划数` and `发生数` to match first.
- If a mismatch needs deeper tracing, follow the target cell formulas backward to the underlying source sheets.
