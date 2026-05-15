# AGENTS.md

Applies to `C:\BMW\03 jks\02 zhongtai`.

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

Use the project GitHub governance workflow by default for project tasks unless the user explicitly says not to use GitHub, asks for local-only work, or the task is a quick read-only question that does not need traceability.

Treat short phrases such as `走GitHub` or `上GitHub` as explicit triggers. The longer equivalent instruction is:

`按本项目 GitHub 治理流程处理：创建/复用 Issue，本地读取敏感 Excel，报告回写 Issue；如需改脚本，用 codex/issue-编号-简述 分支开 PR，敏感文件不得入库。`

Follow these rules:

- Create or reuse a GitHub Issue for task traceability.
- Read sensitive workbooks locally only; never upload Excel, CSV, ledgers, source reports, indicator lists, raw exports, or local outputs.
- Write the local validation result back to the Issue in GitHub-safe form.
- By default, use the Issue/comment trail only. Do not create a new branch or PR unless the user explicitly asks for one or the change genuinely needs review before it lands.
- If scripts or governance docs need to be persisted, prefer a GitHub connector file commit to `main` or the explicitly requested target branch.
- If a PR is created only for traceability, close it after recording the result unless the user asks to keep it open, review it, or merge it.
- If a branch is explicitly needed, use `codex/issue-<issue-number>-<short-topic>`.
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
