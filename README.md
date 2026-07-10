# Privacy-First Metric Validation Governance

隐私优先的指标验证治理仓库

This repository demonstrates a privacy-conscious workflow for validating financial and operational reporting metrics with local-only workbooks and GitHub-based script governance.

本仓库展示一种隐私保护优先的指标验证工作流：真实工作簿只在本地处理，GitHub 只承载脚本治理、流程留痕和无数据检查。

## What This Repository Is

## 本仓库是什么

- A reusable script governance repository for report metric validation.
- 一个可复用的报表指标验证脚本治理仓库。
- A portfolio-friendly example of how to separate sensitive data processing from public engineering artifacts.
- 一个适合作为作品集展示的实践样例：将敏感数据处理与公开工程资产分离。
- A lightweight framework for local validation scripts, no-data CI checks, and GitHub Issue/PR traceability.
- 一个轻量框架，用于本地验证脚本、无数据 CI 检查和 GitHub Issue/PR 留痕。

## What This Repository Does Not Contain

## 本仓库不包含什么

- No customer workbooks, ledgers, source reports, bottom tables, raw exports, or generated validation outputs.
- 不包含客户工作簿、台账、源报表、底表、原始导出或生成的验证输出。
- No indicator lists, full field definitions, raw calculation descriptions, or row-level business records.
- 不包含指标清单、完整字段定义、原始计算口径或行级业务记录。
- No unmasked customer names, project Chinese names, account details, screenshots, emails, or internal materials.
- 不包含未脱敏客户名称、项目中文名、账户细节、截图、邮件或内部材料。

For the formal security boundary, see [`SECURITY.md`](SECURITY.md).

正式安全边界见 [`SECURITY.md`](SECURITY.md)。

## Repository Layout

## 仓库结构

```text
scripts/
  validation/      Reusable metric validation scripts
                   可复用指标验证脚本
  governance/      No-data checks for sensitive files and workflow rules
                   敏感文件与治理规则的无数据检查
docs/              Non-sensitive governance notes and workflow documentation
                   非敏感治理说明与流程文档
.github/           Issue templates, PR template, and no-data GitHub Actions
                   Issue 模板、PR 模板和无数据 GitHub Actions
AGENTS.md          Local Codex execution rules
                   本地 Codex 执行规则
SECURITY.md        Public security and privacy policy
                   公开安全与隐私策略
```

## Workflow

## 工作流

1. Interpret the metric definition locally from approved source materials.
   在本地基于已授权来源材料理解指标口径。
2. Run reusable Python validation scripts against local-only workbooks.
   使用可复用 Python 验证脚本处理仅本地保存的工作簿。
3. Keep raw data and generated outputs out of GitHub.
   原始数据和生成输出不进入 GitHub。
4. Commit only scripts, governance checks, templates, and sanitized documentation.
   只提交脚本、治理检查、模板和脱敏文档。
5. Use GitHub Issues and Pull Requests for traceability without exposing source data.
   使用 GitHub Issue 和 Pull Request 留痕，但不暴露源数据。

## Validation Philosophy

## 验证理念

- Definition first: validation starts from the metric definition, not from report differences.
- 口径优先：验证从指标定义出发，而不是从报表差异倒推。
- Local data, public code: sensitive workbooks stay local; reusable logic is versioned.
- 数据本地、代码公开：敏感工作簿留在本地，可复用逻辑纳入版本管理。
- Reuse over one-off scripts: validation scripts are iterated and improved over time.
- 复用优先于一次性脚本：验证脚本持续迭代优化。
- No-data CI: automated checks verify code and governance rules without loading customer data.
- 无数据 CI：自动化检查只验证代码和治理规则，不加载客户数据。

## No-Data Checks

## 无数据检查

These commands are safe for GitHub Actions because they do not require customer workbooks:

以下命令不依赖客户工作簿，适合在 GitHub Actions 中运行：

```powershell
python scripts\governance\compile_python.py
python scripts\governance\check_sensitive_files.py --tracked
python scripts\governance\check_github_workflow_rules.py --repo-root .
```

Before committing, staged files should also be checked:

提交前还应检查暂存区：

```powershell
python scripts\governance\check_sensitive_files.py --staged
```

## Public Portfolio Note

## 公开作品集说明

This repository is designed to show engineering judgment: data boundaries, repeatable validation scripts, governance automation, and privacy-aware collaboration. It intentionally does not prove results by publishing customer data.

本仓库用于展示工程判断力：数据边界、可复用验证脚本、治理自动化和隐私保护协作。它不会通过公开客户数据来证明结果。

## License

## 许可

No license has been declared yet. Do not reuse this repository's code outside the repository until a license is added.

本仓库暂未声明许可证。在添加许可证前，请勿在仓库外复用代码。
