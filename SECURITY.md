# Security Policy

## 中文版

本仓库用于展示脚本治理、验证自动化和隐私保护工程实践。仓库不得包含客户数据或未公开内部材料。

### 数据边界

不得提交、上传、附加或粘贴：

- 客户工作簿、台账、源报表、底表、原始导出或生成的验证输出。
- 指标清单、字段定义工作簿或完整计算口径摘录。
- 客户名称、项目中文名、未脱敏专业公司名称、账户细节、截图或行级业务记录。
- 内部文档、演示材料、提交材料、导出网页、邮件或压缩包，除非已经明确脱敏并获准公开。

GitHub 只用于脚本治理、Pull Request 历史、Issue 留痕和无数据检查。真实数据处理必须留在本地环境。

### Issue 和 Pull Request 规则

- 不要在公开 Issue、Pull Request、评论、截图或日志中包含敏感源数据。
- 必须引用本地文件时，只使用高层描述或脱敏后的文件类别。
- 项目级验证报告只能显示项目编码和项目名称拼音。
- 专业公司级验证报告只能显示拼音。
- 不要把完整字段定义、原始计算口径或行级数据粘贴到 GitHub。

### 自动化规则

GitHub Actions 不得下载、缓存、上传或处理客户工作簿和原始数据。Actions 仅限运行无数据治理检查，例如 Python 语法检查、敏感文件检查和模板/流程规则检查。

### 报告安全或隐私问题

如果发现仓库中存在敏感材料，不要在公开 Issue 中粘贴敏感内容。请通过私下渠道联系仓库所有者，并只提供最小必要描述和受影响路径或 Pull Request 编号。

### 事件响应

如果敏感材料被误提交：

1. 停止使用受影响分支或 Pull Request。
2. 从工作区移除相关文件，并在 `.gitignore` 和敏感文件检查器中增加拦截规则。
3. 评估是否需要清理 Git 历史、缓存、workflow artifacts 或凭据。
4. 轮换任何可能暴露的凭据或 token。
5. 只使用脱敏语言记录修复过程。

## English Version

This repository demonstrates script governance, validation automation, and privacy-conscious engineering practices. It must not contain customer data or unpublished internal materials.

### Data Boundary

Do not commit, upload, attach, or paste:

- Customer workbooks, ledgers, source reports, bottom tables, raw exports, or generated validation outputs.
- Indicator lists, field-definition workbooks, or full calculation-logic extracts.
- Customer names, project Chinese names, unmasked professional-company names, account details, screenshots, or row-level business records.
- Internal documents, presentation materials, submission materials, exported web pages, emails, or archives unless they are explicitly sanitized and approved for publication.

GitHub is used only for script governance, pull request history, issue traceability, and no-data checks. Real data processing must stay in the local environment.

### Issue And Pull Request Rules

- Do not include sensitive source data in public issues, pull requests, comments, screenshots, or logs.
- Use high-level descriptions or sanitized file categories only when a local file must be referenced.
- Project-level validation reports may show project code and project-name pinyin only.
- Professional-company validation reports may show pinyin only.
- Do not paste complete field definitions, raw calculation logic, or row-level data into GitHub.

### Automation Rules

GitHub Actions must not download, cache, upload, or process customer workbooks or raw data. Actions are limited to no-data governance checks such as Python syntax checks, sensitive-file checks, and template/workflow rule checks.

### Reporting A Security Or Privacy Problem

If you discover sensitive material in this repository, do not open a public issue with the sensitive content. Contact the repository owner privately with a minimal description of the problem and the affected path or pull request number.

### Incident Response

If sensitive material is accidentally committed:

1. Stop using the affected branch or pull request.
2. Remove the file from the working tree and block it in `.gitignore` and the sensitive-file checker.
3. Assess whether Git history, caches, workflow artifacts, or credentials need cleanup.
4. Rotate any exposed credentials or tokens.
5. Record the remediation using sanitized language only.
