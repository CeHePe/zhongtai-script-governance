# Issue #1 Indicator Count Workflow Drill

## Purpose

This document records an end-to-end local Codex governance drill. GitHub is used for traceability only; the sensitive workbook stays local.

## Local Sensitive Read

- Local workbook read by filename only: `JKS_数据中台二期_指标清单.xlsx`
- Upload status: not uploaded
- GitHub content policy: no workbook content, metric-row details, field logic text, raw ledger data, or local output files are included

## Count Result

| Check | Result |
| --- | ---: |
| Workbook opened locally | yes |
| Sheet count | 1 |
| Sheet max row | 2143 |
| Sheet max column | 15 |
| User-requested count: rows excluding header | 2142 |
| Non-empty rows after header | 2142 |
| Blank rows after header | 0 |
| Sequence-column non-empty rows | 2131 |
| Sequence range | 1..2131 |
| Sequence duplicates | 0 |
| Sequence missing count | 0 |
| Metric-name non-empty rows | 2108 |

Primary answer for the drill task: the indicator list has **2142 rows after the header** under the requested ideal row-counting definition.

## Permission Escalation Review

Observed or previously reproduced escalation points:

- Local Git writes to `.git/HEAD.lock` and `.git/index.lock` required elevation. Current evidence points to `.git` ACL behavior rather than business workflow logic.
- Sandbox-local `gh` and Git HTTPS could not read GitHub keyring credentials, so GitHub CLI/API operations required elevation.
- `gh repo create` previously failed because the token lacked `createRepository`; the repository was ultimately created and pushed outside that blocked path.
- This drill uses the GitHub connector for Issue, remote branch, remote commit, and PR operations where possible.

Current policy decision: do not fix `.git` ACL in this drill. Record the blocker and revisit only if local Git writes must run fully inside the sandbox.

## Config Recommendations

Recommended follow-up for `C:\Users\scene\.codex\config.toml`:

- Keep the existing trusted project entry for `\\?\C:\BMW\03 jks\02 zhongtai`.
- Add the normalized trusted project path `C:\BMW\03 jks\02 zhongtai`.
- Persist `C:\BMW\03 jks\02 zhongtai` and `C:\tmp` under `sandbox_workspace_write.writable_roots`.
- Do not store GitHub tokens or PATs in `config.toml`.
- Prefer the GitHub connector over `gh` for future Issue, comment, PR, and remote commit operations.

## Validation

- Local spreadsheet read used `openpyxl` in read-only mode.
- Remote PR file list should contain only this document.
- Governance Actions should run on the PR and perform no-data checks only.
