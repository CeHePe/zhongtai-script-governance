# Local Workbook Row Count Workflow Drill

## Purpose

This document records an end-to-end local Codex governance drill. GitHub is used for traceability only; the sensitive workbook stays local.

## Local Sensitive Read

- Local workbook read: described by category only, not by raw filename
- Upload status: not uploaded
- GitHub content policy: no workbook content, metric-row details, field logic text, raw ledger data, or local output files are included

## Count Result

The drill verified that a local workbook can be opened and counted locally while only a sanitized summary is recorded in GitHub.

| Check | Result |
| --- | --- |
| Workbook opened locally | yes |
| Sheet count captured | yes |
| Row count captured | yes |
| Blank-row check captured | yes |
| Sequence-column check captured | yes |

## Permission Escalation Review

Observed or previously reproduced escalation points are recorded without exposing local workbook names or contents.

- Local Git writes may fail if `.git` ACLs are not aligned with the sandbox.
- Sandbox-local GitHub CLI or Git HTTPS may fail when credentials are unavailable to the sandbox.
- GitHub connector is preferred for Issue, comment, PR, and remote commit operations where possible.

## Config Recommendations

- Keep trusted project paths configured locally.
- Keep local workspace write roots configured locally.
- Do not store GitHub tokens or PATs in local config files committed to the repository.
- Prefer the GitHub connector over local CLI tools when credentials are unstable.

## Validation

- Local spreadsheet reads must remain local-only.
- Remote PR file lists must contain only sanitized documentation, scripts, templates, or governance checks.
- Governance Actions must run no-data checks only.
