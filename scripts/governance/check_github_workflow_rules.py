from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path


REQUIRED_FILES = (
    "AGENTS.md",
    "README.txt",
    "docs/github-governance.md",
    ".github/ISSUE_TEMPLATE/metric-validation.yml",
    ".github/ISSUE_TEMPLATE/governance-task.yml",
    ".github/pull_request_template.md",
    ".github/workflows/governance.yml",
    ".github/workflows/issue-governance.yml",
)

DOCUMENT_SENTINELS = {
    "AGENTS.md": (
        "GitHub Preflight",
        "GitHub Tool Decision Order",
        "L0",
        "L1",
        "L2",
        "走GitHub",
        "上GitHub",
        "GitHub connector",
        "gh pr merge --squash --delete-branch",
        "Sensitive Excel processing",
        "自动合并",
        "PR 合并后",
        "branch -d",
    ),
    "README.txt": (
        "GitHub Preflight",
        "GitHub Tool Decision Order",
        "L0",
        "L1",
        "L2",
        "走GitHub",
        "上GitHub",
        "GitHub connector",
        "gh pr merge --squash --delete-branch",
        "PR 合并后",
    ),
    "docs/github-governance.md": (
        "GitHub Preflight",
        "Tool Decision Order",
        "L0",
        "L1",
        "L2",
        "Issue Rules",
        "PR Rules",
        "Post-Merge Local Cleanup",
        "GitHub connector",
        "gh api",
        "gh pr merge --squash --delete-branch",
        "local Python/scripts",
        "自动合并",
        "http.sslBackend=schannel",
        "http.version=HTTP/1.1",
        "branch -d",
    ),
}

ISSUE_TEMPLATE_SENTINELS = {
    ".github/ISSUE_TEMPLATE/metric-validation.yml": (
        "【指标验证】",
        "codex-task",
        "metric-validation",
        "needs-local-data",
        "不要上传 Excel",
    ),
    ".github/ISSUE_TEMPLATE/governance-task.yml": (
        "【脚本治理】",
        "【演练】",
        "codex-task",
        "script-change",
    ),
}

PR_TEMPLATE_SENTINELS = (
    "## 关联 Issue",
    "Closes #",
    "## 本地验证",
    "## 敏感数据检查",
    "未纳入 Excel",
)

WORKFLOW_SENTINELS = {
    ".github/workflows/governance.yml": (
        "pull_request",
        "check_github_workflow_rules.py",
        "gh pr merge",
        "--squash",
    ),
    ".github/workflows/issue-governance.yml": (
        "issues:",
        "github-script",
        "【指标验证】",
        "addLabels",
    ),
}


def read_text(root: Path, relative_path: str) -> str:
    path = root / relative_path
    if not path.exists():
        raise AssertionError(f"Missing required file: {relative_path}")
    try:
        text = path.read_text(encoding="utf-8")
    except UnicodeDecodeError as exc:
        raise AssertionError(f"{relative_path} is not valid UTF-8: {exc}") from exc
    if "\ufffd" in text:
        raise AssertionError(f"{relative_path} contains replacement characters")
    return text


def require_all(text: str, relative_path: str, sentinels: tuple[str, ...]) -> None:
    missing = [sentinel for sentinel in sentinels if sentinel not in text]
    if missing:
        raise AssertionError(f"{relative_path} is missing required text: {', '.join(missing)}")


def check_repository_rules(root: Path) -> list[str]:
    failures: list[str] = []

    for relative_path in REQUIRED_FILES:
        try:
            read_text(root, relative_path)
        except AssertionError as exc:
            failures.append(str(exc))

    for relative_path, sentinels in DOCUMENT_SENTINELS.items():
        try:
            require_all(read_text(root, relative_path), relative_path, sentinels)
        except AssertionError as exc:
            failures.append(str(exc))

    for relative_path, sentinels in ISSUE_TEMPLATE_SENTINELS.items():
        try:
            require_all(read_text(root, relative_path), relative_path, sentinels)
        except AssertionError as exc:
            failures.append(str(exc))

    try:
        require_all(
            read_text(root, ".github/pull_request_template.md"),
            ".github/pull_request_template.md",
            PR_TEMPLATE_SENTINELS,
        )
    except AssertionError as exc:
        failures.append(str(exc))

    for relative_path, sentinels in WORKFLOW_SENTINELS.items():
        try:
            require_all(read_text(root, relative_path), relative_path, sentinels)
        except AssertionError as exc:
            failures.append(str(exc))

    return failures


def pr_label_names(pr: dict) -> set[str]:
    return {label.get("name", "") for label in pr.get("labels", [])}


def check_pull_request_event(event_path: Path) -> list[str]:
    event = json.loads(event_path.read_text(encoding="utf-8"))
    pr = event.get("pull_request")
    if not pr:
        return []

    failures: list[str] = []
    title = pr.get("title") or ""
    body = pr.get("body") or ""
    labels = pr_label_names(pr)

    if not re.search(r"(?im)\b(closes|fixes|resolves)\s+#\d+\b", body):
        failures.append("PR body must link an Issue with 'Closes #<number>' or equivalent closing keyword.")
    if "## 本地验证" not in body or "## 敏感数据检查" not in body:
        failures.append("PR body must include the local validation and sensitive-data sections from the template.")
    if re.search(r"(?m)- \[ \]", body):
        failures.append("PR body contains unchecked mandatory checklist items.")
    if "needs-review" in labels and "WIP" not in title:
        print("PR has needs-review label; auto-merge will be skipped.")

    return failures


def main() -> int:
    parser = argparse.ArgumentParser(description="Check GitHub workflow guardrails.")
    parser.add_argument("--repo-root", default=".", help="Repository root to inspect.")
    parser.add_argument("--pr-event", help="Optional GitHub pull_request event JSON path.")
    args = parser.parse_args()

    root = Path(args.repo_root).resolve()
    failures = check_repository_rules(root)

    if args.pr_event:
        failures.extend(check_pull_request_event(Path(args.pr_event)))

    if failures:
        print("GitHub workflow rule check failed:", file=sys.stderr)
        for failure in failures:
            print(f"- {failure}", file=sys.stderr)
        return 1

    print("GitHub workflow rule check passed.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
