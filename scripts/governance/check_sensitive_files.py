from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path


SENSITIVE_SUFFIXES = {
    ".xlsx",
    ".xls",
    ".xlsm",
    ".xlsb",
    ".csv",
    ".tsv",
    ".zip",
    ".7z",
    ".rar",
    ".html",
    ".htm",
    ".pdf",
    ".docx",
    ".pptx",
    ".msg",
    ".eml",
}

SENSITIVE_NAMES = {
    "local_outputs",
    "outputs",
    "reports",
    "__pycache__",
    ".codex",
}

SENSITIVE_KEYWORDS = (
    "指标清单",
    "台账",
    "底表",
    "报表",
    "明细",
    "客户",
    "内部",
    "竞赛",
    "投稿",
    "Competition",
    "competition",
    "submission",
    "client",
    "confidential",
)


def tracked_files() -> list[Path]:
    result = subprocess.run(
        ["git", "ls-files"],
        check=True,
        capture_output=True,
        text=True,
        encoding="utf-8",
    )
    return [Path(line) for line in result.stdout.splitlines() if line.strip()]


def staged_files() -> list[Path]:
    result = subprocess.run(
        ["git", "diff", "--cached", "--name-only", "--diff-filter=ACMR"],
        check=True,
        capture_output=True,
        text=True,
        encoding="utf-8",
    )
    return [Path(line) for line in result.stdout.splitlines() if line.strip()]


def is_sensitive(path: Path) -> bool:
    parts = set(path.parts)
    name = path.name
    suffixes = [suffix.lower() for suffix in path.suffixes]

    if parts & SENSITIVE_NAMES:
        return True
    if any(suffix in SENSITIVE_SUFFIXES for suffix in suffixes):
        return True
    if name.endswith("_report.md"):
        return True
    return any(keyword in str(path) for keyword in SENSITIVE_KEYWORDS)


def main() -> int:
    parser = argparse.ArgumentParser(description="Reject files that must remain local-only.")
    group = parser.add_mutually_exclusive_group()
    group.add_argument("--tracked", action="store_true", help="Check git-tracked files only.")
    group.add_argument("--staged", action="store_true", help="Check staged files only.")
    args = parser.parse_args()

    if args.staged:
        files = staged_files()
        scope = "staged"
    elif args.tracked:
        files = tracked_files()
        scope = "tracked"
    else:
        files = sorted(set(tracked_files()) | set(staged_files()))
        scope = "tracked/staged"

    violations = sorted(str(path) for path in files if is_sensitive(path))

    if violations:
        print("Sensitive files are not allowed in GitHub:", file=sys.stderr)
        for item in violations:
            print(f"- {item}", file=sys.stderr)
        return 1

    print(f"Sensitive file check passed for {len(files)} {scope} files.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
