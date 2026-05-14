from __future__ import annotations

import py_compile
import sys
from pathlib import Path


def main() -> int:
    root = Path.cwd()
    scripts = sorted(root.glob("scripts/**/*.py"))
    failures: list[tuple[Path, Exception]] = []

    for script in scripts:
        try:
            py_compile.compile(str(script), doraise=True)
        except Exception as exc:  # pragma: no cover - diagnostic path
            failures.append((script, exc))

    if failures:
        print("Python compile failures:", file=sys.stderr)
        for script, exc in failures:
            print(f"- {script}: {exc}", file=sys.stderr)
        return 1

    print(f"Compiled {len(scripts)} Python files.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
