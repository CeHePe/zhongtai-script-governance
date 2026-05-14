from __future__ import annotations

import hashlib
import re
from typing import Any


_ASCII_RE = re.compile(r"[^A-Za-z0-9]+")


def _fallback_slug(value: str) -> str:
    digest = hashlib.sha1(value.encode("utf-8")).hexdigest()[:10]
    ascii_part = _ASCII_RE.sub("-", value).strip("-").lower()
    return ascii_part or f"cn-{digest}"


def to_pinyin_slug(value: Any) -> str:
    """Convert Chinese names to a GitHub-safe slug when pypinyin is installed."""
    text = "" if value is None else str(value).strip()
    if not text:
        return ""

    try:
        from pypinyin import lazy_pinyin  # type: ignore
    except ImportError:
        return _fallback_slug(text)

    slug = "-".join(lazy_pinyin(text, errors="ignore")).lower()
    return _ASCII_RE.sub("-", slug).strip("-") or _fallback_slug(text)


def mask_project(project_code: Any, project_name: Any) -> str:
    code = "" if project_code is None else str(project_code).strip()
    name = to_pinyin_slug(project_name)
    return " ".join(part for part in (code, name) if part)


def mask_professional_company(company_name: Any) -> str:
    return to_pinyin_slug(company_name)
