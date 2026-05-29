from __future__ import annotations

from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from io import BytesIO
from pathlib import Path
from typing import Iterable
from xml.etree.ElementTree import iterparse
from zipfile import ZipFile

import re
import sys

from _project_root import find_project_root


ROOT = find_project_root(__file__)
NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
TOL = Decimal("0.000001")


def u(text: str) -> str:
    return text.encode("utf-8").decode("unicode_escape")


@dataclass(frozen=True)
class RevenueCheck:
    code: str
    report_amount: Decimal
    source_amount: Decimal
    diff: Decimal
    report_name: str
    source_name: str


def find_workbook(*tokens: str) -> Path:
    matches: list[Path] = []
    for path in ROOT.glob("*.xlsx"):
        name = path.name
        if all(token in name for token in tokens):
            matches.append(path)
    if not matches:
        raise FileNotFoundError(f"Workbook not found for tokens: {tokens!r}")
    if len(matches) > 1:
        names = ", ".join(path.name for path in matches)
        raise FileExistsError(f"Multiple workbooks found for tokens {tokens!r}: {names}")
    return matches[0]


def column_number(cell_ref: str) -> int:
    match = re.match(r"([A-Z]+)", cell_ref)
    if not match:
        raise ValueError(f"Invalid cell reference: {cell_ref!r}")
    number = 0
    for char in match.group(1):
        number = number * 26 + ord(char) - 64
    return number


def shared_strings(zip_file: ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in zip_file.namelist():
        return []
    strings: list[str] = []
    data = BytesIO(zip_file.read("xl/sharedStrings.xml"))
    for _, elem in iterparse(data, events=("end",)):
        if elem.tag == NS + "si":
            strings.append("".join(text.text or "" for text in elem.iter(NS + "t")))
            elem.clear()
    return strings


def read_first_sheet(path: Path) -> list[dict[int, str | None]]:
    rows: list[dict[int, str | None]] = []
    with ZipFile(path) as zip_file:
        strings = shared_strings(zip_file)
        data = BytesIO(zip_file.read("xl/worksheets/sheet1.xml"))
        for _, row in iterparse(data, events=("end",)):
            if row.tag != NS + "row":
                continue
            values: dict[int, str | None] = {}
            for cell in row.findall(NS + "c"):
                col = column_number(cell.attrib.get("r", "A1"))
                cell_type = cell.attrib.get("t")
                value_node = cell.find(NS + "v")
                if cell_type == "s" and value_node is not None:
                    value = strings[int(value_node.text or "0")]
                elif cell_type == "inlineStr":
                    value = "".join(text.text or "" for text in cell.iter(NS + "t"))
                elif value_node is not None:
                    value = value_node.text
                else:
                    value = None
                values[col] = value
            rows.append(values)
            row.clear()
    return rows


def to_decimal(value: object) -> Decimal:
    if value in (None, ""):
        return Decimal("0")
    try:
        return Decimal(str(value))
    except InvalidOperation:
        return Decimal("0")


def normalize_code(value: object) -> str:
    code = str(value or "").strip()
    if len(code) > 1 and code[0] in "DP" and code[1].isdigit():
        return code[1:]
    return code


def add_amount(bucket: dict[str, dict[str, object]], code: object, name: object, amount: Decimal) -> None:
    normalized = normalize_code(code)
    item = bucket.setdefault(
        normalized,
        {"amount": Decimal("0"), "names": set(), "codes": set(), "rows": 0},
    )
    item["amount"] = item["amount"] + amount
    item["rows"] = int(item["rows"]) + 1
    if name:
        item["names"].add(str(name))
    if code:
        item["codes"].add(str(code))


def first(values: Iterable[str]) -> str:
    return sorted(values)[0] if values else ""


def main() -> int:
    report_path = find_workbook("8.1_", "202512")
    source_path = find_workbook("1.5.1-", "202512")

    report_rows = read_first_sheet(report_path)
    source_rows = read_first_sheet(source_path)

    report: dict[str, dict[str, object]] = {}
    for row in report_rows[2:]:
        if not row.get(4):
            continue
        add_amount(report, row.get(4), row.get(5), to_decimal(row.get(11)))

    source: dict[str, dict[str, object]] = {}
    for row in source_rows[5:]:
        if not row.get(2):
            continue
        add_amount(source, row.get(2), row.get(3), to_decimal(row.get(12)) / Decimal("10000"))

    checks: list[RevenueCheck] = []
    missing_source: list[RevenueCheck] = []
    source_only: list[RevenueCheck] = []

    for code in sorted(set(report) | set(source)):
        report_item = report.get(code)
        source_item = source.get(code)
        report_amount = report_item["amount"] if report_item else Decimal("0")
        source_amount = source_item["amount"] if source_item else Decimal("0")
        check = RevenueCheck(
            code=code,
            report_amount=report_amount,
            source_amount=source_amount,
            diff=report_amount - source_amount,
            report_name=first(report_item["names"]) if report_item else "",
            source_name=first(source_item["names"]) if source_item else "",
        )
        if report_item and source_item:
            checks.append(check)
        elif report_item and not source_item:
            missing_source.append(check)
        elif source_item and not report_item and source_amount != 0:
            source_only.append(check)

    mismatches = [check for check in checks if abs(check.diff) > TOL]
    missing_nonzero = [check for check in missing_source if abs(check.report_amount) > TOL]

    print(u(r"\u62a5\u8868\u6587\u4ef6:"), report_path.name)
    print(u(r"\u6765\u6e90\u6587\u4ef6:"), source_path.name)
    print(u(r"\u62a5\u8868\u9879\u76ee\u6570:"), len(report))
    print(u(r"\u6765\u6e90\u9879\u76ee\u6570:"), len(source))
    print(u(r"\u5339\u914d\u9879\u76ee\u6570:"), len(checks))
    print(u(r"\u5dee\u5f02\u9879\u76ee\u6570:"), len(mismatches))
    print(u(r"\u62a5\u8868\u6709\u503c\u4f46\u6765\u6e90\u7f3a\u5931\u9879\u76ee\u6570:"), len(missing_nonzero))
    print(u(r"\u62a5\u8868\u8425\u4e1a\u6536\u5165\u5408\u8ba1(\u4e07\u5143):"), sum(item["amount"] for item in report.values()))
    print(
        u(r"\u5339\u914d\u6765\u6e90\u8425\u4e1a\u6536\u5165\u5408\u8ba1(\u4e07\u5143):"),
        sum(source[code]["amount"] for code in report if code in source),
    )
    print(u(r"\u5339\u914d\u5dee\u5f02\u5408\u8ba1(\u4e07\u5143):"), sum(check.diff for check in checks))
    print(u(r"\u6765\u6e90\u6709\u503c\u4f46\u672a\u57288.1\u62a5\u8868\u51fa\u73b0\u7684\u9879\u76ee\u6570:"), len(source_only))

    if mismatches:
        print(u(r"\n\u5dee\u5f02Top10:"))
        for check in sorted(mismatches, key=lambda item: abs(item.diff), reverse=True)[:10]:
            print(check)
        return 1
    if missing_nonzero:
        print(u(r"\n\u62a5\u8868\u6709\u503c\u4f46\u6765\u6e90\u7f3a\u5931Top10:"))
        for check in sorted(missing_nonzero, key=lambda item: abs(item.report_amount), reverse=True)[:10]:
            print(check)
        return 1

    print(u(r"\u7ed3\u8bba: \u901a\u8fc7"))
    return 0


if __name__ == "__main__":
    sys.exit(main())
