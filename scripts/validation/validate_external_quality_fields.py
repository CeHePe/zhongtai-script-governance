#!/usr/bin/env python3
"""Validate a contiguous field range in a project operating-quality report.

All customer workbooks are supplied through command-line arguments. The script
does not embed local workbook names or customer/entity exceptions, so it can be
kept as a reusable governance artifact while all sensitive data stays local.
"""

from __future__ import annotations

import argparse
import json
import math
import re
from collections import Counter, defaultdict
from datetime import date, datetime
from pathlib import Path
from typing import Any, Iterable
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


AMOUNT_TOLERANCE = 0.01
RATIO_TOLERANCE = 0.000001
STATUS_PASS = "通过"
STATUS_FAIL = "失败"
STATUS_BLOCKED = "缺源阻塞"
STATUS_REVIEW = "待口径确认"

XML_NS = {
    "m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--report", type=Path, required=True)
    parser.add_argument("--template", type=Path, required=True)
    parser.add_argument("--indicator-list", type=Path, required=True)
    parser.add_argument("--project-master", type=Path, required=True)
    parser.add_argument("--saturation-ledger", type=Path, required=True)
    parser.add_argument("--projection-ledger", type=Path, required=True)
    parser.add_argument("--business-aging", type=Path, required=True)
    parser.add_argument("--receivable-aging", type=Path, required=True)
    parser.add_argument("--half-cash-source", type=Path, required=True)
    parser.add_argument(
        "--cash-comparative-report",
        type=Path,
        help="Optional project report carrying current cumulative cash and MoM.",
    )
    parser.add_argument(
        "--missing-row-policy",
        choices=("blocked", "zero"),
        default="blocked",
        help="How to treat a project that is absent from an otherwise present source.",
    )
    parser.add_argument(
        "--cumulative-policy",
        choices=("indicator-year", "entry-month", "both"),
        default="both",
        help="Validate cumulative cash from January, from entry month, or report both.",
    )
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--summary-json", type=Path)
    return parser.parse_args()


def norm_text(value: Any) -> str:
    return "" if value is None else str(value).strip()


def number(value: Any) -> float:
    if value is None or value == "":
        return 0.0
    try:
        return float(value)
    except (TypeError, ValueError):
        return math.nan


def comparable_number(value: Any) -> float | None:
    if value is None or value == "":
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def near_zero(value: float) -> bool:
    return abs(value) < AMOUNT_TOLERANCE


def safe_ratio(numerator: float, denominator: float) -> float:
    if near_zero(denominator):
        return 0.0
    return numerator / denominator


def compare(actual: Any, expected: Any, kind: str) -> tuple[bool, float | None]:
    if kind == "text":
        return norm_text(actual) == norm_text(expected), None
    if kind == "year":
        actual_text = norm_text(actual)
        expected_text = norm_text(expected)
        return actual_text == expected_text, None
    actual_num = comparable_number(actual)
    expected_num = comparable_number(expected)
    if actual_num is None or expected_num is None:
        return actual_num is None and expected_num is None, None
    diff = actual_num - expected_num
    tolerance = RATIO_TOLERANCE if kind == "ratio" else AMOUNT_TOLERANCE
    return abs(diff) <= tolerance, diff


def read_sheet(path: Path, data_only: bool = True):
    workbook = load_workbook(path, read_only=False, data_only=data_only)
    return workbook.active


def rows_by_key(
    path: Path, start_row: int, key_index: int, data_only: bool = True
) -> tuple[Any, dict[str, list[tuple[Any, ...]]]]:
    sheet = read_sheet(path, data_only=data_only)
    result: dict[str, list[tuple[Any, ...]]] = defaultdict(list)
    for row in sheet.iter_rows(min_row=start_row, values_only=True):
        key = norm_text(row[key_index] if key_index < len(row) else None)
        if key:
            result[key].append(row)
    return sheet, result


def add_months(year: int, month: int, offset: int) -> tuple[int, int]:
    serial = year * 12 + (month - 1) + offset
    return serial // 12, serial % 12 + 1


def parse_date(value: Any) -> date | None:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = norm_text(value)
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y-%m", "%Y/%m"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def normalized_entry_month(value: Any) -> str:
    parsed = parse_date(value)
    if parsed is None:
        return ""
    offset = 1 if parsed.day >= 15 else 0
    year, month = add_months(parsed.year, parsed.month, offset)
    return f"{year:04d}-{month:02d}"


def year_from_period(value: Any) -> str:
    text = norm_text(value)
    match = re.match(r"^(\d{4})", text)
    return match.group(1) if match else ""


def xlsx_cells_without_styles(path: Path) -> dict[str, dict[str, str]]:
    """Read cell values directly from OOXML, bypassing malformed style XML."""
    with ZipFile(path) as archive:
        shared: list[str] = []
        if "xl/sharedStrings.xml" in archive.namelist():
            root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
            shared = [
                "".join(node.text or "" for node in item.iterfind(".//m:t", XML_NS))
                for item in root.findall("m:si", XML_NS)
            ]
        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        relationships = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        relation_map = {node.attrib["Id"]: node.attrib["Target"] for node in relationships}
        result: dict[str, dict[str, str]] = {}
        sheets = workbook.find("m:sheets", XML_NS)
        if sheets is None:
            return result
        for sheet in sheets:
            name = sheet.attrib["name"]
            relation_id = sheet.attrib[f"{{{XML_NS['r']}}}id"]
            target = relation_map[relation_id]
            if not target.startswith("xl/"):
                target = "xl/" + target.lstrip("/")
            sheet_root = ET.fromstring(archive.read(target))
            values: dict[str, str] = {}
            for cell in sheet_root.iterfind(".//m:c", XML_NS):
                reference = cell.attrib["r"]
                cell_type = cell.attrib.get("t")
                value_node = cell.find("m:v", XML_NS)
                value = ""
                if cell_type == "inlineStr":
                    value = "".join(
                        node.text or "" for node in cell.iterfind(".//m:t", XML_NS)
                    )
                elif value_node is not None:
                    raw = value_node.text or ""
                    value = shared[int(raw)] if cell_type == "s" else raw
                if value != "":
                    values[reference] = value
            result[name] = values
        return result


def row_number(reference: str) -> int:
    match = re.search(r"\d+", reference)
    return int(match.group()) if match else 0


def indicator_evidence(path: Path) -> dict[str, bool]:
    workbook = xlsx_cells_without_styles(path)
    all_values = [value for sheet in workbook.values() for value in sheet.values()]
    required = {
        "累计回收现金流": False,
        "累计回收现金流_项目": False,
        "年度饱和收入_打折前": False,
    }
    for name in required:
        required[name] = any(value == name for value in all_values)
    return required


def template_evidence(path: Path) -> dict[str, Any]:
    workbook = load_workbook(path, read_only=False, data_only=True)
    for sheet in workbook.worksheets:
        values = {
            norm_text(cell.value): cell.coordinate
            for row in sheet.iter_rows(min_row=1, max_row=min(sheet.max_row, 5))
            for cell in row
            if cell.value is not None
        }
        if "合资公司名称" in values and "1个月时间" in values:
            return {
                "sheet": sheet.title,
                "start": values["合资公司名称"],
                "one_month_time": values["1个月时间"],
                "has_recovery_rate": any("回款率" in key for key in values),
            }
    return {"sheet": None, "start": None, "one_month_time": None, "has_recovery_rate": False}


def first_or_none(items: list[tuple[Any, ...]]) -> tuple[Any, ...] | None:
    return items[0] if len(items) == 1 else None


def sum_index(items: Iterable[tuple[Any, ...]], index: int) -> float:
    total = 0.0
    for item in items:
        if index < len(item):
            value = number(item[index])
            if not math.isnan(value):
                total += value
    return total


def find_header_row(sheet: Any, required: str, max_rows: int = 10) -> int:
    for row_index in range(1, min(max_rows, sheet.max_row) + 1):
        if any(norm_text(cell.value) == required for cell in sheet[row_index]):
            return row_index
    raise ValueError(f"Required header not found: {required}")


def build_cash_comparative(path: Path | None) -> dict[str, tuple[float, float, float]]:
    if path is None:
        return {}
    sheet = read_sheet(path, data_only=True)
    result: dict[str, tuple[float, float, float]] = {}
    for row in sheet.iter_rows(min_row=3, values_only=True):
        code = norm_text(row[1] if len(row) > 1 else None)
        if not code:
            continue
        cumulative = number(row[10]) * 10000
        mom = number(row[11])
        if math.isnan(cumulative) or math.isnan(mom):
            continue
        if near_zero(cumulative):
            monthly = 0.0
        elif abs(mom) <= RATIO_TOLERANCE:
            monthly = cumulative
        elif abs(1 + mom) <= RATIO_TOLERANCE:
            monthly = math.nan
        else:
            monthly = cumulative - cumulative / (1 + mom)
        result[code] = (cumulative, mom, monthly)
    return result


def output_value(value: Any) -> Any:
    if isinstance(value, float) and math.isnan(value):
        return None
    return value


def main() -> int:
    args = parse_args()
    input_paths = [
        args.report,
        args.template,
        args.indicator_list,
        args.project_master,
        args.saturation_ledger,
        args.projection_ledger,
        args.business_aging,
        args.receivable_aging,
        args.half_cash_source,
    ]
    if args.cash_comparative_report:
        input_paths.append(args.cash_comparative_report)
    missing_files = [str(path) for path in input_paths if not path.exists()]
    if missing_files:
        raise FileNotFoundError("Missing inputs: " + ", ".join(missing_files))

    report_sheet = read_sheet(args.report, data_only=True)
    project_sheet, project_map = rows_by_key(args.project_master, 2, 0)
    saturation_sheet, saturation_map = rows_by_key(args.saturation_ledger, 2, 8)
    projection_sheet, projection_map = rows_by_key(args.projection_ledger, 4, 2)
    business_sheet, business_map = rows_by_key(args.business_aging, 5, 1)
    receivable_sheet, receivable_map = rows_by_key(args.receivable_aging, 3, 4)
    half_cash_sheet, half_cash_map = rows_by_key(args.half_cash_source, 6, 1)
    cash_comparative = build_cash_comparative(args.cash_comparative_report)

    # Force structural checks so a wrong workbook fails loudly.
    find_header_row(project_sheet, "立项编码")
    find_header_row(saturation_sheet, "业绩认定年月")
    find_header_row(projection_sheet, "营业收入")
    find_header_row(business_sheet, "应收余额")
    find_header_row(receivable_sheet, "大业主应收金额 （单位：元）")
    find_header_row(half_cash_sheet, "累计回收现金流")

    template_check = template_evidence(args.template)
    indicator_check = indicator_evidence(args.indicator_list)
    if not template_check["sheet"] or not all(indicator_check.values()):
        raise ValueError(
            f"Template/indicator evidence incomplete: {template_check}, {indicator_check}"
        )

    field_specs = [
        (11, "K", "合资公司名称", "text", "项目主数据"),
        (12, "L", "股权比例", "ratio", "项目主数据"),
        (13, "M", "拓展年份", "year", "年饱和台账"),
        (14, "N", "年饱和", "amount", "年饱和台账"),
        (15, "O", "营业收入", "amount", "投模台账"),
        (16, "P", "半收付收入", "amount", "投模台账"),
        (17, "Q", "权责税前利润", "amount", "投模台账"),
        (18, "R", "半收付税前利润", "amount", "投模台账"),
        (19, "S", "营业成本（含增值税附加）", "amount", "投模台账"),
        (20, "T", "自有成本", "amount", "投模台账"),
        (21, "U", "外包成本", "amount", "投模台账"),
        (22, "V", "能耗成本", "amount", "投模台账"),
        (23, "W", "回款率", "ratio", "报表公式"),
        (24, "X", "投前测算入住率", "ratio", "投模台账"),
        (25, "Y", "增值税费附加", "amount", "投模台账"),
        (26, "Z", "管理费", "amount", "投模台账"),
        (27, "AA", "年终奖励/月", "amount", "投模台账"),
        (28, "AB", "带资投入费用", "amount", "投模台账"),
        (29, "AC", "市场营销费", "amount", "投模台账"),
        (30, "AD", "调整后营业收入", "amount", "报表公式"),
        (31, "AE", "调整后半收付收入", "amount", "报表公式"),
        (32, "AF", "调整后权责税前利润", "amount", "报表公式"),
        (33, "AG", "调整后半收付税前利润", "amount", "报表公式"),
        (34, "AH", "权责税前利润率", "ratio", "报表公式"),
        (35, "AI", "半收付税前利润率", "ratio", "报表公式"),
        (36, "AJ", "营业成本（不含带资）", "amount", "报表公式"),
        (37, "AK", "营业成本", "amount", "报表公式"),
        (38, "AL", "1个月时间", "text", "项目主数据"),
        (39, "AM", "累计回收现金流", "amount", "半收付底表"),
        (40, "AN", "应收账款余额", "amount", "账龄底表"),
        (41, "AO", "未到账期金额", "amount", "账龄底表"),
        (42, "AP", "回款率", "ratio", "报表公式"),
    ]

    details: list[dict[str, Any]] = []
    source_gaps: Counter[str] = Counter()

    def add_result(
        *,
        row_number_value: int,
        project_code: str,
        column: str,
        field: str,
        kind: str,
        source: str,
        actual: Any,
        expected: Any,
        status: str | None = None,
        alternate: Any = None,
        note: str = "",
    ) -> None:
        passed, difference = compare(actual, expected, kind)
        final_status = status or (STATUS_PASS if passed else STATUS_FAIL)
        details.append(
            {
                "报表行": row_number_value,
                "项目编码": project_code,
                "列": column,
                "字段": field,
                "报表值": output_value(actual),
                "应有值": output_value(expected),
                "备选口径值": output_value(alternate),
                "差异": output_value(difference),
                "状态": final_status,
                "来源角色": source,
                "说明": note,
            }
        )

    for excel_row, report_row in enumerate(
        report_sheet.iter_rows(min_row=4, values_only=True), start=4
    ):
        project_code = norm_text(report_row[1] if len(report_row) > 1 else None)
        if not project_code:
            continue

        project_rows = project_map.get(project_code, [])
        saturation_rows = saturation_map.get(project_code, [])
        projection_rows = projection_map.get(project_code, [])
        business_rows = business_map.get(project_code, [])
        receivable_rows = receivable_map.get(project_code, [])
        cash_rows = half_cash_map.get(project_code, [])

        project = first_or_none(project_rows)
        if project is None:
            source_gaps["项目主数据缺行或重复"] += 1
            project_expected = {11: None, 12: None, 38: None}
            project_status = STATUS_BLOCKED
        else:
            project_expected = {
                11: project[7],
                12: number(project[9]),
                38: normalized_entry_month(project[12]),
            }
            project_status = None

        for index, column, field, kind, source in field_specs:
            if index not in project_expected:
                continue
            add_result(
                row_number_value=excel_row,
                project_code=project_code,
                column=column,
                field=field,
                kind=kind,
                source=source,
                actual=report_row[index - 1],
                expected=project_expected[index],
                status=project_status,
                note="项目主数据唯一匹配" if project_status is None else "项目主数据无法唯一匹配",
            )

        saturation_status: str | None = None
        if not saturation_rows:
            source_gaps["年饱和台账项目缺行"] += 1
            saturation_expected = {13: "", 14: 0.0}
            saturation_status = (
                STATUS_BLOCKED if args.missing_row_policy == "blocked" else None
            )
        else:
            years = {year_from_period(row[3]) for row in saturation_rows if row[3]}
            if len(years) != 1:
                source_gaps["年饱和台账年度不唯一"] += 1
                saturation_status = STATUS_BLOCKED
            saturation_expected = {
                13: next(iter(years), ""),
                14: sum_index(saturation_rows, 9),
            }
        for index, column, field, kind, source in field_specs:
            if index not in saturation_expected:
                continue
            add_result(
                row_number_value=excel_row,
                project_code=project_code,
                column=column,
                field=field,
                kind=kind,
                source=source,
                actual=report_row[index - 1],
                expected=saturation_expected[index],
                status=saturation_status,
                note="无匹配行" if not saturation_rows else f"匹配行数={len(saturation_rows)}",
            )

        projection_status: str | None = None
        projection = first_or_none(projection_rows)
        if projection is None:
            reason = "项目缺行" if not projection_rows else "项目重复"
            source_gaps[f"投模台账{reason}"] += 1
            projection_status = STATUS_BLOCKED
            if args.missing_row_policy == "zero" and not projection_rows:
                projection_status = None
            projection_values = [0.0] * 15
        else:
            projection_values = [number(value) for value in projection[3:18]]
        direct_projection = {
            15: projection_values[0],
            16: projection_values[1],
            17: projection_values[2],
            18: projection_values[3],
            19: projection_values[4],
            20: projection_values[5],
            21: projection_values[6],
            22: projection_values[7],
            24: projection_values[9],
            25: projection_values[10],
            26: projection_values[11],
            27: projection_values[12],
            28: projection_values[13],
            29: projection_values[14],
        }
        for index, column, field, kind, source in field_specs:
            if index not in direct_projection:
                continue
            add_result(
                row_number_value=excel_row,
                project_code=project_code,
                column=column,
                field=field,
                kind=kind,
                source=source,
                actual=report_row[index - 1],
                expected=direct_projection[index],
                status=projection_status,
                note="无匹配行" if not projection_rows else f"匹配行数={len(projection_rows)}",
            )

        # Report-layer formulas are validated using the report's own upstream fields.
        revenue = number(report_row[14])
        semi_revenue = number(report_row[15])
        accrual_profit = number(report_row[16])
        semi_profit = number(report_row[17])
        vat = number(report_row[24])
        management_fee = number(report_row[25])
        monthly_bonus = number(report_row[26])
        financed_input = number(report_row[27])
        adjusted_revenue = revenue
        adjusted_semi_revenue = semi_revenue
        adjusted_accrual_profit = accrual_profit + vat + management_fee + monthly_bonus * 12
        adjusted_semi_profit = semi_profit + vat + management_fee + monthly_bonus * 12
        adjusted_cost = semi_revenue - adjusted_semi_profit
        formula_expected = {
            23: safe_ratio(semi_revenue + revenue * 0.06, revenue * 1.06),
            30: adjusted_revenue,
            31: adjusted_semi_revenue,
            32: adjusted_accrual_profit,
            33: adjusted_semi_profit,
            34: safe_ratio(adjusted_accrual_profit, revenue),
            35: safe_ratio(adjusted_semi_profit, semi_revenue),
            36: adjusted_cost - financed_input,
            37: adjusted_cost,
        }
        for index, column, field, kind, source in field_specs:
            if index not in formula_expected:
                continue
            add_result(
                row_number_value=excel_row,
                project_code=project_code,
                column=column,
                field=field,
                kind=kind,
                source=source,
                actual=report_row[index - 1],
                expected=formula_expected[index],
                note="按模板报表层公式，以本报表上游字段复算",
            )

        receivable_balance = sum_index(business_rows, 11) + sum_index(receivable_rows, 8)
        not_due = sum_index(receivable_rows, 14)
        for index, expected, rows_present, label in (
            (40, receivable_balance, bool(business_rows or receivable_rows), "应收账款余额"),
            (41, not_due, bool(receivable_rows), "未到账期金额"),
        ):
            status = None
            if not rows_present:
                source_gaps[f"账龄底表{label}项目缺行"] += 1
                if args.missing_row_policy == "blocked":
                    status = STATUS_BLOCKED
            spec = next(item for item in field_specs if item[0] == index)
            add_result(
                row_number_value=excel_row,
                project_code=project_code,
                column=spec[1],
                field=spec[2],
                kind=spec[3],
                source=spec[4],
                actual=report_row[index - 1],
                expected=expected,
                status=status,
                note=f"业务账龄行={len(business_rows)}；应收账龄行={len(receivable_rows)}",
            )

        cumulative_cash = sum_index(cash_rows, 10)
        cash_status: str | None = None
        cash_note = f"半收付底表匹配行={len(cash_rows)}；指标清单口径=当年1月至报表月累计"
        if not cash_rows:
            source_gaps["半收付底表项目缺行"] += 1
            if args.missing_row_policy == "blocked":
                cash_status = STATUS_BLOCKED
        window_cash = (
            0.0
            if (cash_rows and near_zero(cumulative_cash))
            or (not cash_rows and args.missing_row_policy == "zero")
            else None
        )
        if project_code in cash_comparative:
            comparative_cumulative, mom, comparative_month = cash_comparative[project_code]
            window_cash = comparative_month
            cash_note += (
                f"；备选进场月口径由累计值及环比反推，累计={comparative_cumulative:.2f}，环比={mom:.12g}"
            )
        actual_cash = report_row[38]
        cumulative_match, _ = compare(actual_cash, cumulative_cash, "amount")
        window_match = False
        if window_cash is not None and not math.isnan(window_cash):
            window_match, _ = compare(actual_cash, window_cash, "amount")
        expected_cash = cumulative_cash
        alternate_cash = window_cash
        if args.cumulative_policy == "entry-month":
            expected_cash = window_cash
            alternate_cash = cumulative_cash
            cash_note += "；采用已确认口径=从进场月开始累计"
            if expected_cash is None or math.isnan(expected_cash):
                cash_status = STATUS_BLOCKED
                source_gaps["缺少进场月累计所需的上期累计值"] += 1
            elif cash_status is None and not window_match:
                cash_status = STATUS_FAIL
        elif args.cumulative_policy == "indicator-year":
            cash_note += "；采用指标清单通用口径=当年1月至报表月累计"
            if cash_status is None and not cumulative_match:
                cash_status = STATUS_FAIL
        elif window_cash is not None and not math.isnan(window_cash):
            if abs(window_cash - cumulative_cash) > AMOUNT_TOLERANCE:
                cash_status = STATUS_REVIEW
                source_gaps["累计现金流时间窗口待确认"] += 1
            elif cash_status is None and not cumulative_match:
                cash_status = STATUS_FAIL
        elif cash_status is None and not cumulative_match:
            cash_status = STATUS_FAIL
        add_result(
            row_number_value=excel_row,
            project_code=project_code,
            column="AM",
            field="累计回收现金流",
            kind="amount",
            source="半收付底表",
            actual=actual_cash,
            expected=expected_cash,
            alternate=alternate_cash,
            status=cash_status,
            note=cash_note
            + f"；累计匹配={cumulative_match}；备选匹配={window_match}",
        )

        # Recovery rate arithmetic is independent of the disputed cash source window.
        report_cash = number(report_row[38])
        report_receivable = number(report_row[39])
        report_not_due = number(report_row[40])
        recovery_expected = safe_ratio(
            report_cash, report_cash + report_receivable - report_not_due
        )
        add_result(
            row_number_value=excel_row,
            project_code=project_code,
            column="AP",
            field="回款率",
            kind="ratio",
            source="报表公式",
            actual=report_row[41],
            expected=recovery_expected,
            note="按模板公式，以本报表AM、AN、AO复算；源值结论依赖AM口径确认",
        )

    counts = Counter(item["状态"] for item in details)
    per_field: dict[str, Counter[str]] = defaultdict(Counter)
    for item in details:
        per_field[f"{item['列']} {item['字段']}"][item["状态"]] += 1

    summary = {
        "report_rows": len({item["报表行"] for item in details}),
        "field_count": len(field_specs),
        "detail_count": len(details),
        "status_counts": dict(counts),
        "source_gaps": dict(source_gaps),
        "missing_row_policy": args.missing_row_policy,
        "cumulative_policy": args.cumulative_policy,
        "template_evidence": template_check,
        "indicator_evidence": indicator_check,
        "field_status": {key: dict(value) for key, value in per_field.items()},
    }

    args.output.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    summary_sheet = workbook.active
    summary_sheet.title = "验证汇总"
    summary_rows = [
        ("项目", "值"),
        ("报表项目数", summary["report_rows"]),
        ("字段数", summary["field_count"]),
        ("明细检查数", summary["detail_count"]),
        ("通过", counts.get(STATUS_PASS, 0)),
        ("失败", counts.get(STATUS_FAIL, 0)),
        ("缺源阻塞", counts.get(STATUS_BLOCKED, 0)),
        ("待口径确认", counts.get(STATUS_REVIEW, 0)),
        ("缺行政策", args.missing_row_policy),
        ("累计政策", args.cumulative_policy),
        ("说明", "敏感数据仅保存在本地输出；状态按逐项目逐字段统计。"),
    ]
    for row in summary_rows:
        summary_sheet.append(row)
    summary_sheet.append([])
    summary_sheet.append(["字段", STATUS_PASS, STATUS_FAIL, STATUS_BLOCKED, STATUS_REVIEW])
    for field in sorted(per_field, key=lambda value: field_specs[[x[1] for x in field_specs].index(value.split()[0])][0]):
        field_counts = per_field[field]
        summary_sheet.append(
            [
                field,
                field_counts.get(STATUS_PASS, 0),
                field_counts.get(STATUS_FAIL, 0),
                field_counts.get(STATUS_BLOCKED, 0),
                field_counts.get(STATUS_REVIEW, 0),
            ]
        )

    details_sheet = workbook.create_sheet("逐行明细")
    detail_headers = list(details[0].keys()) if details else []
    details_sheet.append(detail_headers)
    for item in details:
        details_sheet.append([item[header] for header in detail_headers])

    gaps_sheet = workbook.create_sheet("缺源与待确认")
    gaps_sheet.append(["事项", "项目字段次数"])
    for key, value in source_gaps.items():
        gaps_sheet.append([key, value])

    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)
    status_fills = {
        STATUS_PASS: PatternFill("solid", fgColor="E2F0D9"),
        STATUS_FAIL: PatternFill("solid", fgColor="F4CCCC"),
        STATUS_BLOCKED: PatternFill("solid", fgColor="FCE5CD"),
        STATUS_REVIEW: PatternFill("solid", fgColor="FFF2CC"),
    }
    for sheet in workbook.worksheets:
        sheet.freeze_panes = "A2"
        sheet.sheet_view.showGridLines = False
        for cell in sheet[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
        for column_cells in sheet.columns:
            width = min(
                60,
                max(10, max(len(norm_text(cell.value)) for cell in column_cells) + 2),
            )
            sheet.column_dimensions[get_column_letter(column_cells[0].column)].width = width
    if details:
        status_column = detail_headers.index("状态") + 1
        for row_index in range(2, details_sheet.max_row + 1):
            cell = details_sheet.cell(row_index, status_column)
            if cell.value in status_fills:
                cell.fill = status_fills[cell.value]
        details_sheet.auto_filter.ref = details_sheet.dimensions
    summary_sheet.auto_filter.ref = f"A{len(summary_rows)+2}:E{summary_sheet.max_row}"
    workbook.save(args.output)

    if args.summary_json:
        args.summary_json.parent.mkdir(parents=True, exist_ok=True)
        args.summary_json.write_text(
            json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8"
        )

    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 0 if counts.get(STATUS_FAIL, 0) == 0 else 2


if __name__ == "__main__":
    raise SystemExit(main())
