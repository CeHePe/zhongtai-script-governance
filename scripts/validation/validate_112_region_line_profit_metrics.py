from __future__ import annotations

import json
import sys
import warnings
from pathlib import Path

import pandas as pd

from _project_root import find_project_root


ROOT = find_project_root(__file__)
REPORT_MONTH = "202512"
PREVIOUS_MONTH = "202412"
VALUE_TOL = 1e-6  # 万元；差异小于0.01元才忽略，0.01元及以上不通过
RATE_TOL = 1e-6
ZERO_DENOMINATOR = 1e-6


def u(text: str) -> str:
    return text.encode("utf-8").decode("unicode_escape")


REGION = u(r"\u533a\u57df")
LINE = u(r"\u6761\u7ebf")
REGION_LINE = u(r"\u533a\u57df\u6761\u7ebf")
PROJECT_QUERY = u(r"\u9879\u76ee\u67e5\u8be2")
HALF_ATTR = u(r"\u534a\u6536\u4ed8\u5f52\u6bcd\u51c0\u5229\u6da6")
MGMT_ATTR = u(r"\u7ba1\u62a5\u5f52\u6bcd\u51c0\u5229\u6da6")


CURRENT_METRICS = [
    ("半收付净利润", "half_net_profit", 1),
    ("半收付归母净利润", "half_attr_profit", 3),
    ("半收付收入", "half_revenue", 5),
    ("管报净利润", "mgmt_net_profit", 7),
    ("管报归母净利润", "mgmt_attr_profit", 9),
]

YOY_METRICS = [
    ("半收付净利润", "half_net_profit", 1, 2),
    ("半收付归母净利润", "half_attr_profit", 3, 4),
    ("半收付收入", "half_revenue", 5, 6),
    ("管报净利润", "mgmt_net_profit", 7, 8),
    ("管报归母净利润", "mgmt_attr_profit", 9, 10),
    ("营业收入", "operating_revenue", 11, 12),
    ("累计回收现金流", "cumulative_cashflow", 13, 14),
    ("折让", "discount", 17, 18),
    ("营业成本", "operating_cost", 21, 22),
    ("管理费用", "management_expense", 25, 26),
]


def configure_stdout() -> None:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")


def find_workbook(*tokens: str, exclude: tuple[str, ...] = ()) -> Path:
    matches = [
        path
        for path in ROOT.iterdir()
        if path.is_file()
        and path.suffix.lower() in {".xlsx", ".xls"}
        and all(token in path.name for token in tokens)
        and not any(token in path.name for token in exclude)
    ]
    if len(matches) != 1:
        raise RuntimeError(f"Expected one workbook for {tokens!r}, got {len(matches)}: {[p.name for p in matches]}")
    return matches[0]


def as_number(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce").fillna(0.0)


def normalize_code(value: object) -> str:
    if value is None or pd.isna(value):
        return ""
    text = str(value).strip().upper()
    if text.endswith(".0"):
        text = text[:-2]
    if text in {"", "NAN", "NONE"}:
        return ""
    return text


def safe_growth(current: float, previous: float) -> float:
    if abs(previous) < ZERO_DENOMINATOR:
        return 0.0
    return (current - previous) / abs(previous)


def signed_growth(current: float, previous: float) -> float:
    if abs(previous) < ZERO_DENOMINATOR:
        return 0.0
    return (current - previous) / previous


def load_112(month: str, dimension: str) -> pd.DataFrame:
    token = REGION if dimension == "region" else LINE
    matches = [
        path
        for path in ROOT.iterdir()
        if path.is_file()
        and path.suffix.lower() == ".xlsx"
        and "1.1.2" in path.name
        and month in path.name
        and path.name.endswith(f"{token}.xlsx")
    ]
    if len(matches) != 1:
        raise RuntimeError(f"Expected one 1.1.2 workbook for {month}/{token}, got {[p.name for p in matches]}")
    path = matches[0]
    raw = pd.read_excel(path, header=None, dtype=object).iloc[3:].reset_index(drop=True).copy()
    if raw.shape[1] != 29:
        raise RuntimeError(f"Unexpected 1.1.2 column count in {path.name}: {raw.shape[1]}")
    out = pd.DataFrame({"key": raw.iloc[:, 0].astype(str).str.strip(), "source_file": path.name})
    for _, field, value_col, yoy_col in YOY_METRICS:
        out[field] = as_number(raw.iloc[:, value_col])
        out[f"{field}_yoy"] = as_number(raw.iloc[:, yoy_col])
    return out


def exact_dimension_source(prefix: str, dimension: str) -> Path:
    if dimension == "region":
        return find_workbook(prefix, REPORT_MONTH, REGION, exclude=(REGION_LINE,))
    return find_workbook(prefix, REPORT_MONTH, LINE, exclude=(REGION, REGION_LINE))


def load_profit_sources(dimension: str) -> pd.DataFrame:
    half_path = exact_dimension_source(HALF_ATTR, dimension)
    mgmt_path = exact_dimension_source(MGMT_ATTR, dimension)

    half = pd.read_excel(half_path, dtype=object)
    half_key_col = 0
    half["key"] = half.iloc[:, half_key_col].astype(str).str.strip()
    half["src_half_net_profit"] = as_number(half.iloc[:, 1]) / 10000.0
    half["src_half_attr_profit"] = as_number(half.iloc[:, 31]) / 10000.0

    mgmt = pd.read_excel(mgmt_path, dtype=object)
    mgmt["key"] = mgmt.iloc[:, 1].astype(str).str.strip()
    mgmt["src_mgmt_net_profit"] = as_number(mgmt.iloc[:, 2]) / 10000.0
    mgmt["src_mgmt_attr_profit"] = as_number(mgmt.iloc[:, 23]) / 10000.0

    return half[["key", "src_half_net_profit", "src_half_attr_profit"]].merge(
        mgmt[["key", "src_mgmt_net_profit", "src_mgmt_attr_profit"]],
        on="key",
        how="outer",
    )


def load_half_revenue_source(dimension: str) -> tuple[pd.DataFrame, dict[str, object]]:
    path = find_workbook("1.5.2", REPORT_MONTH)
    raw = pd.read_excel(path, header=None, dtype=object).iloc[5:].reset_index(drop=True).copy()
    raw["region"] = raw.iloc[:, 0].astype(str).str.strip()
    raw["code_norm"] = raw.iloc[:, 1].map(normalize_code)
    raw["amount"] = as_number(raw.iloc[:, 12]) / 10000.0

    diagnostics: dict[str, object] = {"source_file": path.name, "source_rows": int(len(raw))}
    if dimension == "region":
        source = raw.groupby("region", as_index=False)["amount"].sum().rename(
            columns={"region": "key", "amount": "src_half_revenue"}
        )
        return source, diagnostics

    query_path = find_workbook(PROJECT_QUERY)
    query = pd.read_excel(query_path, dtype=object)
    query["code_norm"] = query.iloc[:, 0].map(normalize_code)
    query["line"] = query.iloc[:, 4].astype(str).str.strip()
    query = query.drop_duplicates("code_norm")
    mapped = raw.merge(query[["code_norm", "line"]], on="code_norm", how="left")
    missing = mapped[mapped["line"].isna()]
    diagnostics.update(
        {
            "mapping_file": query_path.name,
            "unmapped_rows": int(len(missing)),
            "unmapped_nonzero_rows": int((missing["amount"].abs() >= VALUE_TOL).sum()),
            "unmapped_amount_total": float(missing["amount"].sum()),
        }
    )
    source = mapped.groupby("line", as_index=False)["amount"].sum().rename(
        columns={"line": "key", "amount": "src_half_revenue"}
    )
    return source, diagnostics


def summarize_current(base: pd.DataFrame, label: str, report_col: str, calc_col: str) -> dict[str, object]:
    diff = base[report_col] - base[calc_col]
    mismatch = diff.abs() >= VALUE_TOL
    return {
        "check": label,
        "status": "passed" if not mismatch.any() else "failed",
        "rows": int(len(base)),
        "mismatch_rows": int(mismatch.sum()),
        "max_abs_diff": float(diff.abs().max() if len(diff) else 0.0),
        "diff_total": float(diff.sum()),
    }


def summarize_yoy(checks: pd.DataFrame, dimension: str) -> dict[str, object]:
    subset = checks[checks["dimension"].eq(dimension)]
    mismatch = subset["abs_diff"] > RATE_TOL
    return {
        "dimension": "区域" if dimension == "region" else "条线",
        "status": "passed" if not mismatch.any() else "failed",
        "checks": int(len(subset)),
        "mismatch_checks": int(mismatch.sum()),
        "affected_key_metric_pairs": int(subset.loc[mismatch, ["key", "metric"]].drop_duplicates().shape[0]),
        "max_abs_diff": float(subset["abs_diff"].max() if len(subset) else 0.0),
    }


def current_samples(base: pd.DataFrame, report_col: str, calc_col: str, limit: int = 8) -> list[dict[str, object]]:
    diff = base[report_col] - base[calc_col]
    bad = base.loc[diff.abs() >= VALUE_TOL, ["key", report_col, calc_col]].copy()
    bad["diff"] = diff.loc[bad.index]
    return bad.head(limit).to_dict(orient="records")


def build_yoy_checks(current: pd.DataFrame, previous: pd.DataFrame, dimension: str) -> pd.DataFrame:
    base = current.merge(previous, on="key", how="left", suffixes=("", "_previous"))
    records: list[dict[str, object]] = []
    for label, field, _, _ in YOY_METRICS:
        previous_field = f"{field}_previous"
        base[previous_field] = base[previous_field].fillna(0.0)
        growth_func = signed_growth if field == "discount" else safe_growth
        calc = [
            growth_func(current_value, previous_value)
            for current_value, previous_value in zip(base[field], base[previous_field])
        ]
        for idx, calc_value in enumerate(calc):
            report_value = float(base.at[idx, f"{field}_yoy"])
            records.append(
                {
                    "dimension": dimension,
                    "key": base.at[idx, "key"],
                    "metric": label,
                    "current_value": float(base.at[idx, field]),
                    "previous_value": float(base.at[idx, previous_field]),
                    "report_yoy": report_value,
                    "calc_yoy": float(calc_value),
                    "diff": report_value - float(calc_value),
                    "abs_diff": abs(report_value - float(calc_value)),
                }
            )
    return pd.DataFrame(records)


def main() -> None:
    configure_stdout()
    warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

    current_results: dict[str, pd.DataFrame] = {}
    current_summaries: list[dict[str, object]] = []
    current_sample_output: dict[str, dict[str, list[dict[str, object]]]] = {}
    half_revenue_diagnostics: dict[str, dict[str, object]] = {}
    yoy_frames: list[pd.DataFrame] = []

    for dimension in ["region", "line"]:
        current = load_112(REPORT_MONTH, dimension)
        previous = load_112(PREVIOUS_MONTH, dimension)
        profit_sources = load_profit_sources(dimension)
        revenue_source, diagnostics = load_half_revenue_source(dimension)
        half_revenue_diagnostics[dimension] = diagnostics

        base = current.merge(profit_sources, on="key", how="left").merge(revenue_source, on="key", how="left")
        for column in [
            "src_half_net_profit",
            "src_half_attr_profit",
            "src_half_revenue",
            "src_mgmt_net_profit",
            "src_mgmt_attr_profit",
        ]:
            base[column] = base[column].fillna(0.0)

        current_results[dimension] = base
        current_sample_output[dimension] = {}
        for label, field, _ in CURRENT_METRICS:
            calc_col = f"src_{field}"
            summary = summarize_current(base, f"{label}累计完成值", field, calc_col)
            summary["dimension"] = "区域" if dimension == "region" else "条线"
            current_summaries.append(summary)
            current_sample_output[dimension][label] = current_samples(base, field, calc_col)

        yoy_frames.append(build_yoy_checks(current, previous, dimension))

    yoy_checks = pd.concat(yoy_frames, ignore_index=True)
    yoy_samples = yoy_checks[yoy_checks["abs_diff"] > RATE_TOL].head(20).to_dict(orient="records")

    result = {
        "period": REPORT_MONTH,
        "previous_period": PREVIOUS_MONTH,
        "report": "1.1.2 区域及项目利润分析",
        "dimensions": ["区域", "条线"],
        "current_value_rule": "五个指定累计完成值与对应半收付/管报源报表对比；源报表金额单位元，转万元",
        "yoy_rule": "除折让外：(202512累计完成值 - 202412同维度累计完成值) / abs(202412累计完成值)；折让同比使用带符号的202412折让作为分母；上年绝对值小于0.000001万元时按0",
        "yoy_metrics": [label for label, _, _, _ in YOY_METRICS],
        "half_revenue_diagnostics": half_revenue_diagnostics,
        "current_value_summaries": current_summaries,
        "yoy_summaries": [
            summarize_yoy(yoy_checks, "region"),
            summarize_yoy(yoy_checks, "line"),
        ],
        "samples": {
            "current_values": current_sample_output,
            "yoy": yoy_samples,
        },
    }

    output_dir = ROOT / "local_outputs"
    output_dir.mkdir(exist_ok=True)
    json_path = output_dir / f"validate_112_region_line_profit_metrics_{REPORT_MONTH}.json"
    yoy_path = output_dir / f"validate_112_region_line_profit_metrics_{REPORT_MONTH}_yoy_checks.csv"
    region_path = output_dir / f"validate_112_region_line_profit_metrics_{REPORT_MONTH}_region.csv"
    line_path = output_dir / f"validate_112_region_line_profit_metrics_{REPORT_MONTH}_line.csv"
    json_path.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    yoy_checks.to_csv(yoy_path, index=False, encoding="utf-8-sig")
    current_results["region"].to_csv(region_path, index=False, encoding="utf-8-sig")
    current_results["line"].to_csv(line_path, index=False, encoding="utf-8-sig")
    print(json.dumps(result, ensure_ascii=False, indent=2))
    print(f"\nWrote {json_path}")
    print(f"Wrote {yoy_path}")
    print(f"Wrote {region_path}")
    print(f"Wrote {line_path}")


if __name__ == "__main__":
    main()
