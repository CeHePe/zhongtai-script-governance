from __future__ import annotations

import json
import math
import sys
import warnings
from pathlib import Path

import pandas as pd

from _project_root import find_project_root


ROOT = find_project_root(__file__)
REPORT_MONTH = "202512"
PREVIOUS_YEAR_MONTH = "202412"
VALUE_TOL = 1e-4
ZERO_DENOMINATOR = 1e-6


def u(text: str) -> str:
    return text.encode("utf-8").decode("unicode_escape")


PROJECT = u(r"\u9879\u76ee")
INDICATOR_LIST = u(r"\u6307\u6807\u6e05\u5355")


def configure_stdout() -> None:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")


def find_workbook(*tokens: str) -> Path:
    matches = [
        path
        for path in ROOT.iterdir()
        if path.is_file()
        and path.suffix.lower() in {".xlsx", ".xls"}
        and all(token in path.name for token in tokens)
    ]
    if len(matches) != 1:
        raise RuntimeError(f"Expected one workbook for {tokens!r}, got {len(matches)}: {[p.name for p in matches]}")
    return matches[0]


def normalize_code(value: object) -> str:
    if value is None or pd.isna(value):
        return ""
    text = str(value).strip().upper()
    if text.endswith(".0"):
        text = text[:-2]
    if text in {"", "NAN", "NONE"}:
        return ""
    return text


def as_number(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce").fillna(0.0)


def safe_growth(current: float, previous: float) -> float:
    if abs(previous) < ZERO_DENOMINATOR and not math.isclose(abs(previous), ZERO_DENOMINATOR, abs_tol=1e-12):
        return 0.0
    return (current - previous) / abs(previous)


def safe_share(value: float, total_income: float) -> float:
    if abs(total_income) <= ZERO_DENOMINATOR:
        return 0.0
    return value / total_income


METRICS = [
    {"name": "营业收入", "value": "income", "value_col": 4, "yoy_col": 5, "has_share": False},
    {"name": "物业费&公摊收入", "value": "property_public_income", "value_col": 6, "yoy_col": 7, "share_col": 8, "share_yoy_col": 9},
    {"name": "车位物业费收入", "value": "parking_property_income", "value_col": 10, "yoy_col": 11, "share_col": 12, "share_yoy_col": 13},
    {"name": "示范区收入", "value": "demo_area_income", "value_col": 14, "yoy_col": 15, "share_col": 16, "share_yoy_col": 17},
    {"name": "月租停车收入", "value": "monthly_parking_income", "value_col": 18, "yoy_col": 19, "share_col": 20, "share_yoy_col": 21},
    {"name": "临停收入", "value": "temporary_parking_income", "value_col": 22, "yoy_col": 23, "share_col": 24, "share_yoy_col": 25},
    {"name": "场地/特约服务收入", "value": "site_special_income", "value_col": 26, "yoy_col": 27, "share_col": 28, "share_yoy_col": 29},
    {"name": "能源服务收入", "value": "energy_service_income", "value_col": 30, "yoy_col": 31, "share_col": 32, "share_yoy_col": 33},
    {"name": "建渣收入", "value": "construction_waste_income", "value_col": 34, "yoy_col": 35, "share_col": 36, "share_yoy_col": 37},
    {"name": "其他服务收入", "value": "other_service_income", "value_col": 38, "yoy_col": 39, "share_col": 40, "share_yoy_col": 41},
]


def load_111(month: str) -> pd.DataFrame:
    path = find_workbook("1.1.1", month, PROJECT)
    df = pd.read_excel(path, header=None, dtype=object).iloc[3:].reset_index(drop=True).copy()
    if df.shape[1] != 42:
        raise RuntimeError(f"Unexpected 1.1.1 column count in {path.name}: {df.shape[1]}")
    out = pd.DataFrame(
        {
            "region": df.iloc[:, 0],
            "line": df.iloc[:, 1],
            "project_code": df.iloc[:, 2],
            "project_name": df.iloc[:, 3],
            "code_norm": df.iloc[:, 2].map(normalize_code),
            "source_file": path.name,
        }
    )
    out = out[out["code_norm"].ne("")].copy()
    for metric in METRICS:
        prefix = metric["value"]
        out[prefix] = as_number(df.iloc[out.index, metric["value_col"]])
        out[f"{prefix}_yoy"] = as_number(df.iloc[out.index, metric["yoy_col"]])
        if metric.get("has_share", True):
            out[f"{prefix}_share"] = as_number(df.iloc[out.index, metric["share_col"]])
            out[f"{prefix}_share_yoy"] = as_number(df.iloc[out.index, metric["share_yoy_col"]])
    return out.reset_index(drop=True)


def indicator_context() -> list[dict[str, object]]:
    path = find_workbook("JKS_", INDICATOR_LIST)
    df = pd.read_excel(path, sheet_name=0, header=None, dtype=object)
    wanted_rows = [995]
    rows: list[dict[str, object]] = []
    for excel_row in wanted_rows:
        row = df.iloc[excel_row - 1]
        compact = [value for value in row.tolist() if not pd.isna(value) and value != ""]
        rows.append({"excel_row": excel_row, "values": compact})
    return rows


def summarize(checks: pd.DataFrame, kind: str) -> dict[str, object]:
    subset = checks[checks["kind"].eq(kind)].copy()
    mismatch = subset["abs_diff"] > VALUE_TOL
    return {
        "kind": kind,
        "status": "passed" if not mismatch.any() else "failed",
        "checks": int(len(subset)),
        "mismatch_checks": int(mismatch.sum()),
        "affected_project_metric_pairs": int(subset.loc[mismatch, ["code_norm", "metric"]].drop_duplicates().shape[0]),
        "max_abs_diff": float(subset["abs_diff"].max() if len(subset) else 0.0),
        "diff_total": float(subset["diff"].sum() if len(subset) else 0.0),
    }


def sample_mismatches(checks: pd.DataFrame, kind: str, limit: int = 8) -> list[dict[str, object]]:
    subset = checks[(checks["kind"].eq(kind)) & (checks["abs_diff"] > VALUE_TOL)].copy()
    columns = [
        "kind",
        "metric",
        "region",
        "line",
        "project_code",
        "project_name",
        "report_value",
        "calc_value",
        "diff",
        "current_value",
        "previous_value",
        "current_total_income",
        "previous_total_income",
    ]
    return subset[columns].head(limit).to_dict(orient="records")


def build_checks(current: pd.DataFrame, previous: pd.DataFrame) -> pd.DataFrame:
    base = current.merge(
        previous.add_prefix("prev_"),
        left_on="code_norm",
        right_on="prev_code_norm",
        how="left",
    )
    records: list[dict[str, object]] = []
    for metric in METRICS:
        prefix = metric["value"]
        prev_col = f"prev_{prefix}"
        base[prev_col] = base[prev_col].fillna(0.0)
        calc_yoy = [
            safe_growth(current_value, previous_value)
            for current_value, previous_value in zip(base[prefix], base[prev_col])
        ]
        for idx, calc_value in enumerate(calc_yoy):
            report_value = float(base.at[idx, f"{prefix}_yoy"])
            records.append(
                {
                    "kind": "同比增幅（%）",
                    "metric": metric["name"],
                    "region": base.at[idx, "region"],
                    "line": base.at[idx, "line"],
                    "project_code": base.at[idx, "project_code"],
                    "project_name": base.at[idx, "project_name"],
                    "code_norm": base.at[idx, "code_norm"],
                    "report_value": report_value,
                    "calc_value": float(calc_value),
                    "diff": report_value - float(calc_value),
                    "abs_diff": abs(report_value - float(calc_value)),
                    "current_value": float(base.at[idx, prefix]),
                    "previous_value": float(base.at[idx, prev_col]),
                    "current_total_income": float(base.at[idx, "income"]),
                    "previous_total_income": float(base.at[idx, "prev_income"]) if not pd.isna(base.at[idx, "prev_income"]) else 0.0,
                }
            )

        if not metric.get("has_share", True):
            continue

        calc_share = [safe_share(value, income) for value, income in zip(base[prefix], base["income"])]
        # 占比同比按用户指定参考202412报表：当前占比复算值 - 上年报表占比列。
        prev_share = base[f"prev_{prefix}_share"].fillna(0.0)
        calc_share_yoy = [current_share - previous for current_share, previous in zip(calc_share, prev_share)]

        for idx, calc_value in enumerate(calc_share):
            report_value = float(base.at[idx, f"{prefix}_share"])
            records.append(
                {
                    "kind": "占总收入比（%）",
                    "metric": metric["name"],
                    "region": base.at[idx, "region"],
                    "line": base.at[idx, "line"],
                    "project_code": base.at[idx, "project_code"],
                    "project_name": base.at[idx, "project_name"],
                    "code_norm": base.at[idx, "code_norm"],
                    "report_value": report_value,
                    "calc_value": float(calc_value),
                    "diff": report_value - float(calc_value),
                    "abs_diff": abs(report_value - float(calc_value)),
                    "current_value": float(base.at[idx, prefix]),
                    "previous_value": float(base.at[idx, prev_col]),
                    "current_total_income": float(base.at[idx, "income"]),
                    "previous_total_income": float(base.at[idx, "prev_income"]) if not pd.isna(base.at[idx, "prev_income"]) else 0.0,
                }
            )

        for idx, calc_value in enumerate(calc_share_yoy):
            report_value = float(base.at[idx, f"{prefix}_share_yoy"])
            records.append(
                {
                    "kind": "占总收入比的同比（%）",
                    "metric": metric["name"],
                    "region": base.at[idx, "region"],
                    "line": base.at[idx, "line"],
                    "project_code": base.at[idx, "project_code"],
                    "project_name": base.at[idx, "project_name"],
                    "code_norm": base.at[idx, "code_norm"],
                    "report_value": report_value,
                    "calc_value": float(calc_value),
                    "diff": report_value - float(calc_value),
                    "abs_diff": abs(report_value - float(calc_value)),
                    "current_value": float(base.at[idx, prefix]),
                    "previous_value": float(base.at[idx, prev_col]),
                    "current_total_income": float(base.at[idx, "income"]),
                    "previous_total_income": float(base.at[idx, "prev_income"]) if not pd.isna(base.at[idx, "prev_income"]) else 0.0,
                }
            )
    return pd.DataFrame(records)


def main() -> None:
    configure_stdout()
    warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

    current = load_111(REPORT_MONTH)
    previous = load_111(PREVIOUS_YEAR_MONTH)
    checks = build_checks(current, previous)

    result = {
        "period": REPORT_MONTH,
        "previous_period": PREVIOUS_YEAR_MONTH,
        "dimension": "project",
        "rules": {
            "同比增幅（%）": "(202512累计值 - 202412累计值) / abs(202412累计值); 上年绝对值<=1e-6按0",
            "占总收入比（%）": "202512分项累计值 / 202512营业收入; 营业收入绝对值<=1e-6按0",
            "占总收入比的同比（%）": "202512占总收入比 - 202412占总收入比",
        },
        "source_files": [
            find_workbook("1.1.1", REPORT_MONTH, PROJECT).name,
            find_workbook("1.1.1", PREVIOUS_YEAR_MONTH, PROJECT).name,
        ],
        "indicator_context": indicator_context(),
        "coverage": {
            "current_report_rows": int(len(current)),
            "previous_report_rows": int(len(previous)),
            "previous_matched_rows": int(current["code_norm"].isin(set(previous["code_norm"])).sum()),
            "metrics_with_yoy": [metric["name"] for metric in METRICS],
            "metrics_with_share": [metric["name"] for metric in METRICS if metric.get("has_share", True)],
        },
        "summaries": [
            summarize(checks, "同比增幅（%）"),
            summarize(checks, "占总收入比（%）"),
            summarize(checks, "占总收入比的同比（%）"),
        ],
        "samples": {
            "同比增幅（%）": sample_mismatches(checks, "同比增幅（%）"),
            "占总收入比（%）": sample_mismatches(checks, "占总收入比（%）"),
            "占总收入比的同比（%）": sample_mismatches(checks, "占总收入比的同比（%）"),
        },
    }

    output_dir = ROOT / "local_outputs"
    output_dir.mkdir(exist_ok=True)
    output_path = output_dir / f"validate_111_project_income_derived_metrics_{REPORT_MONTH}.json"
    checks_path = output_dir / f"validate_111_project_income_derived_metrics_{REPORT_MONTH}_checks.csv"
    output_path.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    checks.to_csv(checks_path, index=False, encoding="utf-8-sig")
    print(json.dumps(result, ensure_ascii=False, indent=2))
    print(f"\nWrote {output_path}")
    print(f"Wrote {checks_path}")


if __name__ == "__main__":
    main()
