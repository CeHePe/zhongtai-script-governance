from __future__ import annotations

import importlib.util
import json
from pathlib import Path
from _project_root import find_project_root

import pandas as pd


ROOT = find_project_root(__file__)


def load_base_module():
    """复用已有半收付归母净利润脚本中的报表读取和编码标准化函数。"""
    script_path = ROOT / "validate_half_cash_attributable_profit.py"
    spec = importlib.util.spec_from_file_location("half_cash_base", script_path)
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(module)
    return module


def print_record(label: str, record: dict) -> None:
    """统一输出 JSON，便于后续留痕和解析。"""
    print(f"{label} {json.dumps(record, ensure_ascii=False)}")


def summarize_diff(base, label: str, calc: pd.Series, report: pd.Series) -> dict:
    """汇总计算值与报表值的差异。"""
    diff = calc - report
    return {
        "check": label,
        "status": "passed" if (diff.abs() <= base.TOLERANCE).all() else "failed",
        "rows": int(len(diff)),
        "mismatch_rows": int((diff.abs() > base.TOLERANCE).sum()),
        "calc_total": float(calc.sum()),
        "report_total": float(report.sum()),
        "diff_total": float(diff.sum()),
        "max_abs_diff": float(diff.abs().max() if len(diff) else 0.0),
    }


def find_finance_cloud_file(base) -> Path:
    """优先匹配最新的财务云/实际发生数文件，兼容 xlsx 和 xls。"""
    finance_cloud = base.u(r"\u8d22\u52a1\u4e91")
    actual_token = base.u(r"\u5b9e\u9645\u53d1\u751f\u6570")
    matches = [
        path
        for path in ROOT.iterdir()
        if path.is_file()
        and path.suffix.lower() in {".xlsx", ".xls"}
        and (finance_cloud in path.name or actual_token in path.name)
    ]
    if not matches:
        raise FileNotFoundError("未找到财务云/实际发生数文件")
    matches.sort(key=lambda path: (path.suffix.lower() != ".xls", path.name, path.stat().st_mtime))
    return matches[-1]


def load_finance_cloud_actuals(base, query_df: pd.DataFrame) -> pd.DataFrame:
    """读取财务云实际发生数，并按已验证口径映射为三个发生数字段。"""
    cloud_path = find_finance_cloud_file(base)
    df = pd.read_excel(cloud_path, dtype=object)
    rename_map = {
        "sub_project_code": "立项编码",
        "share_type_name": "摊销类型名称",
        "no_tax_dist_amount": "分摊不含税金额",
    }
    for expected, current in rename_map.items():
        if expected not in df.columns and current in df.columns:
            df[expected] = df[current]

    df["code_norm"] = df["sub_project_code"].map(base.normalize_code)
    df["share_type_name"] = df["share_type_name"].astype(str).str.strip()
    df["no_tax_amount"] = pd.to_numeric(df["no_tax_dist_amount"], errors="coerce").fillna(0.0)

    type_to_component = {
        base.u(r"\u5e26\u8d44\u8fdb\u573a"): "actual_capital",
        base.u(r"\u6280\u672f\u6539\u9020"): "actual_smart",
        base.u(r"\u7f8e\u597d\u5bb6\u56ed"): "actual_quality",
        base.u(r"\u5b58\u91cf\u7cbe\u8015"): "actual_quality",
        base.u(r"\u6807\u6746\u9879\u76ee"): "actual_quality",
    }
    df["component"] = df["share_type_name"].map(type_to_component)
    df = df.dropna(subset=["component"]).copy()

    df = df.merge(query_df[["code_norm", "penetration_ratio"]].drop_duplicates("code_norm"), on="code_norm", how="left")
    df["penetration_ratio"] = df["penetration_ratio"].fillna(1.0)

    # 统一口径：财务云不含税金额 * 0.81 * 项目穿透比例。
    df["amount"] = df["no_tax_amount"] * 0.81 * df["penetration_ratio"]
    return df


def pivot_actuals(df: pd.DataFrame, index_cols: list[str]) -> pd.DataFrame:
    """把长表发生数转成带资、智能化、质效三列。"""
    wide = (
        df.groupby(index_cols + ["component"], as_index=False)["amount"]
        .sum()
        .pivot(index=index_cols, columns="component", values="amount")
        .reset_index()
        .fillna(0.0)
    )
    for component in ["actual_capital", "actual_smart", "actual_quality"]:
        if component not in wide.columns:
            wide[component] = 0.0
    return wide


def compare_project_actuals(base, project_df: pd.DataFrame, actuals: pd.DataFrame) -> None:
    """项目维度：财务云源表复算结果对比半收付项目报表。"""
    wide = pivot_actuals(actuals, ["code_norm"])
    check_df = project_df[
        ["region", "project_code", "project_name", "code_norm", "actual_capital", "actual_smart", "actual_quality"]
    ].merge(wide[["code_norm", "actual_capital", "actual_smart", "actual_quality"]], on="code_norm", how="left", suffixes=("_report", "_calc"))
    check_df = check_df.fillna(0.0)

    for component in ["actual_capital", "actual_smart", "actual_quality"]:
        diff = check_df[f"{component}_calc"] - check_df[f"{component}_report"]
        print_record(
            f"PROJECT_{component}",
            summarize_diff(base, f"project_{component}", check_df[f"{component}_calc"], check_df[f"{component}_report"]),
        )
        bad = check_df.loc[
            diff.abs() > base.TOLERANCE,
            ["region", "project_code", "project_name", f"{component}_report", f"{component}_calc"],
        ].copy()
        if not bad.empty:
            bad["diff"] = diff[bad.index]
            print(
                bad.sort_values("diff", key=lambda series: series.abs(), ascending=False)
                .head(20)
                .to_json(orient="records", force_ascii=False)
            )


def compare_region_actuals(base, region_df: pd.DataFrame, query_df: pd.DataFrame, actuals: pd.DataFrame) -> None:
    """区域维度：财务云源表按项目查询映射区域后汇总，对比半收付区域报表。"""
    actuals_with_region = actuals.merge(query_df[["code_norm", "region"]].drop_duplicates("code_norm"), on="code_norm", how="left")
    wide = pivot_actuals(actuals_with_region, ["region"])
    check_df = region_df[["region", "actual_capital", "actual_smart", "actual_quality"]].merge(
        wide[["region", "actual_capital", "actual_smart", "actual_quality"]],
        on="region",
        how="left",
        suffixes=("_report", "_calc"),
    )
    check_df = check_df.fillna(0.0)

    for component in ["actual_capital", "actual_smart", "actual_quality"]:
        diff = check_df[f"{component}_calc"] - check_df[f"{component}_report"]
        print_record(
            f"REGION_{component}",
            summarize_diff(base, f"region_{component}", check_df[f"{component}_calc"], check_df[f"{component}_report"]),
        )
        bad = check_df.loc[
            diff.abs() > base.TOLERANCE,
            ["region", f"{component}_report", f"{component}_calc"],
        ].copy()
        if not bad.empty:
            bad["diff"] = diff[bad.index]
            print(bad.to_json(orient="records", force_ascii=False))


def compare_region_from_project_report(base, project_df: pd.DataFrame, region_df: pd.DataFrame) -> None:
    """诊断项：区域报表是否等于项目报表汇总。"""
    rollup = project_df.groupby("region", as_index=False)[["actual_capital", "actual_smart", "actual_quality"]].sum()
    check_df = region_df[["region", "actual_capital", "actual_smart", "actual_quality"]].merge(
        rollup,
        on="region",
        how="left",
        suffixes=("_report", "_project"),
    )
    check_df = check_df.fillna(0.0)

    for component in ["actual_capital", "actual_smart", "actual_quality"]:
        print_record(
            f"REGION_PROJECT_ROLLUP_{component}",
            summarize_diff(
                base,
                f"region_project_rollup_{component}",
                check_df[f"{component}_project"],
                check_df[f"{component}_report"],
            ),
        )


def compare_space_actuals(base, region_df: pd.DataFrame, actuals: pd.DataFrame) -> None:
    """空间维度：当前报表专业公司发生数全为 0，先按区域结果汇总对比空间报表。"""
    from validate_half_cash_requested_metrics import load_space_report

    space_df = load_space_report()
    actual_total = pivot_actuals(actuals, ["code_norm"])[["actual_capital", "actual_smart", "actual_quality"]].sum()
    region_total = region_df[["actual_capital", "actual_smart", "actual_quality"]].sum()

    for component in ["actual_capital", "actual_smart", "actual_quality"]:
        report_series = space_df[component]
        calc_series = pd.Series([float(actual_total[component])])
        print_record(
            f"SPACE_{component}",
            {
                **summarize_diff(base, f"space_{component}", calc_series, report_series),
                "region_report_total": float(region_total[component]),
                "space_report_total": float(report_series.sum()),
            },
        )


def main() -> None:
    base = load_base_module()
    project_df = base.load_project_report()
    region_df = base.load_region_report()
    query_df = base.load_query()
    actuals = load_finance_cloud_actuals(base, query_df)

    print_record(
        "SOURCE_SUMMARY",
        {
            "rows": int(len(actuals)),
            "project_count": int(actuals["code_norm"].nunique()),
            "amount_total": float(actuals["amount"].sum()),
        },
    )
    compare_project_actuals(base, project_df, actuals)
    compare_region_actuals(base, region_df, query_df, actuals)
    compare_region_from_project_report(base, project_df, region_df)
    compare_space_actuals(base, region_df, actuals)


if __name__ == "__main__":
    main()
