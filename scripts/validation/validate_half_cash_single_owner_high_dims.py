from __future__ import annotations

import importlib.util
from pathlib import Path
from _project_root import find_project_root

import pandas as pd


ROOT = find_project_root(__file__)
TOLERANCE = 1e-6


def load_base_module():
    """复用半收付归母净利润脚本中的公共读取逻辑。"""
    script_path = ROOT / "validate_half_cash_attributable_profit.py"
    spec = importlib.util.spec_from_file_location("half_cash_base", script_path)
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(module)
    return module


def find_workbook_by_tokens(*tokens: str) -> Path:
    """按中文 token 安全匹配工作簿文件名。"""
    matches = [
        path
        for path in ROOT.iterdir()
        if path.suffix.lower() == ".xlsx" and all(token in path.name for token in tokens)
    ]
    if len(matches) != 1:
        raise RuntimeError(f"Expected one workbook for {tokens}, got {len(matches)}: {[p.name for p in matches]}")
    return matches[0]


def load_indicator_rows(base) -> pd.DataFrame:
    """读取两个单一业权未到账期指标在条线和专业公司相关维度的指标清单行。"""
    indicator = pd.read_excel(base.find_workbook(base.INDICATOR), sheet_name=0, dtype=object)
    cols = list(indicator.columns)
    metric_current = base.u(r"\u5355\u4e00\u4e1a\u6743\u5f53\u5e74\u5e94\u6536\u5728\u5f53\u671f\u672a\u5230\u8d26\u671f\u91d1\u989d")
    metric_prev = base.u(r"\u5355\u4e00\u4e1a\u6743\u4e0a\u4e00\u5e74\u672b\u5e94\u6536\u672a\u5230\u8d26\u671f\u91d1\u989d")
    target_dims = [
        base.u(r"\u4f4f\u5b85"),
        base.u(r"\u653f\u4f01"),
        base.u(r"\u5b89\u4fdd"),
        base.u(r"\u91d1\u4ee4\u91d1\u5320"),
        base.u(r"\u91d1\u9890"),
        base.u(r"\u97f5\u6db5"),
    ]
    rows = indicator[
        indicator[cols[3]].astype(str).isin([metric_current, metric_prev])
        & indicator[cols[4]].astype(str).isin(target_dims)
    ][[cols[3], cols[4], cols[5], cols[8], cols[10], cols[12]]].fillna("")
    rows = rows.copy()
    rows.insert(0, "excel_row", rows.index + 2)
    rows.columns = ["excel_row", "metric_name", "dimension", "cycle", "method", "source_table", "logic"]
    return rows


def load_line_report(base) -> pd.DataFrame:
    """读取新增的条线维度半收付报表。"""
    path = find_workbook_by_tokens(base.HALF_ATTRIBUTABLE, base.u(r"\u6761\u7ebf"))
    df = pd.read_excel(path, sheet_name=0, dtype=object)
    df = df[df.iloc[:, 0].notna()].copy()
    amount_cols = [
        base.u(r"\u5355\u4e00\u4e1a\u6743\u5f53\u5e74\u5e94\u6536\u5728\u5f53\u671f\u672a\u5230\u8d26\u671f\u91d1\u989d"),
        base.u(r"\u5355\u4e00\u4e1a\u6743\u4e0a\u4e00\u5e74\u672b\u5e94\u6536\u672a\u5230\u8d26\u671f\u91d1\u989d"),
    ]
    for column in amount_cols:
        df[column] = pd.to_numeric(df[column], errors="coerce").fillna(0.0)
    return df.rename(columns={df.columns[0]: "dimension_value"})


def load_professional_report(base) -> pd.DataFrame:
    """读取新增的专业公司维度半收付报表。"""
    path = find_workbook_by_tokens(base.HALF_ATTRIBUTABLE, base.u(r"\u4e13\u4e1a\u516c\u53f8"))
    df = pd.read_excel(path, sheet_name=0, dtype=object)
    df = df[df.iloc[:, 0].notna()].copy()
    amount_cols = [
        base.u(r"\u5355\u4e00\u4e1a\u6743\u5f53\u5e74\u5e94\u6536\u5728\u5f53\u671f\u672a\u5230\u8d26\u671f\u91d1\u989d"),
        base.u(r"\u5355\u4e00\u4e1a\u6743\u4e0a\u4e00\u5e74\u672b\u5e94\u6536\u672a\u5230\u8d26\u671f\u91d1\u989d"),
    ]
    for column in amount_cols:
        df[column] = pd.to_numeric(df[column], errors="coerce").fillna(0.0)
    return df.rename(columns={df.columns[0]: "dimension_value"})


def build_line_calc(base) -> pd.DataFrame:
    """按指标清单要求，从项目报表汇总住宅/政企条线的累计值。"""
    project_df = base.load_project_report()
    profit_non_assess_codes = base.load_profit_non_assess_codes()
    single_owner = base.u(r"\u5355\u4e00")
    source = project_df[
        project_df["code_exact"].ne("")
        & ~project_df["code_norm"].isin(profit_non_assess_codes)
        & project_df["ownership_attr"].astype(str).str.contains(single_owner, na=False, regex=False)
    ].copy()
    rollup = (
        source.groupby("line", as_index=False)[["single_current_unreceived_2", "single_prev_unreceived_2"]]
        .sum()
        .rename(
            columns={
                "line": "dimension_value",
                "single_current_unreceived_2": "calc_current",
                "single_prev_unreceived_2": "calc_prev",
            }
        )
    )
    return rollup


def compare_line(base) -> pd.DataFrame:
    """条线报表与项目汇总结果对比。"""
    metric_current = base.u(r"\u5355\u4e00\u4e1a\u6743\u5f53\u5e74\u5e94\u6536\u5728\u5f53\u671f\u672a\u5230\u8d26\u671f\u91d1\u989d")
    metric_prev = base.u(r"\u5355\u4e00\u4e1a\u6743\u4e0a\u4e00\u5e74\u672b\u5e94\u6536\u672a\u5230\u8d26\u671f\u91d1\u989d")
    report = load_line_report(base)[["dimension_value", metric_current, metric_prev]].rename(
        columns={metric_current: "report_current", metric_prev: "report_prev"}
    )
    calc = build_line_calc(base)
    compare = report.merge(calc, on="dimension_value", how="left").fillna(0.0)
    compare["diff_current"] = compare["calc_current"] - compare["report_current"]
    compare["diff_prev"] = compare["calc_prev"] - compare["report_prev"]
    compare["status"] = (
        compare["diff_current"].abs().le(TOLERANCE) & compare["diff_prev"].abs().le(TOLERANCE)
    ).map({True: "passed", False: "failed"})
    return compare


def find_professional_source(base) -> list[Path]:
    """查找指标清单指定的专业公司回款营收比调整项台账或含目标字段的替代台账。"""
    source_name = base.u(r"\u4e13\u4e1a\u516c\u53f8\u7684\u56de\u6b3e\u8425\u6536\u6bd4\u8c03\u6574\u9879\u53f0\u8d26")
    current_field = base.u(r"\u672a\u8fbe\u8d26\u671f\u91d1\u989d\uff1a\u672c\u5e74\u622a\u81f3\u5230\u5f53\u524d\u6708\u4efd").replace("e", "")
    prev_field = base.u(r"\u672a\u8fbe\u8d26\u671f\u91d1\u989d\uff1a\u672c\u5e74\u671f\u521d\u672a\u5230\u8d26\u671f\u91d1\u989d").replace("e", "")

    direct = [
        path
        for path in ROOT.iterdir()
        if path.suffix.lower() == ".xlsx" and source_name in path.name
    ]
    if direct:
        return direct

    candidates: list[Path] = []
    for path in ROOT.iterdir():
        if path.suffix.lower() != ".xlsx":
            continue
        try:
            xl = pd.ExcelFile(path)
            for sheet_name in xl.sheet_names:
                preview = pd.read_excel(path, sheet_name=sheet_name, nrows=20, dtype=object)
                text = "\n".join(map(str, preview.columns)) + "\n" + preview.astype(str).to_string()
                if current_field in text and prev_field in text:
                    candidates.append(path)
                    break
        except Exception:
            continue
    return candidates


def load_professional_source(base) -> pd.DataFrame:
    """按指标清单口径计算专业公司两个未到账期指标。"""
    sources = find_professional_source(base)
    if not sources:
        return pd.DataFrame(
            columns=[
                "dimension_value",
                "legal_company",
                "penetration_ratio",
                "source_current",
                "source_prev",
                "calc_current",
                "calc_prev",
            ]
        )

    source_path = sources[0]
    source_df = pd.read_excel(source_path, sheet_name=0, header=[1, 2], dtype=object)
    source_df = source_df.iloc[:, [0, 1, 41, 38]].copy()
    source_df.columns = ["dimension_value", "report_month", "source_current", "source_prev"]
    source_df["dimension_value"] = source_df["dimension_value"].astype(str).str.strip()
    source_df["report_month"] = source_df["report_month"].astype(str).str.strip()
    source_df = source_df[source_df["report_month"].eq("2025-12")].copy()
    source_df["source_current"] = pd.to_numeric(source_df["source_current"], errors="coerce").fillna(0.0)
    source_df["source_prev"] = pd.to_numeric(source_df["source_prev"], errors="coerce").fillna(0.0)

    config_path = find_workbook_by_tokens(base.u(r"\u4e13\u4e1a\u516c\u53f8\u6cd5\u4eba\u516c\u53f8\u914d\u7f6e\u8868"))
    config_df = pd.read_excel(config_path, header=None, dtype=object)
    config_df = config_df.iloc[:, :2].copy()
    config_df.columns = ["dimension_value", "legal_company"]
    config_df["dimension_value"] = config_df["dimension_value"].astype(str).str.strip()
    config_df["legal_company"] = config_df["legal_company"].astype(str).str.strip()

    ratio_path = find_workbook_by_tokens(base.u(r"\u9879\u76ee\u6cd5\u4eba\u516c\u53f8\u7a7f\u900f\u6bd4\u4f8b\u914d\u7f6e"))
    ratio_df = pd.read_excel(ratio_path, dtype=object)
    ratio_df = ratio_df.iloc[:, :3].copy()
    ratio_df.columns = ["legal_company", "legal_code", "penetration_ratio"]
    ratio_df["legal_company"] = ratio_df["legal_company"].astype(str).str.strip()
    ratio_df["penetration_ratio"] = pd.to_numeric(ratio_df["penetration_ratio"], errors="coerce")

    # 同一法人如果配置了多条不同穿透比例，先标记出来，避免静默取错。
    ratio_unique = (
        ratio_df.groupby("legal_company", as_index=False)["penetration_ratio"]
        .agg(lambda values: values.dropna().unique().tolist())
    )
    ratio_unique["ratio_count"] = ratio_unique["penetration_ratio"].map(len)
    ratio_unique["penetration_ratio"] = ratio_unique["penetration_ratio"].map(
        lambda values: values[0] if len(values) == 1 else None
    )

    result = config_df.merge(source_df, on="dimension_value", how="left").merge(
        ratio_unique[["legal_company", "penetration_ratio", "ratio_count"]],
        on="legal_company",
        how="left",
    )
    result["source_current"] = pd.to_numeric(result["source_current"], errors="coerce").fillna(0.0)
    result["source_prev"] = pd.to_numeric(result["source_prev"], errors="coerce").fillna(0.0)
    result["penetration_ratio"] = pd.to_numeric(result["penetration_ratio"], errors="coerce")
    result["ratio_count"] = pd.to_numeric(result["ratio_count"], errors="coerce").fillna(0).astype(int)
    result["calc_current"] = result["source_current"] * result["penetration_ratio"].fillna(0.0)
    result["calc_prev"] = result["source_prev"] * result["penetration_ratio"].fillna(0.0)
    return result[
        [
            "dimension_value",
            "legal_company",
            "penetration_ratio",
            "ratio_count",
            "source_current",
            "source_prev",
            "calc_current",
            "calc_prev",
        ]
    ]


def compare_professional(base) -> pd.DataFrame:
    """专业公司报表与专业公司的回款营收比调整项台账对比。"""
    metric_current = base.u(r"\u5355\u4e00\u4e1a\u6743\u5f53\u5e74\u5e94\u6536\u5728\u5f53\u671f\u672a\u5230\u8d26\u671f\u91d1\u989d")
    metric_prev = base.u(r"\u5355\u4e00\u4e1a\u6743\u4e0a\u4e00\u5e74\u672b\u5e94\u6536\u672a\u5230\u8d26\u671f\u91d1\u989d")
    report = load_professional_report(base)[["dimension_value", metric_current, metric_prev]].rename(
        columns={metric_current: "report_current", metric_prev: "report_prev"}
    )
    calc = load_professional_source(base)
    compare = report.merge(calc, on="dimension_value", how="left").fillna(0.0)
    compare["diff_current"] = compare["calc_current"] - compare["report_current"]
    compare["diff_prev"] = compare["calc_prev"] - compare["report_prev"]
    compare["status"] = (
        compare["diff_current"].abs().le(TOLERANCE) & compare["diff_prev"].abs().le(TOLERANCE)
    ).map({True: "passed", False: "failed"})
    return compare


def summarize_compare(label: str, compare: pd.DataFrame) -> dict:
    """对两个目标字段同时输出通过情况。"""
    return {
        "dimension": label,
        "rows": int(len(compare)),
        "status": "passed" if compare["status"].eq("passed").all() else "failed",
        "current_calc_total": float(compare["calc_current"].sum()),
        "current_report_total": float(compare["report_current"].sum()),
        "current_diff_total": float(compare["diff_current"].sum()),
        "prev_calc_total": float(compare["calc_prev"].sum()),
        "prev_report_total": float(compare["report_prev"].sum()),
        "prev_diff_total": float(compare["diff_prev"].sum()),
        "mismatch_rows": int(compare["status"].ne("passed").sum()),
    }


def main() -> None:
    base = load_base_module()
    indicator_rows = load_indicator_rows(base)
    line_compare = compare_line(base)
    professional_compare = compare_professional(base)
    professional_sources = find_professional_source(base)
    indicator_prof_dims = set(
        indicator_rows.loc[
            indicator_rows["cycle"].eq(base.u(r"\u7d2f\u8ba1"))
            & indicator_rows["dimension"].isin(
                [
                    base.u(r"\u5b89\u4fdd"),
                    base.u(r"\u91d1\u4ee4\u91d1\u5320"),
                    base.u(r"\u91d1\u9890"),
                    base.u(r"\u97f5\u6db5"),
                ]
            ),
            "dimension",
        ]
    )

    metric_current = base.u(r"\u5355\u4e00\u4e1a\u6743\u5f53\u5e74\u5e94\u6536\u5728\u5f53\u671f\u672a\u5230\u8d26\u671f\u91d1\u989d")
    metric_prev = base.u(r"\u5355\u4e00\u4e1a\u6743\u4e0a\u4e00\u5e74\u672b\u5e94\u6536\u672a\u5230\u8d26\u671f\u91d1\u989d")

    print("[indicator_rows]")
    print(indicator_rows.to_json(force_ascii=False, orient="records"))

    print("[line_compare_summary]")
    print(pd.DataFrame([summarize_compare("条线", line_compare)]).to_json(force_ascii=False, orient="records"))

    print("[line_compare_detail]")
    print(line_compare.to_json(force_ascii=False, orient="records"))

    print("[professional_source_status]")
    print(
        pd.DataFrame(
            [
                {
                    "status": "available" if professional_sources else "blocked_missing_source",
                    "matched_sources": [path.name for path in professional_sources],
                    "required_source": base.u(r"\u4e13\u4e1a\u516c\u53f8\u7684\u56de\u6b3e\u8425\u6536\u6bd4\u8c03\u6574\u9879\u53f0\u8d26"),
                    "note": ""
                    if professional_sources
                    else base.u(
                        r"\u6307\u6807\u6e05\u5355\u4e13\u4e1a\u516c\u53f8\u7ef4\u5ea6\u8981\u6c42\u4ece\u8be5\u53f0\u8d26\u53d6\u672a\u8fbe\u8d26\u671f\u91d1\u989d\uff0c\u5f53\u524d\u76ee\u5f55\u672a\u627e\u5230\u542b\u76ee\u6807\u5b57\u6bb5\u7684\u6e90\u8868"
                    ),
                }
            ]
        ).to_json(force_ascii=False, orient="records")
    )

    print("[professional_compare_summary]")
    print(
        pd.DataFrame([summarize_compare("专业公司", professional_compare)]).to_json(
            force_ascii=False, orient="records"
        )
    )

    print("[professional_compare_detail]")
    print(
        professional_compare.assign(
            indicator_cumulative_row_exists=lambda df: df["dimension_value"].isin(indicator_prof_dims)
        ).to_json(force_ascii=False, orient="records")
    )


if __name__ == "__main__":
    main()
