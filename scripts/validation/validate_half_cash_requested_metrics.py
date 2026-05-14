from __future__ import annotations

import json

import pandas as pd

from validate_half_cash_attributable_profit import (
    HALF_ATTRIBUTABLE,
    METRIC_COLUMNS,
    REGION,
    REPORT_MONTH,
    REPORT_MONTH_DASHED,
    ROOT,
    TOLERANCE,
    find_workbook,
    load_other_adjust_project,
    load_other_adjust_region,
    load_profit_non_assess_codes,
    load_project_report,
    load_query,
    load_region_report,
    normalize_code,
    print_section,
    summarize_diff,
    u,
)


PROJECT = u(r"\u9879\u76ee")
LINE = u(r"\u6761\u7ebf")
PROFESSIONAL_COMPANY = u(r"\u4e13\u4e1a\u516c\u53f8")
SPACE = u(r"\u7a7a\u95f4")
SINGLE_OWNER = u(r"\u5355\u4e00\u4e1a\u6743")
NON_RELATED = u(r"\u975e\u5173\u8054\u65b9")

# 源台账名称统一用 Unicode 转义构造，避免终端编码影响中文文件名匹配。
EBT_LEDGER = u(r"\u5e74\u5ea6\u5730\u4ea7\u5173\u8054\u65b9EBT\u6210\u672c\u53f0\u8d26")
ALLOCATED_WATER_LEDGER = u(r"\u57ab\u652f\u6c34\u7535\u8d39\u7269\u4e1a\u5206\u644a\u636e\u5b9e\u5206\u644a")
WATER_FEE_UNRECEIVED_LEDGER = u(r"\u6c34\u7535\u8d39\u672a\u5230\u8d26\u671f\u91d1\u989d\u5e74\u5ea6\u5206\u5e03")
RECEIVABLE_AGE_LEDGER = u(r"\u5e94\u6536\u8d26\u9f84\u53ca\u672a\u5230\u8d26\u671f\u91d1\u989d\u5e74\u5ea6\u5206\u5e03")
OVERDUE_BOND_LEDGER = u(r"\u4fdd\u8bc1\u91d1")
LEGAL_ORG_LEDGER = u(r"\u6cd5\u4eba\u4e0e\u7ec4\u7ec7\u5173\u7cfb")
LEGAL_RATIO_LEDGER = u(r"\u9879\u76ee\u6cd5\u4eba\u516c\u53f8\u7a7f\u900f\u6bd4\u4f8b\u914d\u7f6e")
JV_INFO_LEDGER = u(r"\u5408\u8d44\u516c\u53f8")
OPERATING_1_5_1 = "1.5.1-"
YES = u(r"\u662f")
ENABLED = u(r"\u542f\u7528")
WATER_DETAIL_LEDGER = u(r"\u57ab\u652f\u6c34\u7535\u8d39\u660e\u7ec6\u8868")


def month_key(value: object) -> str:
    # 兼容 Excel 日期、202512、2025-12 等月份格式，统一转成 yyyy-mm。
    if pd.isna(value):
        return ""
    if hasattr(value, "strftime"):
        return value.strftime("%Y-%m")
    text = str(value).strip()
    if not text or text.lower() == "nan":
        return ""
    if text.endswith(".0"):
        text = text[:-2]
    if len(text) == 6 and text.isdigit():
        return f"{text[:4]}-{text[4:]}"
    return text[:7]


def as_number(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce").fillna(0.0)


def find_optional_workbook(*tokens: str) -> Path | None:
    matches = [
        path
        for path in ROOT.iterdir()
        if path.suffix.lower() == ".xlsx" and all(token in path.name for token in tokens)
    ]
    if len(matches) == 1:
        return matches[0]
    return None


def valid_text_mask(series: pd.Series) -> pd.Series:
    text = series.astype(str).str.strip()
    return series.notna() & text.ne("") & ~text.isin(["0", "0.0", "nan", "None"])


def normalize_header_text(value: object) -> str:
    return "".join(str(value).strip().split())


def find_column_by_normalized_name(columns: list[str], target: str) -> str | None:
    target_norm = normalize_header_text(target)
    for column in columns:
        if normalize_header_text(column) == target_norm:
            return column
    return None


def load_professional_company_report() -> pd.DataFrame:
    # 专业公司报表尾部有空行，先按专业公司名称过滤，再统一重命名指标列。
    df = pd.read_excel(find_workbook(HALF_ATTRIBUTABLE, PROFESSIONAL_COMPANY), sheet_name=0, dtype=object)
    df.columns = ["professional_company", *METRIC_COLUMNS, "ownership_attr"]
    df = df[valid_text_mask(df["professional_company"])].copy()
    return df.assign(professional_company=lambda data: data["professional_company"].astype(str).str.strip()).pipe(
        lambda data: data.assign(**{column: pd.to_numeric(data[column], errors="coerce").fillna(0.0) for column in METRIC_COLUMNS})
    )


def load_space_report() -> pd.DataFrame:
    # 空间报表只有一条有效汇总行，按指标列是否存在非零值过滤掉尾部脏行。
    df = pd.read_excel(find_workbook(HALF_ATTRIBUTABLE, SPACE), sheet_name=0, dtype=object)
    df.columns = [*METRIC_COLUMNS, "ownership_attr"]
    for column in METRIC_COLUMNS:
        df[column] = pd.to_numeric(df[column], errors="coerce").fillna(0.0)
    metric_mask = df[METRIC_COLUMNS].abs().sum(axis=1) > TOLERANCE
    df = df[metric_mask].copy()
    df["space"] = SPACE
    return df


def nonzero_details(
    df: pd.DataFrame,
    calc_col: str,
    report_col: str,
    columns: list[str],
    limit: int = 30,
) -> list[dict]:
    diff = df[calc_col] - df[report_col]
    cols = columns + [report_col, calc_col]
    result = df.loc[diff.abs() > TOLERANCE, cols].copy()
    result["diff"] = diff.loc[result.index]
    return result.head(limit).to_dict(orient="records")


def compare_professional_company(
    label: str,
    report_df: pd.DataFrame,
    source_df: pd.DataFrame,
    source_value: str,
    report_value: str,
) -> None:
    merged = report_df[["professional_company", report_value]].merge(
        source_df[["professional_company", source_value]],
        on="professional_company",
        how="left",
    )
    merged[source_value] = merged[source_value].fillna(0.0)
    print_section(
        label,
        [summarize_diff(label, merged[source_value], merged[report_value])]
        + nonzero_details(merged, source_value, report_value, ["professional_company"]),
    )


def compare_space_value(label: str, report_df: pd.DataFrame, calc_value: float, report_value: str) -> None:
    result = report_df[["space", report_value]].copy()
    result[f"{report_value}_calc"] = calc_value
    print_section(
        label,
        [
            summarize_diff(label, result[f"{report_value}_calc"], result[report_value]),
            *nonzero_details(result, f"{report_value}_calc", report_value, ["space"]),
        ],
    )


def compare_project(
    label: str,
    project_df: pd.DataFrame,
    source_df: pd.DataFrame,
    source_value: str,
    report_value: str,
) -> None:
    # 项目维度统一按项目编码归并源表金额，再与半收付项目报表列比对。
    merged = project_df[["region", "project_code", "project_name", "code_norm", report_value]].merge(
        source_df[["code_norm", source_value]],
        on="code_norm",
        how="left",
    )
    merged[source_value] = merged[source_value].fillna(0.0)
    print_section(
        label,
        [summarize_diff(label, merged[source_value], merged[report_value])]
        + nonzero_details(
            merged,
            source_value,
            report_value,
            ["region", "project_code", "project_name"],
        ),
    )


def compare_region(
    label: str,
    region_df: pd.DataFrame,
    source_df: pd.DataFrame,
    source_value: str,
    report_value: str,
) -> None:
    # 区域维度统一按区域名对齐源表复算值和区域报表值。
    merged = region_df[["region", report_value]].merge(
        source_df[["region", source_value]],
        on="region",
        how="left",
    )
    merged[source_value] = merged[source_value].fillna(0.0)
    print_section(
        label,
        [summarize_diff(label, merged[source_value], merged[report_value])]
        + nonzero_details(merged, source_value, report_value, ["region"]),
    )


def attach_ratio(df: pd.DataFrame, query_df: pd.DataFrame) -> pd.DataFrame:
    # 项目查询中的穿透比例是项目口径指标的权威比例来源。
    ratio = query_df[["code_norm", "penetration_ratio"]].drop_duplicates("code_norm")
    result = df.merge(ratio, on="code_norm", how="left")
    result["penetration_ratio"] = result["penetration_ratio"].fillna(1.0)
    return result


def load_jka_sources(
    query_df: pd.DataFrame,
    profit_non_assess_codes: set[str],
) -> tuple[pd.DataFrame, pd.DataFrame]:
    # JKA 项目口径取“类型=项目”；区域口径按指标清单为项目利润考核汇总 + 区域级台账。
    df = pd.read_excel(find_workbook(EBT_LEDGER), header=1)
    type_col = next(column for column in df.columns if u(r"\u7c7b\u578b") in str(column))
    month_col = next(column for column in df.columns if u(r"\u6570\u636e\u5e74\u6708") in str(column))
    code_col = next(column for column in df.columns if u(r"\u7acb\u9879\u7f16\u7801") in str(column))
    region_col = next(column for column in df.columns if u(r"\u533a\u57df") in str(column))
    amount_col = next(column for column in df.columns if "JKA" in str(column))

    df = df[df[month_col].map(month_key).eq(REPORT_MONTH_DASHED)].copy()
    df["code_norm"] = df[code_col].map(normalize_code)
    df["amount"] = as_number(df[amount_col]) * 0.81

    project = (
        df[df[type_col].astype(str).str.strip().eq(PROJECT)]
        .groupby("code_norm", as_index=False)["amount"]
        .sum()
        .rename(columns={"amount": "jka_calc"})
    )

    project_region = df[df[type_col].astype(str).str.strip().eq(PROJECT)].copy()
    if region_col not in project_region.columns or project_region[region_col].isna().all():
        project_region = project_region.merge(
            query_df[["code_norm", "region"]].drop_duplicates("code_norm"),
            on="code_norm",
            how="left",
        )
        region_name = "region"
    else:
        region_name = region_col
    project_region = project_region[~project_region["code_norm"].isin(profit_non_assess_codes)]
    project_rollup = (
        project_region.groupby(region_name, as_index=False)["amount"]
        .sum()
        .rename(columns={region_name: "region", "amount": "project_amount"})
    )

    region_manual = (
        df[
            df[type_col].astype(str).str.strip().eq(REGION)
            & df[code_col].isna()
        ]
        .groupby(region_col, as_index=False)["amount"]
        .sum()
        .rename(columns={region_col: "region", "amount": "region_amount"})
    )
    region = project_rollup.merge(region_manual, on="region", how="outer").fillna(0.0)
    region["jka_calc"] = region["project_amount"] + region["region_amount"]
    return project, region[["region", "jka_calc", "project_amount", "region_amount"]]


def load_allocated_water_source(query_df: pd.DataFrame) -> pd.DataFrame:
    # 据实分摊水电：回款取当期/上一年/往年实收；当期应收按指标清单第 2103 行，只取当期应收并乘穿透比例。
    raw = pd.read_excel(find_workbook(ALLOCATED_WATER_LEDGER), header=None)
    df = raw.iloc[5:].reset_index(drop=True).copy()
    df.columns = [f"c{i}" for i in range(df.shape[1])]
    df = df[df["c0"].map(month_key).eq(REPORT_MONTH_DASHED)].copy()
    df["code_norm"] = df["c1"].map(normalize_code)
    df = attach_ratio(df, query_df)
    df["water_backflow_alloc_calc"] = (as_number(df["c5"]) + as_number(df["c7"]) + as_number(df["c9"])) * df[
        "penetration_ratio"
    ]
    df["water_current_receivable_alloc_calc"] = as_number(df["c4"]) * df["penetration_ratio"]
    return df.groupby("code_norm", as_index=False)[
        ["water_backflow_alloc_calc", "water_current_receivable_alloc_calc"]
    ].sum()


def load_region_water_source(
    query_df: pd.DataFrame,
    profit_non_assess_codes: set[str],
) -> tuple[pd.DataFrame, dict | None]:
    # 区域代收水电和剩余应收按用户刚更新的指标原文走 A-B-C：
    # A=区域下全部明细总额
    # B=区域下利润非考核项目总额
    # C=区域下利润考核项目按指定穿透口径折算后的总额
    source_path = find_optional_workbook(WATER_DETAIL_LEDGER)
    if source_path is None:
        return pd.DataFrame(), {
            "status": "blocked",
            "reason": "missing_required_ledger",
            "required_ledger": WATER_DETAIL_LEDGER,
            "note": "当前工作区未找到“垫支水电费明细表”，不能用“垫支水电费物业分摊据实分摊.xlsx”替代正式口径。",
        }
    raw = pd.read_excel(source_path, header=None)
    header = raw.iloc[2].fillna("").astype(str).str.strip().tolist()
    df = raw.iloc[4:].reset_index(drop=True).copy()
    df.columns = header

    region_col = find_column_by_normalized_name(df.columns.tolist(), u(r"\u533a\u57df"))
    code_col = find_column_by_normalized_name(df.columns.tolist(), u(r"\u7acb\u9879/\u6210\u672c\u4e2d\u5fc3\u7f16\u7801"))
    year_col = find_column_by_normalized_name(df.columns.tolist(), u(r"\u5e74\u4efd"))
    current_col = find_column_by_normalized_name(df.columns.tolist(), u(r"\u57ab\u652f\u6c34\u7535\u6b20\u8d39 \u5f53\u671f\u5e94\u6536"))
    history_real_col = find_column_by_normalized_name(df.columns.tolist(), u(r"\u57ab\u652f\u6c34\u7535\u6b20\u8d39 \u5386\u6b20\u5b9e\u6536"))
    current_real_col = find_column_by_normalized_name(df.columns.tolist(), u(r"\u57ab\u652f\u6c34\u7535\u6b20\u8d39 \u5f53\u671f\u5b9e\u6536"))

    resolved_columns = {
        "region": region_col,
        "code": code_col,
        "year": year_col,
        "current_receivable": current_col,
        "history_real": history_real_col,
        "current_real": current_real_col,
    }
    missing_columns = [name for name, column in resolved_columns.items() if column is None]
    if missing_columns:
        return pd.DataFrame(), {
            "status": "blocked",
            "reason": "missing_required_columns",
            "required_ledger": source_path.name,
            "missing_columns": missing_columns,
        }

    df = df[df[year_col].astype(str).str.strip().eq(REPORT_MONTH[:4])].copy()
    if df.empty:
        return pd.DataFrame(), {
            "status": "blocked",
            "reason": "no_rows_for_report_year",
            "required_ledger": source_path.name,
            "report_year": REPORT_MONTH[:4],
        }
    df["region"] = df[region_col].astype(str).str.strip()
    df["code_norm"] = df[code_col].map(normalize_code)

    ratio = query_df[["code_norm", "penetration_ratio"]].drop_duplicates("code_norm")
    query_codes = set(ratio["code_norm"])
    df = df.merge(ratio, on="code_norm", how="left")
    df["penetration_ratio"] = df["penetration_ratio"].fillna(1.0)

    df["current_receivable_raw"] = as_number(df[current_col])
    df["backflow_raw"] = as_number(df[history_real_col]) + as_number(df[current_real_col])
    df["is_query_project"] = df["code_norm"].isin(query_codes)
    df["is_profit_non_assess"] = df["code_norm"].isin(profit_non_assess_codes)
    df["is_assess_project"] = df["is_query_project"] & ~df["is_profit_non_assess"]

    records: list[dict] = []
    for region, group in df.groupby("region", dropna=False):
        current_a = float(group["current_receivable_raw"].sum())
        current_b = float(group.loc[group["is_profit_non_assess"], "current_receivable_raw"].sum())
        current_c = float(
            (group.loc[group["is_assess_project"], "current_receivable_raw"]
             * (1 - group.loc[group["is_assess_project"], "penetration_ratio"])).sum()
        )
        backflow_a = float(group["backflow_raw"].sum())
        backflow_b = float(group.loc[group["is_profit_non_assess"], "backflow_raw"].sum())
        backflow_c = float(
            (group.loc[group["is_assess_project"], "backflow_raw"]
             * (1 - group.loc[group["is_assess_project"], "penetration_ratio"])).sum()
        )
        records.append(
            {
                "region": str(region).strip(),
                "water_remaining_receivable_calc": current_a - current_b - current_c,
                "water_backflow_calc": backflow_a - backflow_b - backflow_c,
                "current_A_total": current_a,
                "current_B_non_assess": current_b,
                "current_C_assess_penetrated": current_c,
                "backflow_A_total": backflow_a,
                "backflow_B_non_assess": backflow_b,
                "backflow_C_assess_unpenetrated": backflow_c,
            }
        )
    return pd.DataFrame(records), None


def load_energy_income_source(query_df: pd.DataFrame) -> pd.DataFrame:
    # 报表列为“能耗收入_半收付*0.06”，因此源表能源服务收入先乘穿透比例，再乘 0.06 比对。
    raw = pd.read_excel(
        find_workbook(OPERATING_1_5_1, u(r"\u7ecf\u8425\u6536\u652f\u8868\u67e5\u8be2\u5e95\u8868")),
        header=None,
    )
    df = raw.iloc[5:].reset_index(drop=True).copy()
    df.columns = [f"c{i}" for i in range(df.shape[1])]
    df = df[df["c1"].astype(str).str.contains(r"\d", na=False)].copy()
    df["code_norm"] = df["c1"].map(normalize_code)
    df = attach_ratio(df, query_df)
    df["energy_income_calc"] = as_number(df["c19"]) * df["penetration_ratio"]
    df["energy_income_tax_calc"] = df["energy_income_calc"] * 0.06
    return df.groupby("code_norm", as_index=False)[["energy_income_calc", "energy_income_tax_calc"]].sum()


def load_region_energy_income_source(
    query_df: pd.DataFrame,
    profit_non_assess_codes: set[str],
) -> pd.DataFrame:
    # 区域能耗收入按指标清单取区域 1.5.1 的 A-B-C；非项目组织行只留在 A，不进入 B/C。
    token = u(r"\u7efc\u7ba1\u533a\u57df\u53e3\u5f84")
    workbooks = sorted(
        path
        for path in ROOT.iterdir()
        if path.suffix.lower() == ".xlsx" and OPERATING_1_5_1 in path.name and token in path.name and "(" in path.name
    )
    ratio = query_df[["code_norm", "penetration_ratio"]].drop_duplicates("code_norm")
    query_codes = set(ratio["code_norm"])
    records = []
    for path in workbooks:
        region = path.name.split("(", 1)[1].split(")", 1)[0].strip()
        raw = pd.read_excel(path, header=None)
        df = raw.iloc[5:].reset_index(drop=True).copy()
        df.columns = [f"c{i}" for i in range(df.shape[1])]
        total_mask = df["c0"].astype(str).str.strip().eq(u(r"\u603b\u8ba1"))
        a_total = float(as_number(df.loc[total_mask, "c19"]).sum())

        detail = df[~total_mask].copy()
        detail["energy_income_raw"] = as_number(detail["c19"])
        detail = detail[detail["energy_income_raw"].abs() > TOLERANCE].copy()
        detail["code_norm"] = detail["c1"].map(normalize_code)
        detail = detail.merge(ratio, on="code_norm", how="left")
        detail["is_query_project"] = detail["code_norm"].isin(query_codes)
        detail["is_profit_non_assess"] = detail["code_norm"].isin(profit_non_assess_codes)

        b_amount = float(detail.loc[detail["is_profit_non_assess"], "energy_income_raw"].sum())
        c_rows = detail[detail["is_query_project"] & ~detail["is_profit_non_assess"]].copy()
        c_amount = float((c_rows["energy_income_raw"] * (1 - c_rows["penetration_ratio"])).sum())
        retained_non_project = float(detail.loc[~detail["is_query_project"], "energy_income_raw"].sum())
        energy_income_calc = a_total - b_amount - c_amount
        records.append(
            {
                "region": region,
                "source_file": path.name,
                "A_total": a_total,
                "B_profit_non_assess": b_amount,
                "C_assess_unpenetrated": c_amount,
                "retained_non_project_in_A": retained_non_project,
                "energy_income_calc": energy_income_calc,
                "energy_income_tax_calc": energy_income_calc * 0.06,
                "nonzero_detail_rows": int(len(detail)),
                "nonzero_project_rows": int(detail["is_query_project"].sum()),
            }
        )
    return pd.DataFrame(records)


def load_overdue_bond_region_source(
    project_df: pd.DataFrame,
    profit_non_assess_codes: set[str],
) -> pd.DataFrame:
    # 区域保证金=A项目指标库汇总（排除利润非考核）+B区域手工台账；预计回款时间晚于报表月的不算逾期。
    ledger = pd.read_excel(find_workbook(OVERDUE_BOND_LEDGER, u(r"\u903e\u671f")), dtype=object)
    type_col = next(column for column in ledger.columns if u(r"\u7c7b\u578b") in str(column))
    month_col = next(column for column in ledger.columns if u(r"\u6570\u636e\u5e74\u6708") in str(column))
    region_col = next(column for column in ledger.columns if u(r"\u533a\u57df") in str(column))
    nature_col = next(column for column in ledger.columns if u(r"\u6b3e\u9879\u6027\u8d28") in str(column))
    amount_col = next(column for column in ledger.columns if u(r"\u903e\u671f\u91d1\u989d") in str(column))
    expected_col = next(column for column in ledger.columns if u(r"\u9884\u8ba1\u56de\u6b3e\u65f6\u95f4") in str(column))

    assessment_project = project_df[~project_df["code_norm"].isin(profit_non_assess_codes)].copy()
    project_part = assessment_project.groupby("region", as_index=False)[
        ["overdue_performance_bond", "overdue_bid_bond"]
    ].sum()

    ledger["expected_month"] = ledger[expected_col].map(month_key)
    ledger = ledger[
        ledger[type_col].astype(str).str.strip().eq(REGION)
        & ledger[month_col].map(month_key).eq(REPORT_MONTH_DASHED)
        & ledger["expected_month"].ne("")
        & (ledger["expected_month"] <= REPORT_MONTH_DASHED)
    ].copy()
    ledger["amount"] = as_number(ledger[amount_col])

    bond_map = {
        u(r"\u5c65\u7ea6\u4fdd\u8bc1\u91d1"): "region_performance_amount",
        u(r"\u6295\u6807\u4fdd\u8bc1\u91d1"): "region_bid_amount",
    }
    region_parts = []
    for nature, output_col in bond_map.items():
        part = (
            ledger[ledger[nature_col].astype(str).str.strip().eq(nature)]
            .groupby(region_col, as_index=False)["amount"]
            .sum()
            .rename(columns={region_col: "region", "amount": output_col})
        )
        region_parts.append(part)

    source = project_part
    for part in region_parts:
        source = source.merge(part, on="region", how="outer")
    source = source.fillna(0.0)
    source["overdue_performance_bond_calc"] = (
        source["overdue_performance_bond"] + source.get("region_performance_amount", 0.0)
    )
    source["overdue_bid_bond_calc"] = source["overdue_bid_bond"] + source.get("region_bid_amount", 0.0)
    return source


def normalize_professional_company(value: object) -> str:
    text = str(value).strip()
    mapping = {
        u(r"\u5a01\u9707\u4fdd\u5b89"): u(r"\u5b89\u4fdd"),
        u(r"\u5b89\u4fdd"): u(r"\u5b89\u4fdd"),
        u(r"\u91d1\u4ee4"): u(r"\u91d1\u4ee4\u91d1\u5320"),
        u(r"\u91d1\u5320"): u(r"\u91d1\u4ee4\u91d1\u5320"),
        u(r"\u91d1\u9890"): u(r"\u91d1\u9890"),
        u(r"\u97f5\u6db5"): u(r"\u97f5\u6db5"),
    }
    return mapping.get(text, "")


def unique_text_mapping(df: pd.DataFrame, key_col: str, value_col: str) -> tuple[pd.DataFrame, list[dict]]:
    records: list[dict] = []
    blocked: list[dict] = []
    for key, group in df.groupby(key_col):
        values = sorted({str(value).strip() for value in group[value_col].dropna() if str(value).strip()})
        if len(values) == 1:
            records.append({key_col: key, value_col: values[0]})
        elif len(values) > 1:
            blocked.append({key_col: key, value_col: values})
    return pd.DataFrame(records), blocked


def unique_ratio_mapping(df: pd.DataFrame, key_col: str, value_col: str) -> tuple[pd.DataFrame, list[dict]]:
    records: list[dict] = []
    blocked: list[dict] = []
    for key, group in df.groupby(key_col):
        values = sorted({round(float(value), 10) for value in group[value_col].dropna()})
        if len(values) == 1:
            records.append({key_col: key, value_col: values[0]})
        elif len(values) > 1:
            blocked.append({key_col: key, value_col: values})
    return pd.DataFrame(records), blocked


def load_half_interest_sources(region_df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame, list[dict]]:
    # 半收付资金利息必须按法人组织关系和法人穿透比例逐条复算。
    fund = pd.read_excel(find_workbook(u(r"\u8d44\u91d1\u5229\u606f")), header=1, dtype=object)
    rel = pd.read_excel(find_workbook(LEGAL_ORG_LEDGER), dtype=object)
    ratio = pd.read_excel(find_workbook(LEGAL_RATIO_LEDGER), dtype=object)
    jv = pd.read_excel(find_workbook(JV_INFO_LEDGER), header=1, dtype=object)

    fund = fund.rename(columns={fund.columns[0]: "period", fund.columns[1]: "legal", fund.columns[2]: "amount_tax"})
    rel = rel.rename(columns={rel.columns[0]: "legal", rel.columns[1]: "org_name", rel.columns[2]: "primary_flag", rel.columns[3]: "enabled_flag"})
    ratio = ratio.rename(columns={ratio.columns[0]: "legal", ratio.columns[2]: "penetration_ratio"})
    jv = jv.rename(columns={jv.columns[2]: "legal", jv.columns[3]: "jv_tag"})

    fund = fund[fund["period"].map(month_key).eq(REPORT_MONTH_DASHED)].copy()
    fund["amount_tax"] = as_number(fund["amount_tax"])

    # 资金利息口径已确认：不过滤“未启用”，只按“是否主要=是”取组织归属。
    rel = rel[
        rel["primary_flag"].astype(str).str.strip().eq(YES)
    ][["legal", "org_name"]].copy()
    rel["org_name"] = rel["org_name"].astype(str).str.strip()
    rel_map, rel_ambiguous = unique_text_mapping(rel, "legal", "org_name")

    ratio["penetration_ratio"] = pd.to_numeric(ratio["penetration_ratio"], errors="coerce")
    ratio_map, ratio_ambiguous = unique_ratio_mapping(ratio[["legal", "penetration_ratio"]], "legal", "penetration_ratio")

    jv["jv_tag"] = jv["jv_tag"].astype(str).str.strip()
    jv_map, jv_ambiguous = unique_text_mapping(jv[["legal", "jv_tag"]], "legal", "jv_tag")

    merged = fund.merge(rel_map, on="legal", how="left").merge(ratio_map, on="legal", how="left").merge(jv_map, on="legal", how="left")
    merged = merged[~merged["jv_tag"].astype(str).isin(["C", "D"])].copy()
    merged["interest_calc"] = merged["amount_tax"] / 1.06 * 0.81 * merged["penetration_ratio"]
    merged["professional_company"] = merged["org_name"].map(normalize_professional_company)

    valid_regions = set(region_df["region"].dropna().astype(str).str.strip())
    region_rows = merged[merged["org_name"].astype(str).isin(valid_regions)].copy()
    prof_rows = merged[merged["professional_company"].ne("")].copy()

    diagnostics: list[dict] = []
    if rel_ambiguous:
        diagnostics.append({"type": "ambiguous_org_mapping", "rows": rel_ambiguous[:30]})
    if ratio_ambiguous:
        diagnostics.append({"type": "ambiguous_penetration_ratio", "rows": ratio_ambiguous[:30]})
    if jv_ambiguous:
        diagnostics.append({"type": "ambiguous_jv_tag", "rows": jv_ambiguous[:30]})

    missing_ratio = merged[
        (merged["org_name"].astype(str).isin(valid_regions) | merged["professional_company"].ne(""))
        & merged["penetration_ratio"].isna()
    ][["legal", "org_name", "professional_company", "amount_tax"]].copy()
    if not missing_ratio.empty:
        diagnostics.append({"type": "missing_penetration_ratio", "rows": missing_ratio.head(30).to_dict(orient="records")})

    missing_org = merged[merged["org_name"].isna()][["legal", "amount_tax"]].copy()
    if not missing_org.empty:
        diagnostics.append({"type": "missing_org_mapping", "rows": missing_org.head(30).to_dict(orient="records")})

    region_source = (
        region_rows.dropna(subset=["penetration_ratio"])
        .groupby("org_name", as_index=False)["interest_calc"]
        .sum()
        .rename(columns={"org_name": "region"})
    )
    prof_source = (
        prof_rows.dropna(subset=["penetration_ratio"])
        .groupby("professional_company", as_index=False)["interest_calc"]
        .sum()
    )
    return region_source, prof_source, diagnostics


def load_single_owner_water_sources(
    query_df: pd.DataFrame,
    profit_non_assess_codes: set[str],
) -> tuple[pd.DataFrame, pd.DataFrame]:
    # 单一业权水电未到账期：项目取单一业权项目行；区域按指标清单为区域行 A + 项目行 B。
    # 此处使用指标清单指定的水电费未到账期台账，不使用“大业主应收账龄”台账。
    df = pd.read_excel(find_workbook(WATER_FEE_UNRECEIVED_LEDGER), header=None)
    df = df.iloc[2:].reset_index(drop=True).copy()
    df.columns = [f"c{i}" for i in range(df.shape[1])]
    df = df[df["c0"].map(month_key).eq(REPORT_MONTH_DASHED)].copy()
    df["type"] = df["c1"].astype(str).str.strip()
    df["region"] = df["c2"].astype(str).str.strip()
    df["code_norm"] = df["c3"].map(normalize_code)
    df["ownership"] = df["c6"].astype(str).str.strip()
    df["current_unreceived_raw"] = as_number(df["c7"])
    df["prev_unreceived_raw"] = as_number(df["c8"])

    project_rows = df[(df["type"].eq(PROJECT)) & (df["ownership"].eq(SINGLE_OWNER))].copy()
    project_rows = attach_ratio(project_rows, query_df)
    project_rows["single_current_unreceived_calc"] = (
        project_rows["current_unreceived_raw"] * project_rows["penetration_ratio"]
    )
    project_rows["single_prev_unreceived_calc"] = (
        project_rows["prev_unreceived_raw"] * project_rows["penetration_ratio"]
    )
    project = project_rows.groupby("code_norm", as_index=False)[
        ["single_current_unreceived_calc", "single_prev_unreceived_calc"]
    ].sum()

    region_manual = df[df["type"].eq(REGION)].groupby("region", as_index=False)[
        ["current_unreceived_raw", "prev_unreceived_raw"]
    ].sum()
    region_manual = region_manual.rename(
        columns={
            "current_unreceived_raw": "region_current_amount",
            "prev_unreceived_raw": "region_prev_amount",
        }
    )

    project_for_region = project_rows[
        ~project_rows["code_norm"].isin(profit_non_assess_codes)
    ].copy()
    project_rollup = project_for_region.groupby("region", as_index=False)[
        ["single_current_unreceived_calc", "single_prev_unreceived_calc"]
    ].sum()
    region = project_rollup.merge(region_manual, on="region", how="outer").fillna(0.0)
    region["single_current_unreceived_calc"] = (
        region["single_current_unreceived_calc"] + region["region_current_amount"]
    )
    region["single_prev_unreceived_calc"] = (
        region["single_prev_unreceived_calc"] + region["region_prev_amount"]
    )
    return project, region[
        [
            "region",
            "single_current_unreceived_calc",
            "single_prev_unreceived_calc",
            "region_current_amount",
            "region_prev_amount",
        ]
    ]


def load_receivable_age_ledger(period: str) -> pd.DataFrame:
    # 应收账龄台账有两行表头：第 1 行为大类，第 2 行为年度分布；数据从第 3 行开始。
    df = pd.read_excel(find_workbook(RECEIVABLE_AGE_LEDGER, period), header=None)
    df = df.iloc[2:].reset_index(drop=True).copy()
    df.columns = [f"c{i}" for i in range(df.shape[1])]
    df = df[df["c1"].map(month_key).eq(f"{period[:4]}-{period[4:]}")].copy()
    df["type"] = df["c2"].astype(str).str.strip()
    df["region"] = df["c3"].astype(str).str.strip()
    df["code_norm"] = df["c4"].map(normalize_code)
    df["customer"] = df["c6"].astype(str).str.strip()
    df["ownership"] = df["c7"].astype(str).str.strip()
    df["ownership_check"] = df["c20"].astype(str).str.strip()
    df["unreceived_total"] = as_number(df["c14"])
    df["unreceived_current_year"] = as_number(df["c15"])
    return df


def load_single_owner_receivable_sources(query_df: pd.DataFrame) -> pd.DataFrame:
    # 单一业权当年应收在当期未到账期金额：取报表月台账“未到账期金额年度分布：当年”。
    # 单一业权上一年末应收未到账期金额：取报表年的上一年 12 月台账“其中未到账期金额”。
    previous_december = f"{int(REPORT_MONTH[:4]) - 1}12"
    current_df = load_receivable_age_ledger(REPORT_MONTH)
    previous_df = load_receivable_age_ledger(previous_december)

    def project_values(df: pd.DataFrame, source_col: str, result_col: str) -> pd.DataFrame:
        rows = df[
            df["type"].eq(PROJECT)
            & (df["ownership"].eq(SINGLE_OWNER) | df["ownership_check"].eq(SINGLE_OWNER))
        ].copy()
        rows = attach_ratio(rows, query_df)
        rows[result_col] = rows[source_col] * rows["penetration_ratio"]
        return rows.groupby("code_norm", as_index=False)[result_col].sum()

    current = project_values(current_df, "unreceived_current_year", "single_current_unreceived_2_calc")
    previous = project_values(previous_df, "unreceived_total", "single_prev_unreceived_2_calc")
    return current.merge(previous, on="code_norm", how="outer").fillna(0.0)


def project_rollup_excluding_profit_non_assess(
    project_df: pd.DataFrame,
    value_cols: list[str],
    profit_non_assess_codes: set[str],
) -> pd.DataFrame:
    # 多数区域指标由项目指标库汇总，并要求排除利润非考核项目。
    filtered = project_df[~project_df["code_norm"].isin(profit_non_assess_codes)].copy()
    return filtered.groupby("region", as_index=False)[value_cols].sum()


def emit_rollup_check(
    label: str,
    project_df: pd.DataFrame,
    region_df: pd.DataFrame,
    value_col: str,
    profit_non_assess_codes: set[str],
) -> None:
    rollup = project_rollup_excluding_profit_non_assess(project_df, [value_col], profit_non_assess_codes)
    compare_region(label, region_df, rollup.rename(columns={value_col: f"{value_col}_rollup"}), f"{value_col}_rollup", value_col)


def main() -> None:
    project_df = load_project_report()
    region_df = load_region_report()
    # 区域报表尾部有空行/无效行，参与汇总前先剔除。
    region_df = region_df[
        region_df["region"].notna()
        & region_df["region"].astype(str).str.strip().ne("")
        & ~region_df["region"].astype(str).str.strip().isin(["0", "0.0"])
    ].copy()
    professional_company_df = load_professional_company_report()
    space_df = load_space_report()
    query_df = load_query()
    profit_non_assess_codes = load_profit_non_assess_codes()

    jka_project, jka_region = load_jka_sources(query_df, profit_non_assess_codes)
    water_source = load_allocated_water_source(query_df)
    region_water_source, region_water_block = load_region_water_source(query_df, profit_non_assess_codes)
    energy_source = load_energy_income_source(query_df)
    region_energy_source = load_region_energy_income_source(query_df, profit_non_assess_codes)
    single_project, single_region = load_single_owner_water_sources(query_df, profit_non_assess_codes)
    receivable_source = load_single_owner_receivable_sources(query_df)
    overdue_bond_region = load_overdue_bond_region_source(project_df, profit_non_assess_codes)

    print_section(
        "requested_metric_scope",
        [
            {
                "report_month": REPORT_MONTH,
                "project_rows": int(len(project_df)),
                "region_rows": int(len(region_df)),
                "professional_company_rows": int(len(professional_company_df)),
                "space_rows": int(len(space_df)),
                "profit_non_assess_codes": int(len(profit_non_assess_codes)),
                "rules": [
                    "project checks use source ledgers and penetration_ratio where required",
                    "region checks use project rollup excluding profit-non-assessment projects plus regional ledger rows where the indicator row requires A+B",
                    "D-exit exclusion is not applied to these final project/region rows unless the exact indicator row says non-D; these requested final rows say profit assessment for region rollup",
                ],
            }
        ],
    )

    compare_project("project_jka_2023_adjust", project_df, jka_project, "jka_calc", "jka_2023_adjust")
    compare_region("region_jka_2023_adjust", region_df, jka_region, "jka_calc", "jka_2023_adjust")

    other_project = project_df[["region", "project_code", "project_name", "code_norm", "other_adjust"]].merge(
        load_other_adjust_project().rename(columns={"amount": "other_adjust_calc"}),
        on="code_norm",
        how="left",
    )
    other_project["other_adjust_calc"] = other_project["other_adjust_calc"].fillna(0.0)
    print_section(
        "project_other_assessment_half_cash_parent_profit",
        [summarize_diff("project_other_assessment_half_cash_parent_profit", other_project["other_adjust_calc"], other_project["other_adjust"])]
        + nonzero_details(
            other_project,
            "other_adjust_calc",
            "other_adjust",
            ["region", "project_code", "project_name"],
        ),
    )

    other_project_rollup = project_rollup_excluding_profit_non_assess(
        other_project.rename(columns={"other_adjust_calc": "amount"}),
        ["amount"],
        profit_non_assess_codes,
    ).rename(columns={"amount": "project_amount"})
    other_region_manual = load_other_adjust_region().rename(columns={"amount": "region_amount"})
    other_region = other_project_rollup.merge(other_region_manual, on="region", how="outer").fillna(0.0)
    other_region["other_adjust_calc"] = other_region["project_amount"] + other_region["region_amount"]
    print_section(
        "region_other_assessment_half_cash_parent_profit",
        region_df[["region", "other_adjust"]]
        .merge(other_region, on="region", how="left")
        .fillna(0.0)
        .assign(diff=lambda df: df["other_adjust_calc"] - df["other_adjust"])
        .assign(status=lambda df: df["diff"].abs().le(TOLERANCE).map({True: "passed", False: "failed"}))
        .to_dict(orient="records"),
    )

    for source_col, report_col, label in [
        ("water_backflow_alloc_calc", "water_backflow_alloc", "allocated_water_backflow"),
        ("water_current_receivable_alloc_calc", "water_current_receivable_alloc", "allocated_water_current_receivable"),
    ]:
        compare_project(f"project_{label}", project_df, water_source, source_col, report_col)
        source_region = (
            project_df[["region", "code_norm"]]
            .merge(water_source[["code_norm", source_col]], on="code_norm", how="left")
            .fillna({source_col: 0.0})
        )
        source_region = project_rollup_excluding_profit_non_assess(
            source_region,
            [source_col],
            profit_non_assess_codes,
        )
        compare_region(f"region_{label}", region_df, source_region, source_col, report_col)

    if region_water_block is not None:
        print_section("region_water_source_blocked", [region_water_block])
    else:
        print_section("region_water_source_breakdown", region_water_source.to_dict(orient="records"))
        compare_region(
            "region_water_backflow",
            region_df,
            region_water_source,
            "water_backflow_calc",
            "water_backflow",
        )
        compare_region(
            "region_water_remaining_receivable",
            region_df,
            region_water_source,
            "water_remaining_receivable_calc",
            "water_remaining_receivable",
        )

    compare_project("project_energy_income_tax", project_df, energy_source, "energy_income_tax_calc", "energy_income_tax")
    if not region_energy_source.empty:
        print_section("region_energy_income_source_breakdown", region_energy_source.to_dict(orient="records"))
        energy_regions = region_energy_source["region"].dropna().astype(str).unique()
        compare_region(
            "region_energy_income_tax",
            region_df[region_df["region"].astype(str).isin(energy_regions)].copy(),
            region_energy_source,
            "energy_income_tax_calc",
            "energy_income_tax",
        )

    compare_region(
        "region_overdue_performance_bond",
        region_df,
        overdue_bond_region,
        "overdue_performance_bond_calc",
        "overdue_performance_bond",
    )
    compare_region(
        "region_overdue_bid_bond",
        region_df,
        overdue_bond_region,
        "overdue_bid_bond_calc",
        "overdue_bid_bond",
    )

    for source_col, report_col, label in [
        ("single_current_unreceived_calc", "single_current_unreceived", "single_owner_current_water_unreceived"),
        ("single_prev_unreceived_calc", "single_prev_unreceived", "single_owner_previous_water_unreceived"),
    ]:
        compare_project(f"project_{label}", project_df, single_project, source_col, report_col)
        compare_region(f"region_{label}", region_df, single_region, source_col, report_col)

    for source_col, report_col, label in [
        ("single_current_unreceived_2_calc", "single_current_unreceived_2", "single_owner_current_receivable_unreceived"),
        ("single_prev_unreceived_2_calc", "single_prev_unreceived_2", "single_owner_previous_year_end_unreceived"),
    ]:
        compare_project(f"project_{label}", project_df, receivable_source, source_col, report_col)
        source_region = (
            project_df[["region", "code_norm"]]
            .merge(receivable_source[["code_norm", source_col]], on="code_norm", how="left")
            .fillna({source_col: 0.0})
        )
        source_region = project_rollup_excluding_profit_non_assess(
            source_region,
            [source_col],
            profit_non_assess_codes,
        )
        compare_region(f"region_{label}", region_df, source_region, source_col, report_col)

    interest_region_source, interest_prof_source, interest_diagnostics = load_half_interest_sources(region_df)
    print_section("half_interest_source_diagnostics", interest_diagnostics or [{"type": "ok", "rows": []}])
    compare_region("region_half_interest", region_df, interest_region_source, "interest_calc", "interest")
    compare_professional_company(
        "professional_company_half_interest",
        professional_company_df,
        interest_prof_source,
        "interest_calc",
        "interest",
    )

    space_interest_calc = float(interest_region_source["interest_calc"].sum() + interest_prof_source["interest_calc"].sum())
    compare_space_value("space_half_interest", space_df, space_interest_calc, "interest")

    # 空间利润非考核项目半收付净利润按指标清单直接汇总区域类型结果，不从项目台账重算。
    space_non_assess_calc = float(region_df["non_assess_half_net_profit"].sum())
    compare_space_value(
        "space_profit_non_assess_half_net_profit",
        space_df,
        space_non_assess_calc,
        "non_assess_half_net_profit",
    )

    print(json.dumps({"done": True}, ensure_ascii=True))


if __name__ == "__main__":
    main()
