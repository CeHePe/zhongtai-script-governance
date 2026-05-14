from __future__ import annotations

from pathlib import Path
from _project_root import find_project_root

import pandas as pd


ROOT = find_project_root(__file__)


def u(text: str) -> str:
    return text.encode("utf-8").decode("unicode_escape")


METRIC_NAME = u(r"\u4e0a\u4e00\u5e74\u56de\u6b3e\u7387")
PROJECT_MARK = u(r"\u9879\u76ee")

REPORT_FILE = u(r"\u4e0a\u4e00\u5e74\u56de\u6b3e\u7387202512\u9879\u76ee.xlsx")
PROJECT_QUERY_FILE = u(r"\u9879\u76ee\u67e5\u8be2.xlsx")
NON_ASSESS_FILE = u(r"\u975e\u8003\u6838\u9879\u76ee\u53f0\u8d26.xlsx")
RELATED_PARTY_FILE = u(r"\u5173\u8054\u65b9\u6c34\u7535\u8d39\u5e94\u6536\u53ca\u5b9e\u6536\u5e74\u5ea6\u5206\u5e03.xlsx")
AGING_2024_FILE = u(r"\u5e94\u6536\u8d26\u9f84\u53ca\u672a\u5230\u8d26\u671f\u91d1\u989d\u5e74\u5ea6\u5206\u5e03202412.xlsx")
AGING_2025_FILE = u(r"\u5e94\u6536\u8d26\u9f84\u53ca\u672a\u5230\u8d26\u671f\u91d1\u989d\u5e74\u5ea6\u5206\u5e03202512.xlsx")
BUSINESS_AGING_2024_FILE = u(r"\u4e1a\u52a1\u5e10\u9f84-\u5e74\u5ea6\u5206\u5e03202412.xlsx")
BUSINESS_AGING_2025_FILE = u(r"\u4e1a\u52a1\u5e10\u9f84-\u5e74\u5ea6\u5206\u5e03202512.xlsx")
COIN_BALANCE_2024_FILE = u(r"\u91d1\u5e01\u4f59\u989d\u53f0\u8d262024.xlsx")

SKIPPED_COMPONENTS = [
    (u(r"\u4e0a\u4e00\u5e74\u4ee3\u6536\u6c34\u8d39\u5e94\u6536\u5728\u5f53\u671f\u56de\u6b3e"), u(r"\u57ab\u652f\u6c34\u7535\u8d39\u660e\u7ec6\u8868"), "2025-12"),
    (u(r"\u4e0a\u4e00\u5e74\u4ee3\u6536\u6c34\u7535\u8d39\u5269\u4f59\u5e94\u6536"), u(r"\u57ab\u652f\u6c34\u7535\u8d39\u660e\u7ec6\u8868"), "2025-12"),
]

ASSUMED_ZERO_COMPONENTS = [
    (u(r"\u5c0f\u4e1a\u4e3b\u91d1\u5e01\u56de\u6b3e\u91d1\u989d_\u5f53\u5e74"), u(r"\u5c0f\u4e1a\u4e3b\u91d1\u5e01\u56de\u6b3e\u91d1\u989d\u53f0\u8d26"), "2025-12"),
]

RELATED_PARTY_EXCLUDE_NAMES = {
    u(r"JKA\u975e\u5168\u8d44\u5b50\u516c\u53f8"),
    u(r"JKA\u5408\u8425\u516c\u53f8"),
    u(r"JKA\u8054\u8425\u516c\u53f8"),
    u(r"JKA\u5168\u8d44\u5b50\u516c\u53f8"),
    u(r"JKS\u5168\u8d44\u5b50\u516c\u53f8"),
    u(r"JKS\u8054\u8425\u516c\u53f8"),
    u(r"JKS\u975e\u5168\u8d44\u5b50\u516c\u53f8"),
    u(r"JKS\u5408\u8425\u516c\u53f8"),
}


def normalize_code(value: object) -> str:
    if value is None or pd.isna(value):
        return ""
    text = str(value).strip().upper()
    if text in {"", "NAN", "NONE"}:
        return ""
    if len(text) >= 2 and text[0].isalpha() and any(ch.isdigit() for ch in text[1:]):
        return text[1:]
    return text


def load_indicator_rows() -> pd.DataFrame:
    indicator = next(path for path in ROOT.glob("*.xlsx") if path.name.startswith("JKS_"))
    df = pd.read_excel(indicator, sheet_name=0)
    df = df.rename(columns={column: str(column).replace(" ", "").strip() for column in df.columns})
    name_col = u(r"\u6307\u6807\u540d\u79f0")
    dim_col = u(r"\u7ec4\u7ec7\u7ef4\u5ea6")
    targets = [
        METRIC_NAME,
        METRIC_NAME + u(r"_\u5206\u5b50"),
        METRIC_NAME + u(r"_\u5206\u6bcd"),
    ]
    return df[df[name_col].isin(targets) & (df[dim_col] == PROJECT_MARK)].copy()


def load_report_df() -> pd.DataFrame:
    df = pd.read_excel(ROOT / REPORT_FILE)
    header_row = df.iloc[0].fillna("").astype(str).str.strip().tolist()
    rename_map = {
        u(r"\u4e0a\u4e00\u5e74\u5e94\u6536\u5728\u5f53\u671f\u56de\u6b3e"): "report_prev_ar_recovery",
        u(r"\u4e0a\u4e00\u5e74\u4ee3\u6536\u6c34\u8d39\u5e94\u6536\u5728\u5f53\u671f\u56de\u6b3e"): "report_water_recv",
        u(r"\u4e0a\u4e00\u5e74\u4ee3\u6536\u6c34\u7535\u56de\u6b3e\u636e\u5b9e\u5206\u644a"): "report_water_recv_alloc",
        u(r"\u5173\u8054\u65b9\u4ee3\u6536\u6c34\u7535\u5386\u6b20\u56de\u6b3e"): "report_rel_paid",
        u(r"\u4e0a\u4e00\u5e74\u672a\u5230\u8d26\u671f\u5728\u5f53\u671f\u56de\u6b3e"): "report_notdue_last_current",
        u(r"\u5f80\u5e74\u672a\u5230\u8d26\u671f\u5728\u5f53\u671f\u56de\u6b3e"): "report_notdue_old_current",
        u(r"\u91d1\u5e01\u56de\u6b3e\u91d1\u989d\uff08\u4e0a\u4e00\u5e74\uff09"): "report_coin_recovery",
        u(r"\u4e0a\u4e00\u5e74\u5269\u4f59\u5e94\u6536"): "report_prev_ar_balance",
        u(r"\u4e0a\u4e00\u5e74\u4ee3\u6536\u6c34\u7535\u8d39\u5269\u4f59\u5e94\u6536"): "report_water_balance",
        u(r"\u4e0a\u4e00\u5e74\u4ee3\u6536\u6c34\u7535\u8d39\u5269\u4f59\u5e94\u6536\u636e\u5b9e\u5206\u644a"): "report_water_balance_alloc",
        u(r"\u5173\u8054\u65b9\u4ee3\u6536\u6c34\u7535\u5386\u6b20\u5e94\u6536"): "report_rel_recv",
        u(r"\u4e0a\u4e00\u5e74\u91d1\u5e01\uff08\u91d1\u5e01\u4f59\u989d\uff09"): "report_coin_balance",
        u(r"\u4e0a\u4e00\u5e74\u672a\u5230\u8d26\u671f\u5728\u4e0a\u4e00\u5e74\u672b\u91d1\u989d"): "report_notdue_last_end",
        u(r"\u5f80\u5e74\u672a\u5230\u8d26\u671f\u5728\u4e0a\u4e00\u5e74\u672b\u91d1\u989d"): "report_notdue_old_end",
    }
    column_rename = {}
    for column in df.columns:
        label = header_row[df.columns.get_loc(column)]
        if label in rename_map:
            column_rename[column] = rename_map[label]
    df = df.rename(columns=column_rename)
    df = df[df[u(r"\u7acb\u9879\u7f16\u7801")].notna()].copy()
    df["code_norm"] = df[u(r"\u7acb\u9879\u7f16\u7801")].map(normalize_code)
    numeric_cols = [
        "report_prev_ar_recovery",
        "report_water_recv",
        "report_water_recv_alloc",
        "report_rel_paid",
        "report_notdue_last_current",
        "report_notdue_old_current",
        "report_coin_recovery",
        u(r"\u4e0a\u4e00\u5e74\u56de\u6b3e\u7387\u5206\u5b50"),
        "report_prev_ar_balance",
        "report_water_balance",
        "report_water_balance_alloc",
        "report_rel_recv",
        "report_coin_balance",
        "report_notdue_last_end",
        "report_notdue_old_end",
        u(r"\u4e0a\u4e00\u5e74\u56de\u6b3e\u7387\u5206\u6bcd"),
        u(r"\u4e0a\u4e00\u5e74\u56de\u6b3e\u7387"),
    ]
    for column in numeric_cols:
        if column in df.columns:
            df[column] = pd.to_numeric(df[column], errors="coerce").fillna(0.0)
    return df


def load_aging_df(filename: str) -> pd.DataFrame:
    df = pd.read_excel(ROOT / filename, header=None).iloc[2:].reset_index(drop=True)
    df.columns = [
        "idx",
        "data_month",
        "type_name",
        "region",
        "project_code",
        "project_name",
        "customer_attr",
        "ownership_attr",
        "big_owner_total",
        "big_owner_current",
        "big_owner_prev1",
        "big_owner_prev2",
        "big_owner_prev3",
        "big_owner_prev4",
        "not_due_total",
        "not_due_current",
        "not_due_prev1",
        "not_due_prev2",
        "not_due_prev3",
        "not_due_prev4",
        "ownership_attr_dup",
        "assessment_flag",
    ]
    df = df[
        (df["type_name"].astype(str) == PROJECT_MARK)
        & (df["customer_attr"].astype(str) == u(r"\u975e\u5173\u8054\u65b9"))
    ].copy()
    df["code_norm"] = df["project_code"].map(normalize_code)
    for column in ["not_due_current", "not_due_prev1", "not_due_prev2", "not_due_prev3", "not_due_prev4"]:
        df[column] = pd.to_numeric(df[column], errors="coerce").fillna(0.0)
    return (
        df.groupby("code_norm", as_index=False)
        .agg(
            not_due_current=("not_due_current", "sum"),
            not_due_prev1=("not_due_prev1", "sum"),
            not_due_prev2=("not_due_prev2", "sum"),
            not_due_prev3=("not_due_prev3", "sum"),
            not_due_prev4=("not_due_prev4", "sum"),
        )
    )


def load_related_party_df() -> pd.DataFrame:
    df = pd.read_excel(ROOT / RELATED_PARTY_FILE, header=None).iloc[2:].reset_index(drop=True)
    df.columns = [
        "idx",
        "data_month",
        "project_code",
        "project_name",
        "region",
        "hist_recv",
        "hist_paid",
        "curr_recv",
        "curr_paid",
    ]
    df = df[df["data_month"].astype(str) == "2025-12"].copy()
    df["code_norm"] = df["project_code"].map(normalize_code)
    for column in ["hist_recv", "hist_paid"]:
        df[column] = pd.to_numeric(df[column], errors="coerce").fillna(0.0)
    return df.groupby("code_norm", as_index=False).agg(rel_recv=("hist_recv", "sum"), rel_paid=("hist_paid", "sum"))


def load_business_aging_df(filename: str) -> pd.DataFrame:
    df = pd.read_excel(ROOT / filename, header=None).iloc[4:].reset_index(drop=True)
    df.columns = [
        "region",
        "project_code",
        "project_name",
        "customer_code",
        "customer_name",
        "manage_area",
        "assessment_attr",
        "project_status",
        "enter_date",
        "exit_date",
        "acquisition_mode",
        "ar_total",
        "year_current",
        "year_prev1",
        "year_prev2",
        "year_prev3",
        "year_prev4_plus",
        "subtotal",
        "water_current",
        "water_prev1",
        "water_prev2",
        "water_prev3",
        "water_prev4_plus",
    ]
    df["code_norm"] = df["project_code"].map(normalize_code)
    df["customer_name"] = df["customer_name"].astype(str).str.strip()
    for column in ["ar_total", "year_current", "year_prev1", "year_prev2", "year_prev3", "year_prev4_plus"]:
        df[column] = pd.to_numeric(df[column], errors="coerce").fillna(0.0)
    return df


def load_coin_balance_df() -> pd.DataFrame:
    df = pd.read_excel(ROOT / COIN_BALANCE_2024_FILE)
    code_col = u(r"\u7acb\u9879\u7f16\u7801")
    amount_col = u(r"\u5e74\u5ea6\u5e94\u8d60\u9001\u91d1\u989d\uff08\u5355\u4f4d\uff1a\u5143\uff09")
    year_col = u(r"\u5e74\u4efd")
    df["code_norm"] = df[code_col].map(normalize_code)
    df[amount_col] = pd.to_numeric(df[amount_col], errors="coerce").fillna(0.0)
    year_mask = df[year_col].astype(str).str.replace(".0", "", regex=False).eq("2024")
    return df.loc[year_mask].groupby("code_norm", as_index=False).agg(report_col_15=(amount_col, "sum"))


def build_component_source_df() -> pd.DataFrame:
    aging_2024 = load_aging_df(AGING_2024_FILE)
    aging_2025 = load_aging_df(AGING_2025_FILE)
    aging = aging_2024.merge(aging_2025, on="code_norm", how="outer", suffixes=("_2024", "_2025")).fillna(0.0)
    aging["report_col_8"] = aging["not_due_current_2024"] - aging["not_due_prev1_2025"]
    aging["report_col_9"] = (
        aging["not_due_prev1_2024"]
        + aging["not_due_prev2_2024"]
        + aging["not_due_prev3_2024"]
        + aging["not_due_prev4_2024"]
        - aging["not_due_prev2_2025"]
        - aging["not_due_prev3_2025"]
        - aging["not_due_prev4_2025"]
    )
    aging["report_col_16"] = aging["not_due_current_2024"]
    aging["report_col_17"] = (
        aging["not_due_prev1_2024"]
        + aging["not_due_prev2_2024"]
        + aging["not_due_prev3_2024"]
        + aging["not_due_prev4_2024"]
    )

    related_party = load_related_party_df()
    business_aging_2024 = load_business_aging_df(BUSINESS_AGING_2024_FILE)
    business_aging_2025 = load_business_aging_df(BUSINESS_AGING_2025_FILE)
    business_aging_2024 = business_aging_2024.loc[
        ~business_aging_2024["customer_name"].isin(RELATED_PARTY_EXCLUDE_NAMES)
    ]
    business_aging_2025 = business_aging_2025.loc[
        ~business_aging_2025["customer_name"].isin(RELATED_PARTY_EXCLUDE_NAMES)
    ]
    business_aging_2024 = business_aging_2024.groupby("code_norm", as_index=False).agg(
        report_col_5=("year_current", "sum")
    )
    business_aging_2025 = business_aging_2025.groupby("code_norm", as_index=False).agg(
        year_prev1_2025=("year_prev1", "sum")
    )
    business_aging = business_aging_2024.merge(business_aging_2025, on="code_norm", how="outer").fillna(0.0)
    business_aging["report_col_12"] = business_aging["report_col_5"]
    business_aging["report_col_5"] = business_aging["report_col_5"] - business_aging["year_prev1_2025"]

    coin_balance = load_coin_balance_df()

    source_df = (
        aging.merge(related_party, on="code_norm", how="outer")
        .merge(business_aging[["code_norm", "report_col_5", "report_col_12"]], on="code_norm", how="outer")
        .merge(coin_balance, on="code_norm", how="outer")
        .fillna(0.0)
    )
    source_df["report_col_10"] = 0.0
    return source_df


def print_indicator_rows(df: pd.DataFrame) -> None:
    acc_col = u(r"\u7d2f\u8ba1/\u6708\u5ea6")
    method_col = u(r"\u53d6\u6570\u65b9\u5f0f")
    source_col = u(r"\u53d6\u6570\u5bf9\u5e94\u8868")
    logic_col = u(r"\u8ba1\u7b97\u903b\u8f91")
    name_col = u(r"\u6307\u6807\u540d\u79f0")
    print("[indicator_rows]")
    for _, row in df.iterrows():
        print(row[name_col])
        print(f"  {acc_col}: {row[acc_col]}")
        print(f"  {method_col}: {row[method_col]}")
        print(f"  {source_col}: {row[source_col]}")
        print(f"  {logic_col}: {str(row[logic_col]).replace(chr(10), ' | ')}")


def print_formula_check(report_df: pd.DataFrame) -> None:
    rate_calc = pd.Series(0.0, index=report_df.index)
    denominator_nonzero = report_df[u(r"\u4e0a\u4e00\u5e74\u56de\u6b3e\u7387\u5206\u6bcd")] != 0
    rate_calc.loc[denominator_nonzero] = (
        report_df.loc[denominator_nonzero, u(r"\u4e0a\u4e00\u5e74\u56de\u6b3e\u7387\u5206\u5b50")]
        / report_df.loc[denominator_nonzero, u(r"\u4e0a\u4e00\u5e74\u56de\u6b3e\u7387\u5206\u6bcd")]
    )
    rate_diff = (rate_calc - report_df[u(r"\u4e0a\u4e00\u5e74\u56de\u6b3e\u7387")]).abs()

    print("[report_formula_check]")
    print(f"rows: {len(report_df)}")
    print("numerator_mismatch_rows: not_checked_due_component_layout_change")
    print("denominator_visible_mismatch_rows: not_checked_due_component_layout_change")
    print(f"rate_mismatch_rows: {int((rate_diff > 1e-9).sum())}")


def compare_component(report_df: pd.DataFrame, source_df: pd.DataFrame, label: str, report_col: str, source_col: str) -> None:
    merged = report_df[[u(r"\u7acb\u9879\u7f16\u7801"), u(r"\u9879\u76ee"), "code_norm", report_col]].merge(
        source_df[["code_norm", source_col]], on="code_norm", how="left"
    )
    merged[source_col] = merged[source_col].fillna(0.0)
    diff = (merged[report_col] - merged[source_col]).abs()
    print(label)
    print(f"  report_total: {float(merged[report_col].sum())}")
    print(f"  source_total: {float(merged[source_col].sum())}")
    print(f"  diff_total: {float(merged[report_col].sum() - merged[source_col].sum())}")
    print(f"  mismatch_rows: {int((diff > 1e-6).sum())}")


def print_available_component_checks(report_df: pd.DataFrame, source_df: pd.DataFrame) -> None:
    print("[available_component_checks]")
    compare_component(report_df, source_df, u(r"\u4e0a\u4e00\u5e74\u5e94\u6536\u5728\u5f53\u671f\u56de\u6b3e"), "report_prev_ar_recovery", "report_col_5")
    compare_component(report_df, source_df, u(r"\u4e0a\u4e00\u5e74\u5269\u4f59\u5e94\u6536"), "report_prev_ar_balance", "report_col_12")
    compare_component(report_df, source_df, u(r"\u5173\u8054\u65b9\u4ee3\u6536\u6c34\u7535\u5386\u6b20\u56de\u6b3e"), "report_rel_paid", "rel_paid")
    compare_component(report_df, source_df, u(r"\u5173\u8054\u65b9\u4ee3\u6536\u6c34\u7535\u5386\u6b20\u5e94\u6536"), "report_rel_recv", "rel_recv")
    compare_component(report_df, source_df, u(r"\u4e0a\u4e00\u5e74\u672a\u5230\u8d26\u671f\u5728\u5f53\u671f\u56de\u6b3e"), "report_notdue_last_current", "report_col_8")
    compare_component(report_df, source_df, u(r"\u5f80\u5e74\u672a\u5230\u8d26\u671f\u5728\u5f53\u671f\u56de\u6b3e"), "report_notdue_old_current", "report_col_9")
    compare_component(report_df, source_df, u(r"\u4e0a\u4e00\u5e74\u672a\u5230\u8d26\u671f\u5728\u4e0a\u4e00\u5e74\u672b\u91d1\u989d"), "report_notdue_last_end", "report_col_16")
    compare_component(report_df, source_df, u(r"\u5f80\u5e74\u672a\u5230\u8d26\u671f\u5728\u4e0a\u4e00\u5e74\u672b\u91d1\u989d"), "report_notdue_old_end", "report_col_17")
    compare_component(report_df, source_df, u(r"\u5c0f\u4e1a\u4e3b\u91d1\u5e01\u56de\u6b3e\u91d1\u989d_\u5f53\u5e74"), "report_coin_recovery", "report_col_10")
    compare_component(report_df, source_df, u(r"\u4e0a\u4e00\u5e74\u91d1\u5e01\u4f59\u989d"), "report_coin_balance", "report_col_15")


def print_scope_check(report_df: pd.DataFrame) -> None:
    project_df = pd.read_excel(ROOT / PROJECT_QUERY_FILE)
    non_assess_df = pd.read_excel(ROOT / NON_ASSESS_FILE)
    level_col = u(r"\u9879\u76ee\u7b49\u7ea7")
    status_col = u(r"\u9879\u76ee\u72b6\u6001")
    code_col = u(r"\u7acb\u9879\u7f16\u7801")

    project_df["code_norm"] = project_df[code_col].map(normalize_code)
    non_assess_df["code_norm"] = non_assess_df[code_col].map(normalize_code)
    non_assess_codes = set(non_assess_df["code_norm"].dropna())

    merged = report_df.merge(
        project_df[["code_norm", level_col, status_col]].drop_duplicates("code_norm"),
        on="code_norm",
        how="left",
    )
    merged["is_non_assess"] = merged["code_norm"].isin(non_assess_codes)
    merged["is_d_exit"] = (
        merged[level_col].astype(str).str.startswith("D")
        & (merged[status_col].astype(str) == u(r"\u5df2\u64a4\u573a"))
    )

    print("[project_scope_check]")
    print(f"project_query_unmatched_rows: {int(merged[level_col].isna().sum())}")
    print(f"non_assess_rows_in_report: {int(merged['is_non_assess'].sum())}")
    print(f"d_exit_rows_in_report: {int(merged['is_d_exit'].sum())}")


def print_assumptions_and_blockers(report_df: pd.DataFrame) -> None:
    numerator_total_col = u(r"\u4e0a\u4e00\u5e74\u56de\u6b3e\u7387\u5206\u5b50")
    denominator_total_col = u(r"\u4e0a\u4e00\u5e74\u56de\u6b3e\u7387\u5206\u6bcd")

    print("[assumptions]")
    for metric, source_name, months in SKIPPED_COMPONENTS:
        print(f"skip: {source_name} {months} ({metric})")
    for metric, source_name, months in ASSUMED_ZERO_COMPONENTS:
        print(f"assume_zero: {source_name} {months} ({metric})")

    print("[remaining_blockers]")
    print("none")

    print("[report_totals_by_status]")
    print(f"numerator_total: {float(report_df[numerator_total_col].sum())}")
    print(f"numerator_business_aging_component: {float(report_df['report_prev_ar_recovery'].sum())}")
    print(f"numerator_skipped_water_detail: {float(report_df['report_water_recv'].sum())}")
    print(f"numerator_verified_components: {float((report_df['report_rel_paid'] + report_df['report_notdue_last_current'] + report_df['report_notdue_old_current'] + report_df['report_coin_recovery']).sum())}")
    print(f"denominator_total: {float(report_df[denominator_total_col].sum())}")
    print(f"denominator_business_aging_component: {float(report_df['report_prev_ar_balance'].sum())}")
    print(f"denominator_skipped_water_detail: {float(report_df['report_water_balance'].sum())}")
    print(f"denominator_coin_balance_component: {float(report_df['report_coin_balance'].sum())}")
    print(f"denominator_verified_components: {float((report_df['report_rel_recv'] + report_df['report_notdue_last_end'] + report_df['report_notdue_old_end']).sum())}")


def print_overall_status() -> None:
    print("[overall_status]")
    print("status: blocked_partial")
    print("reason: water detail remains user-requested skip, and business aging / coin balance now have project-level mismatches to resolve")


def main() -> None:
    indicator_rows = load_indicator_rows()
    report_df = load_report_df()
    source_df = build_component_source_df()

    print_indicator_rows(indicator_rows)
    print_formula_check(report_df)
    print_available_component_checks(report_df, source_df)
    print_scope_check(report_df)
    print_assumptions_and_blockers(report_df)
    print_overall_status()


if __name__ == "__main__":
    main()
