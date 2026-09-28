"""원본 캡스 로그의 행과 사업장 정보를 정제한다."""

from __future__ import annotations

import numpy as np
import pandas as pd

from config.data_config import (
    COMPANY_COL,
    DATA_CLEANER_FILTER_ITEMS,
    TAEIL_CABLE,
    TAEIL_MATERIAL,
)


def _normalize_text(value: object) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def _to_normalized_set(value: object) -> set[str]:
    if value is None:
        return set()

    if isinstance(value, str):
        normalized = _normalize_text(value)
        return {normalized} if normalized else {""}

    if isinstance(value, (list, tuple, set, frozenset)):
        return {_normalize_text(item) for item in value}

    normalized = _normalize_text(value)
    return {normalized} if normalized else {""}


def _build_rule_sets(rule: object) -> tuple[set[str], set[str]]:
    """한 컬럼의 필터 규칙을 include/exclude 집합으로 정규화한다."""
    if isinstance(rule, dict):
        include_values = _to_normalized_set(rule.get("include"))
        exclude_values = _to_normalized_set(rule.get("exclude"))
        return include_values, exclude_values

    return set(), _to_normalized_set(rule)


def filter_rows(df: pd.DataFrame) -> pd.DataFrame:
    """설정된 필터 규칙에 따라 계산에 불필요한 행을 제거한다."""
    if df is None or df.empty:
        return df.copy()

    working = df.copy()
    keep_mask = pd.Series(True, index=working.index)

    for column, rule in DATA_CLEANER_FILTER_ITEMS.items():
        if column not in working.columns:
            continue

        include_values, exclude_values = _build_rule_sets(rule)
        normalized_column = working[column].map(_normalize_text)

        if include_values:
            keep_mask &= normalized_column.isin(include_values)

        if exclude_values:
            keep_mask &= ~normalized_column.isin(exclude_values)

    return working.loc[keep_mask].reset_index(drop=True)


def add_company_column(
    df,
    taeil_cable=TAEIL_CABLE,
    taeil_material=TAEIL_MATERIAL,
    user_col="이름",
):
    """입력 DataFrame에 사업장 컬럼을 추가한다."""
    conditions = [
        df[user_col].isin(taeil_cable),
        df[user_col].isin(taeil_material),
    ]
    choices = ["태일전선", "태일소재"]
    df[COMPANY_COL] = np.select(conditions, choices, default="미등록")
    return df
