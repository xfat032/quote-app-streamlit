"""Normalize quote item names and display categories."""

from __future__ import annotations

import pandas as pd

from .quote_constants import (
    CATEGORY_MAP,
    CATEGORY_ORDER,
    CONTENT_KEYWORDS,
    EQUIPMENT_KEYWORDS,
    LOGISTICS_KEYWORDS,
    MATERIAL_KEYWORDS,
    PERSONNEL_KEYWORDS,
    PROMOTION_KEYWORDS,
    SCENE_KEYWORDS,
)
from .quote_utils import _matches_keywords, _normalize_item_name, _resolve_item_column


def normalize_category(category: str, item_name: str = "") -> str:
    item_name = _normalize_item_name(item_name)
    category = str(category or "").strip()

    if _matches_keywords(item_name, SCENE_KEYWORDS):
        return "场景搭建类"
    if _matches_keywords(item_name, MATERIAL_KEYWORDS):
        return "活动物料类"
    if _matches_keywords(item_name, CONTENT_KEYWORDS):
        return "活动内容类"
    if _matches_keywords(item_name, EQUIPMENT_KEYWORDS):
        return "设备租赁类"
    if _matches_keywords(item_name, PERSONNEL_KEYWORDS):
        return "人员服务类"
    if _matches_keywords(item_name, LOGISTICS_KEYWORDS):
        return "后勤保障类"
    if _matches_keywords(item_name, PROMOTION_KEYWORDS):
        return "宣传推广类"

    mapped_category = CATEGORY_MAP.get(category)
    if mapped_category:
        return mapped_category
    return "后勤保障类"


def normalize_quote_categories(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty or "项目分类" not in df.columns:
        return df.copy()

    normalized_df = df.copy()
    item_column = _resolve_item_column(normalized_df)
    if item_column is None:
        normalized_df["项目分类"] = normalized_df["项目分类"].map(lambda value: normalize_category(str(value), ""))
        return normalized_df

    normalized_df["项目分类"] = normalized_df.apply(
        lambda row: normalize_category(row.get("项目分类", ""), row.get(item_column, "")),
        axis=1,
    )
    return normalized_df


def sort_quote_items(df: pd.DataFrame) -> pd.DataFrame:
    """Normalize final categories and keep same-category items contiguous."""
    if df.empty or "项目分类" not in df.columns:
        return df.copy()

    sorted_df = normalize_quote_categories(df)
    category_order = {category: index for index, category in enumerate(CATEGORY_ORDER)}
    item_column = _resolve_item_column(sorted_df)

    sorted_df["_category_rank"] = sorted_df["项目分类"].map(lambda value: category_order.get(str(value), len(CATEGORY_ORDER)))
    sorted_df["_item_sort"] = sorted_df[item_column].astype(str) if item_column else ""
    sorted_df = sorted_df.sort_values(
        by=["_category_rank", "项目分类", "_item_sort"],
        kind="stable",
    ).drop(columns=["_category_rank", "_item_sort"])
    return sorted_df
