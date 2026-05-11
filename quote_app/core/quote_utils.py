"""Small shared helpers used by quote building modules."""

from __future__ import annotations

import re
from typing import Any

import pandas as pd


def _resolve_item_column(df: pd.DataFrame) -> str | None:
    for column in ("标准项目", "项目", "建议补充项"):
        if column in df.columns:
            return column
    return None


def _matches_keywords(item_name: str, keywords: list[str]) -> bool:
    return any(keyword in item_name for keyword in keywords if keyword)


def _safe_first_match_position(value: Any) -> int | None:
    if isinstance(value, list):
        for item in value:
            try:
                return int(item)
            except (TypeError, ValueError):
                continue
        return None
    try:
        return int(value)
    except (TypeError, ValueError):
        return None


def _normalize_item_name(item_name: str) -> str:
    item_name = str(item_name or "").strip()
    if item_name in {"市集摊位", "帐篷摊位"}:
        return "帐篷摊位"
    if item_name in {
        "开场表演",
        "节目表演",
        "乐队演绎",
        "舞蹈演绎",
        "脱口秀演绎",
        "巡游表演",
        "非遗演绎",
        "服饰走秀",
        "音乐会",
        "琼剧表演",
    }:
        return "节目演出"
    if item_name in {"工作牌", "工作卡", "胸卡", "人员证件"}:
        return "工作证"
    if item_name in {
        "主视觉设计",
        "KV设计",
        "活动海报设计",
        "倒计时海报设计",
        "延展设计",
        "物料画面设计",
        "导视系统设计",
        "舞台视觉设计",
        "摊位视觉设计",
        "IP形象设计",
        "品牌视觉设计",
        "宣传物料设计",
        "活动视觉设计",
        "美陈画面设计",
        "展板画面设计",
        "背景板画面设计",
    }:
        return "设计服务"
    return item_name


def _append_note(existing: str, note: str) -> str:
    parts = [part for part in str(existing or "").split("；") if part]
    for note_part in [part for part in str(note or "").split("；") if part]:
        if note_part not in parts:
            parts.append(note_part)
    return "；".join(parts)


CONFIRMATION_CONTENT_RE = re.compile(
    r"(需确认|待确认|人工确认|确认|核价|核定|报价口径|计价|测算)"
)


def sanitize_quote_content_text(value: Any, fallback: Any = "") -> str:
    """Keep only item content in the content/spec column; move confirmation needs to remarks."""
    text = str(value or "").strip()
    fallback_text = str(fallback or "").strip()
    if not text:
        return fallback_text

    parts = [
        part.strip()
        for part in re.split(r"[，；。,.]", re.sub(r"\s+", " ", text))
        if part.strip()
    ]
    kept_parts = [part for part in parts if not CONFIRMATION_CONTENT_RE.search(part)]
    cleaned = "，".join(kept_parts).strip("，；。,. ")
    if cleaned:
        return cleaned
    return fallback_text or re.sub(CONFIRMATION_CONTENT_RE, "", text).strip("，；。,. ")


def sanitize_quote_content_df(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty or "内容/尺寸/工艺" not in df.columns:
        return df.copy()

    sanitized = df.copy()
    item_column = _resolve_item_column(sanitized)
    for index, row in sanitized.iterrows():
        fallback = row.get(item_column, "") if item_column else ""
        sanitized.at[index, "内容/尺寸/工艺"] = sanitize_quote_content_text(row.get("内容/尺寸/工艺", ""), fallback)
    return sanitized


def _merge_text_values(left: Any, right: Any, separator: str = "、") -> str:
    values: list[str] = []
    for raw_value in (left, right):
        raw_parts = raw_value if isinstance(raw_value, list) else str(raw_value or "").split(separator)
        for part in raw_parts:
            value = str(part).strip()
            if value and value not in values:
                values.append(value)
    return separator.join(values)


def _merge_list_values(left: Any, right: Any) -> list[Any]:
    values: list[Any] = []
    for raw_value in (left, right):
        raw_items = raw_value if isinstance(raw_value, list) else [raw_value]
        for item in raw_items:
            if item in ("", None):
                continue
            if item not in values:
                values.append(item)
    return values


def _is_empty(value: Any) -> bool:
    return pd.isna(value) or value == ""


def _to_number(value: Any) -> float | None:
    if _is_empty(value):
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None
