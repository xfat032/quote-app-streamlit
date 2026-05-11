"""Build editable quote rows and enrich them with the empty price DB template."""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any

import pandas as pd
from openpyxl import Workbook

from .quote_constants import DESIGN_TRIGGER_KEYWORDS, PRICE_DB_COLUMNS, PUBLIC_QUOTE_SECTIONS, QUOTE_COLUMNS
from .quote_categories import normalize_category, normalize_quote_categories, sort_quote_items
from .quote_utils import (
    _append_note,
    _is_empty,
    _merge_list_values,
    _merge_text_values,
    _normalize_item_name,
    _resolve_item_column,
    _safe_first_match_position,
    sanitize_quote_content_df,
    sanitize_quote_content_text,
    _to_number,
)
from .section_recognition import (
    assign_quote_section,
    classify_section_candidate,
    classify_section_level,
    diagnose_activity_content_ranges,
    diagnose_section_candidates,
    extract_activity_content_ranges,
    extract_activity_sections,
    find_section_for_match,
    infer_activity_sections_from_quote_rows,
    infer_section_from_item,
    is_public_merge_item,
    normalize_section_name,
    prepare_confirmed_activity_sections,
    split_compound_section_name,
    validate_section_source,
)


def make_quote_item_key(item: dict[str, Any]) -> str:
    standard_item = _normalize_item_name(item.get("标准项目", item.get("项目", "")))
    if is_public_merge_item(standard_item) and str(item.get("evidence_type", "")) != "module_completion":
        return f"PUBLIC::{standard_item}"

    quote_section = str(item.get("quote_section") or item.get("项目分类") or "").strip()
    if not quote_section:
        quote_section = assign_quote_section(
            item,
            _safe_first_match_position(item.get("匹配位置")),
            str(item.get("source_context_text", "")),
            [],
        )
    evidence_type = str(item.get("evidence_type") or "").strip()
    trigger_section = str(item.get("trigger_section") or item.get("matched_section_name") or "").strip()
    evidence_text = str(item.get("evidence_text") or item.get("trigger_module") or "").strip()
    trigger = trigger_section or evidence_text
    return f"{quote_section}::{standard_item}::{evidence_type}::{trigger}"


QUANTITY_SOURCE_RE = re.compile(r"根据“([^”]+)”识别[^；]*数量")


def _quantity_source_texts(remark: Any) -> set[str]:
    return set(QUANTITY_SOURCE_RE.findall(str(remark or "")))


def _has_same_quantity_source(left_remark: Any, right_remark: Any) -> bool:
    return bool(_quantity_source_texts(left_remark) & _quantity_source_texts(right_remark))


def _merge_quantities(base: dict[str, Any], incoming: dict[str, Any]) -> None:
    base_quantity = _to_number(base.get("数量"))
    incoming_quantity = _to_number(incoming.get("数量"))
    same_unit = str(base.get("单位", "")) == str(incoming.get("单位", ""))
    base_remark = str(base.get("备注", ""))
    incoming_remark = str(incoming.get("备注", ""))
    quantity_is_clear = "未明确数量" not in base_remark and "未明确数量" not in incoming_remark

    if (
        base_quantity is not None
        and incoming_quantity is not None
        and same_unit
        and base_quantity == incoming_quantity
        and _has_same_quantity_source(base_remark, incoming_remark)
    ):
        base["数量"] = base_quantity
        base["备注"] = _append_note(base.get("备注", ""), "同一数量线索多别名命中，未重复相加")
        return

    if base_quantity is not None and incoming_quantity is not None and same_unit and quantity_is_clear:
        base["数量"] = base_quantity + incoming_quantity
        return

    base["数量"] = base.get("数量") or 1
    base["备注"] = _append_note(base.get("备注", ""), "多处命中，数量需确认")


def _merge_quote_rows(rows: list[dict[str, Any]], activity_sections: list[dict[str, Any]]) -> list[dict[str, Any]]:
    merged: dict[str, dict[str, Any]] = {}
    source_sections: dict[str, list[str]] = {}

    for row in rows:
        key = make_quote_item_key(row)
        source_section = str(row.get("trigger_section") or row.get("matched_section_name") or row.get("quote_section") or "").strip()
        if key not in merged:
            merged[key] = dict(row)
            source_sections[key] = [source_section] if source_section else []
            continue

        target = merged[key]
        if source_section and source_section not in source_sections[key]:
            source_sections[key].append(source_section)

        target["原始命中词"] = _merge_text_values(target.get("原始命中词"), row.get("原始命中词"))
        target["evidence_text"] = _merge_text_values(target.get("evidence_text"), row.get("evidence_text"))
        target["trigger_module"] = _merge_text_values(target.get("trigger_module"), row.get("trigger_module"))
        target["匹配位置"] = _merge_list_values(target.get("匹配位置"), row.get("匹配位置"))
        target["命中模块"] = _merge_list_values(target.get("命中模块"), row.get("命中模块"))
        _merge_quantities(target, row)
        target["备注"] = _append_note(target.get("备注", ""), str(row.get("备注", "")))

    for key, row in merged.items():
        sections = [section for section in source_sections.get(key, []) if section]
        if key.startswith("PUBLIC::") and len(sections) > 1:
            row["备注"] = _append_note(row.get("备注", ""), f"由多个活动板块共同命中：{'、'.join(sections)}")

    return sort_by_quote_section(pd.DataFrame(merged.values()), activity_sections).to_dict("records")


def reassign_quote_sections(
    df: pd.DataFrame,
    source_text: str,
    activity_sections: list[dict[str, Any]] | None = None,
) -> pd.DataFrame:
    if df.empty:
        return df.copy()

    reassigned_df = df.copy()
    confirmed_sections = prepare_confirmed_activity_sections(activity_sections or [])
    section_order = _build_quote_section_order(confirmed_sections)

    for index, row in reassigned_df.iterrows():
        match_start = _safe_first_match_position(row.get("匹配位置"))
        matched_section = find_section_for_match(match_start, confirmed_sections)
        quote_section = assign_quote_section(row.to_dict(), match_start, str(row.get("source_context_text", "") or source_text), confirmed_sections)
        reassigned_df.at[index, "quote_section"] = quote_section
        reassigned_df.at[index, "quote_section_order"] = section_order.get(quote_section, len(section_order) + 1)
        reassigned_df.at[index, "项目分类"] = quote_section
        reassigned_df.at[index, "matched_section_name"] = matched_section["name"] if matched_section else ""
        reassigned_df.at[index, "trigger_section"] = matched_section["name"] if matched_section else quote_section
        reassigned_df.at[index, "source_context_start"] = match_start if match_start is not None else ""
        if "source_context_text" in reassigned_df.columns and not str(row.get("source_context_text", "")).strip() and match_start is not None:
            reassigned_df.at[index, "source_context_text"] = source_text[max(0, match_start - 40) : min(len(source_text), match_start + 80)]

        remark = str(reassigned_df.at[index, "备注"] if "备注" in reassigned_df.columns else "").strip()
        remark = remark.replace("；未能识别所属活动板块，需人工确认", "").replace("未能识别所属活动板块，需人工确认", "").strip("；")
        if quote_section == "未归属板块":
            addition = "未能识别所属活动板块，需人工确认"
            remark = f"{remark}；{addition}" if remark else addition
        if "备注" in reassigned_df.columns:
            reassigned_df.at[index, "备注"] = remark

    return sort_by_quote_section(reassigned_df, confirmed_sections)


def _build_quote_section_order(
    activity_sections: list[dict[str, Any]],
    observed_sections: list[str] | None = None,
) -> dict[str, int]:
    order: dict[str, int] = {"活动宣传": 0}
    index = 1
    for section in activity_sections:
        name = section["name"]
        if name in PUBLIC_QUOTE_SECTIONS or name == "活动宣传":
            continue
        if name not in order:
            order[name] = index
            index += 1

    for name in observed_sections or []:
        if not name or name in order or name in PUBLIC_QUOTE_SECTIONS or name == "活动宣传":
            continue
        order[name] = index
        index += 1

    order["美陈搭建类"] = index
    order["其他搭建类"] = index + 1
    order["未归属板块"] = index + 2
    order["人员类及其他"] = index + 3
    return order


def sort_by_quote_section(df: pd.DataFrame, activity_sections: list[dict[str, Any]] | None = None) -> pd.DataFrame:
    if df.empty:
        return df.copy()

    sorted_df = df.copy()
    confirmed_sections = prepare_confirmed_activity_sections(activity_sections or [])
    if "quote_section" not in sorted_df.columns or sorted_df["quote_section"].astype(str).eq("").all():
        if "项目分类" in sorted_df.columns and any(str(value) in PUBLIC_QUOTE_SECTIONS or str(value) in {section["name"] for section in confirmed_sections} or str(value) == "未归属板块" for value in sorted_df["项目分类"].fillna("")):
            sorted_df["quote_section"] = sorted_df["项目分类"].astype(str)
        else:
            item_column = _resolve_item_column(sorted_df)
            if item_column:
                sorted_df["quote_section"] = sorted_df.apply(
                    lambda row: assign_quote_section(
                        row.to_dict(),
                        _safe_first_match_position(row.get("匹配位置")),
                        str(row.get("source_context_text", "")),
                        confirmed_sections,
                    ),
                    axis=1,
                )
            else:
                sorted_df["quote_section"] = "未归属板块"

    observed_sections = [str(value) for value in sorted_df["quote_section"].fillna("").tolist() if str(value).strip()]
    section_order = _build_quote_section_order(confirmed_sections, observed_sections)

    sorted_df["quote_section_order"] = sorted_df["quote_section"].map(
        lambda value: section_order.get(str(value), len(section_order) + 1)
    )

    item_column = _resolve_item_column(sorted_df)
    sorted_df["_item_sort"] = sorted_df[item_column].map(_normalize_item_name).astype(str) if item_column else ""
    sorted_df = sorted_df.sort_values(
        by=["quote_section_order", "quote_section", "_item_sort"],
        kind="stable",
    ).drop(columns=["_item_sort"])
    if "项目分类" in sorted_df.columns:
        sorted_df["项目分类"] = sorted_df["quote_section"]
    return sorted_df


def _maybe_add_design_service(rows: list[dict[str, Any]], text: str) -> list[dict[str, Any]]:
    if any(str(row.get("标准项目", "")) == "设计服务" for row in rows):
        return rows

    hit_keyword = next((keyword for keyword in DESIGN_TRIGGER_KEYWORDS if keyword in text), "")
    if not hit_keyword:
        return rows

    design_row = {
        "是否保留": True,
        "项目分类": "宣传推广类",
        "标准项目": "设计服务",
        "原始命中词": hit_keyword,
        "内容/尺寸/工艺": "方案提及设计服务，具体设计范围需确认",
        "数量": 1,
        "单位": "项",
        "单价": 0,
        "合计": 0,
        "报价类型": "模糊报价",
        "来源状态": "需确认",
        "备注": "根据方案中的设计相关表述汇总；需确认设计范围",
        "匹配位置": [text.find(hit_keyword)],
        "命中模块": [],
        "evidence_type": "explicit_text",
        "evidence_text": hit_keyword,
        "trigger_module": "",
    }
    return [*rows, design_row]



def ensure_price_db(path: str | Path) -> Path:
    """Create an empty price DB workbook when missing."""
    path = Path(path)
    if path.exists():
        return path

    path.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "价格库"
    sheet.append(PRICE_DB_COLUMNS)
    workbook.save(path)
    return path


def load_price_db(path: str | Path) -> pd.DataFrame:
    path = ensure_price_db(path)
    try:
        df = pd.read_excel(path)
    except Exception:
        df = pd.DataFrame(columns=PRICE_DB_COLUMNS)

    for column in PRICE_DB_COLUMNS:
        if column not in df.columns:
            df[column] = ""

    return df[PRICE_DB_COLUMNS]


def recalculate_totals(df: pd.DataFrame) -> pd.DataFrame:
    """Recalculate totals without forcing users to fill unit prices."""
    df = df.copy()

    for index, row in df.iterrows():
        quantity = _to_number(row.get("数量"))
        unit_price = _to_number(row.get("单价"))
        if quantity is None or unit_price is None:
            df.at[index, "合计"] = 0
        else:
            df.at[index, "合计"] = quantity * unit_price

    return df


def dedupe_final_quote_items(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df.copy()

    df = sanitize_quote_content_df(df)
    deduped_rows: list[dict[str, Any]] = []
    key_to_index: dict[str, int] = {}

    item_column = _resolve_item_column(df)
    if item_column is None:
        return df.copy()

    for _, row in df.iterrows():
        row_dict = row.to_dict()
        section = str(row_dict.get("quote_section") or row_dict.get("项目分类") or "").strip()
        item_name = str(row_dict.get(item_column, "")).strip()
        key = f"{section}::{item_name}"

        if key not in key_to_index:
            key_to_index[key] = len(deduped_rows)
            deduped_rows.append(row_dict)
            continue

        target = deduped_rows[key_to_index[key]]
        source_values = _merge_text_values(
            _merge_text_values(target.get("原始命中词"), row_dict.get("原始命中词")),
            _merge_text_values(target.get("evidence_text"), row_dict.get("evidence_text")),
        )
        if not source_values:
            source_values = item_name

        target_quantity = _to_number(target.get("数量"))
        row_quantity = _to_number(row_dict.get("数量"))
        same_unit = str(target.get("单位", "")) == str(row_dict.get("单位", ""))
        same_quantity_source = target_quantity == row_quantity and _has_same_quantity_source(target.get("备注", ""), row_dict.get("备注", ""))
        need_text = "；".join(
            str(value)
            for value in [
                target.get("确认状态", ""),
                target.get("需要确认什么", ""),
                target.get("备注", ""),
                row_dict.get("确认状态", ""),
                row_dict.get("需要确认什么", ""),
                row_dict.get("备注", ""),
            ]
            if value
        )

        if (
            target_quantity is not None
            and row_quantity is not None
            and same_unit
            and not same_quantity_source
            and not (target_quantity == 1 and row_quantity == 1 and ("需确认数量" in need_text or "缺数量" in need_text or "数量需确认" in need_text or "未明确数量" in need_text))
        ):
            target["数量"] = target_quantity + row_quantity
        elif same_quantity_source:
            target["数量"] = target_quantity
            target["备注"] = _append_note(target.get("备注", ""), "同一数量线索多别名命中，未重复相加")

        target["备注"] = _append_note(target.get("备注", ""), str(row_dict.get("备注", "")))
        target["备注"] = _append_note(target.get("备注", ""), f"多处命中：{source_values}；数量需确认")

        for column in ("确认状态", "需要确认什么"):
            if column in row_dict:
                target[column] = _merge_text_values(target.get(column, ""), row_dict.get(column, ""), " / ")

        target["原始命中词"] = _merge_text_values(target.get("原始命中词"), row_dict.get("原始命中词"))
        target["evidence_text"] = _merge_text_values(target.get("evidence_text"), row_dict.get("evidence_text"))
        target["trigger_module"] = _merge_text_values(target.get("trigger_module"), row_dict.get("trigger_module"))
        target["匹配位置"] = _merge_list_values(target.get("匹配位置"), row_dict.get("匹配位置"))
        target["命中模块"] = _merge_list_values(target.get("命中模块"), row_dict.get("命中模块"))

    deduped_df = pd.DataFrame(deduped_rows)
    return recalculate_totals(deduped_df) if "合计" in deduped_df.columns else deduped_df


def finalize_quote_rows(
    rows: list[dict[str, Any]],
    source_text: str,
    activity_sections: list[dict[str, Any]] | None = None,
) -> list[dict[str, Any]]:
    all_activity_sections = activity_sections or extract_activity_sections(source_text)
    activity_sections = prepare_confirmed_activity_sections(all_activity_sections)
    section_order = _build_quote_section_order(activity_sections)
    finalized_rows = _maybe_add_design_service(rows, source_text)
    kept_rows: list[dict[str, Any]] = []

    for row in finalized_rows:
        normalized_item_name = _normalize_item_name(row.get("标准项目", row.get("项目", "")))
        match_start = _safe_first_match_position(row.get("匹配位置"))
        matched_section = find_section_for_match(match_start, activity_sections)
        matched_section_name = matched_section["name"] if matched_section else ""
        evidence_type = str(row.get("evidence_type", "")).strip()

        if evidence_type == "module_completion":
            if not matched_section_name:
                continue
            row["trigger_module"] = matched_section_name
            row["evidence_text"] = f"由【{matched_section_name}】模块补全"
        elif evidence_type == "explicit_text":
            if not str(row.get("evidence_text", "")).strip():
                continue
        elif evidence_type and evidence_type != "user_selected_suggestion":
            continue

        row["标准项目"] = normalized_item_name
        row["内容/尺寸/工艺"] = sanitize_quote_content_text(row.get("内容/尺寸/工艺", ""), normalized_item_name)
        row["internal_category"] = normalize_category(row.get("项目分类", ""), normalized_item_name)
        row["quote_section"] = assign_quote_section(row, match_start, source_text, activity_sections)
        row["quote_section_order"] = section_order.get(row["quote_section"], len(section_order) + 1)
        row["项目分类"] = row["quote_section"]
        row["matched_section_name"] = matched_section_name
        row["trigger_section"] = matched_section_name or row["quote_section"]
        row["source_context_start"] = match_start if match_start is not None else ""
        row["source_context_text"] = (
            source_text[max(0, match_start - 40) : min(len(source_text), match_start + 80)]
            if match_start is not None
            else ""
        )
        if row["quote_section"] == "未归属板块":
            remark = str(row.get("备注", "")).strip()
            addition = "未能识别所属活动板块，需人工确认"
            row["备注"] = f"{remark}；{addition}" if remark else addition
        kept_rows.append(row)

    return _merge_quote_rows(kept_rows, activity_sections)


def build_quote_rows(
    extracted_rows: list[dict[str, Any]],
    price_db_path: str | Path,
    source_text: str = "",
    activity_sections: list[dict[str, Any]] | None = None,
) -> list[dict[str, Any]]:
    """Merge extracted rows with price DB defaults."""
    price_db = load_price_db(price_db_path)
    built_rows: list[dict[str, Any]] = []

    for extracted in extracted_rows:
        row = {column: extracted.get(column, "") for column in QUOTE_COLUMNS}
        row["是否保留"] = bool(row.get("是否保留", True))
        row["单价"] = 0 if _is_empty(row.get("单价")) else row["单价"]
        row["匹配位置"] = extracted.get("匹配位置", [])
        row["命中模块"] = extracted.get("命中模块", [])

        matches = price_db[
            (price_db["项目分类"].fillna("") == row["项目分类"])
            & (price_db["标准项目"].fillna("") == row["标准项目"])
        ]

        if not matches.empty:
            price_row = matches.iloc[0]
            if not _is_empty(price_row.get("默认规格")):
                row["内容/尺寸/工艺"] = price_row["默认规格"]
            if not _is_empty(price_row.get("单位")):
                row["单位"] = price_row["单位"]
            if not _is_empty(price_row.get("默认单价")):
                row["单价"] = price_row["默认单价"]
            if not _is_empty(price_row.get("报价类型")):
                row["报价类型"] = price_row["报价类型"]
            if not _is_empty(price_row.get("备注")):
                row["备注"] = f"{row['备注']}；{price_row['备注']}" if row["备注"] else price_row["备注"]

        row["内容/尺寸/工艺"] = sanitize_quote_content_text(row.get("内容/尺寸/工艺", ""), row.get("标准项目", ""))
        built_rows.append(row)

    df = recalculate_totals(pd.DataFrame(built_rows, columns=QUOTE_COLUMNS))
    return finalize_quote_rows(df.to_dict("records"), source_text, activity_sections=activity_sections)
