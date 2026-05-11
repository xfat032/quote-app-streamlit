"""State management and dataframe services for the Streamlit app."""

from __future__ import annotations

import hashlib
import json

import pandas as pd
import streamlit as st

from app_config import (
    FEEDBACK_COLUMNS,
    FINAL_QUOTE_COLUMNS,
    IGNORED_TERMS_PATH,
    KEY_SUGGESTION_ITEMS,
    NO_QUOTE_TERMS,
    PRICE_DB_PATH,
    QUOTE_CANDIDATE_ROOTS,
    RULES_PATH,
    SECTION_CONFIRM_COLUMNS,
    SUGGESTION_EDITOR_COLUMNS,
)
from core.activity_classifier import build_suggested_items, classify_activity_types
from core.extractor import calculate_coverage, extract_quote_items, extract_unrecognized_candidates
from core.normalizer import load_rules
from core.quote_builder import (
    QUOTE_COLUMNS,
    assign_quote_section,
    build_quote_rows,
    diagnose_activity_content_ranges,
    diagnose_section_candidates,
    dedupe_final_quote_items,
    extract_activity_sections,
    make_quote_item_key,
    prepare_confirmed_activity_sections,
    recalculate_totals,
    sort_by_quote_section,
)
from core.quote_utils import sanitize_quote_content_text
from core.rule_feedback import ensure_ignored_terms, load_ignored_terms

rules: dict = {}


def configure_app_services(current_rules: dict) -> None:
    global rules
    rules = current_rules


def initialize_session_state() -> None:
    defaults = {
        "plan_text_input": "",
        "plan_text_value": "",
        "quote_df": pd.DataFrame(columns=QUOTE_COLUMNS),
        "raw_quote_df": pd.DataFrame(columns=QUOTE_COLUMNS),
        "working_quote_df": pd.DataFrame(columns=QUOTE_COLUMNS),
        "review_quote_df": pd.DataFrame(columns=QUOTE_COLUMNS),
        "recognized_df": pd.DataFrame(columns=QUOTE_COLUMNS),
        "candidate_df": pd.DataFrame(columns=["候选词", "所在上下文", "建议动作", "是否加入规则库"]),
        "candidate_display_df": pd.DataFrame(columns=FEEDBACK_COLUMNS),
        "feedback_df": pd.DataFrame(columns=FEEDBACK_COLUMNS),
        "activity_type_df": pd.DataFrame(columns=["活动类型", "命中关键词", "重点规则"]),
        "activity_sections": [],
        "activity_sections_all": [],
        "activity_content_diagnostics": {},
        "activity_section_df": pd.DataFrame(columns=SECTION_CONFIRM_COLUMNS),
        "suggested_df": pd.DataFrame(columns=["是否加入报价单", "建议补充项", "项目分类", "报价类型", "来源状态", "建议原因", "备注"]),
        "suggestion_editor_df": pd.DataFrame(columns=SUGGESTION_EDITOR_COLUMNS),
        "coverage": None,
        "current_file_signature": "",
        "last_text_hash": "",
        "excel_bytes": None,
        "upload_notice": "",
        "quote_editor_version": 0,
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value

def append_default_ignored_terms() -> None:
    existing = load_ignored_terms(IGNORED_TERMS_PATH)
    updated = list(dict.fromkeys([*existing, *NO_QUOTE_TERMS]))
    if updated != existing:
        IGNORED_TERMS_PATH.write_text(json.dumps(updated, ensure_ascii=False, indent=2), encoding="utf-8")

def summarize_text(value: str, limit: int = 40) -> str:
    value = " ".join(str(value).split())
    if len(value) <= limit:
        return value
    if limit <= 3:
        return value[:limit]
    return f"{value[: limit - 3]}..."


def get_file_signature(uploaded_file) -> str:
    file_bytes = uploaded_file.getvalue()
    file_hash = hashlib.md5(file_bytes).hexdigest()
    return f"{uploaded_file.name}_{uploaded_file.size}_{file_hash}"


def get_text_hash(text: str) -> str:
    return hashlib.md5(str(text).encode("utf-8")).hexdigest()


def clear_result_state() -> None:
    reset_editor_widget_state()
    st.session_state["raw_quote_df"] = pd.DataFrame(columns=QUOTE_COLUMNS)
    st.session_state["working_quote_df"] = pd.DataFrame(columns=QUOTE_COLUMNS)
    st.session_state["review_quote_df"] = pd.DataFrame(columns=QUOTE_COLUMNS)
    st.session_state["recognized_df"] = pd.DataFrame(columns=QUOTE_COLUMNS)
    st.session_state["quote_df"] = pd.DataFrame(columns=QUOTE_COLUMNS)
    st.session_state["candidate_df"] = pd.DataFrame(columns=["候选词", "所在上下文", "建议动作", "是否加入规则库"])
    st.session_state["candidate_display_df"] = pd.DataFrame(columns=FEEDBACK_COLUMNS)
    st.session_state["feedback_df"] = pd.DataFrame(columns=FEEDBACK_COLUMNS)
    st.session_state["activity_type_df"] = pd.DataFrame(columns=["活动类型", "命中关键词", "重点规则"])
    st.session_state["activity_sections"] = []
    st.session_state["activity_sections_all"] = []
    st.session_state["activity_content_diagnostics"] = {}
    st.session_state["section_candidate_diagnostics"] = pd.DataFrame(columns=["候选标题", "分类", "过滤原因", "位置", "原始文本"])
    st.session_state["activity_section_df"] = pd.DataFrame(columns=SECTION_CONFIRM_COLUMNS)
    st.session_state["suggested_df"] = pd.DataFrame(columns=["是否加入报价单", "建议补充项", "项目分类", "报价类型", "来源状态", "建议原因", "备注"])
    st.session_state["suggestion_editor_df"] = pd.DataFrame(columns=SUGGESTION_EDITOR_COLUMNS)
    st.session_state["coverage"] = None
    st.session_state["excel_bytes"] = None
    st.session_state["raw_text"] = ""
    st.session_state["recognized_items"] = []
    st.session_state["quote_items"] = []
    st.session_state["final_quote_df"] = pd.DataFrame(columns=FINAL_QUOTE_COLUMNS)
    st.session_state["confirmed_quote_df"] = pd.DataFrame(columns=FINAL_QUOTE_COLUMNS)
    st.session_state["confirmed_sections"] = []
    st.session_state["section_editor_df"] = pd.DataFrame(columns=SECTION_CONFIRM_COLUMNS)
    st.session_state["unrecognized_candidates"] = []
    st.session_state["suggested_items"] = []
    st.session_state["generated_excel"] = None
    st.session_state["last_result"] = None
    st.session_state["last_text_hash"] = ""
    st.session_state["quote_editor_version"] = st.session_state.get("quote_editor_version", 0) + 1


def reset_editor_widget_state() -> None:
    for key in [
        "review_quote_editor",
        "final_quote_editor",
        "activity_section_editor",
        "suggestion_quote_editor",
        "candidate_feedback_editor",
    ]:
        if key in st.session_state:
            del st.session_state[key]


def sync_working_quote_state(working_df: pd.DataFrame, activity_sections: list[dict] | None = None) -> None:
    normalized_df = sort_by_quote_section(working_df.copy(), activity_sections or st.session_state.get("activity_sections", []))
    normalized_df = dedupe_final_quote_items(normalized_df)
    normalized_df = sort_by_quote_section(normalized_df, activity_sections or st.session_state.get("activity_sections", []))
    st.session_state["working_quote_df"] = normalized_df.copy()
    st.session_state["review_quote_df"] = make_judgement_df(normalized_df)
    st.session_state["quote_df"] = normalized_df.copy()
    st.session_state["recognized_df"] = normalized_df.copy()
    final_df = make_final_quote_df(normalized_df)
    st.session_state["final_quote_df"] = final_df.copy()
    st.session_state["confirmed_quote_df"] = final_df.copy()


def infer_candidate_item(candidate: str) -> str:
    mapping = [
        (["舞台"], "舞台搭建"),
        (["灯光", "音响"], "灯光音响套装"),
        (["视频"], "视频快剪"),
        (["摄影", "拍摄"], "摄影摄像服务"),
        (["直播"], "图片直播"),
        (["海报"], "倒计时海报"),
        (["推文"], "公众号软文"),
        (["媒体"], "主流媒体宣传"),
        (["达人", "KOL"], "达人KOL推广"),
        (["市集", "摊位", "帐篷"], "帐篷摊位"),
        (["展板"], "展板"),
        (["展览"], "艺术展陈"),
        (["装置"], "美陈装置"),
        (["打卡"], "打卡装置"),
        (["导视", "指引"], "导视指引"),
        (["互动", "游戏"], "趣味互动游戏"),
        (["道具"], "互动道具"),
        (["礼品", "奖品"], "文创礼品"),
        (["人员"], "执行人员"),
        (["主持"], "主持人"),
        (["安保"], "安保人员"),
        (["医疗"], "医疗保障"),
        (["志愿者"], "志愿者"),
        (["交通", "停车"], "交通停车引导"),
        (["电力", "发电"], "电力保障"),
        (["围挡", "铁马"], "铁马围挡"),
        (["餐费"], "餐费"),
        (["饮用水"], "饮用水"),
        (["运输"], "运输费"),
        (["手作", "体验"], "手作体验材料"),
        (["花灯", "灯笼", "河灯"], "灯光艺术装置"),
        (["巡游", "NPC"], "NPC互动服务"),
        (["服饰", "妆造"], "演艺服化道"),
        (["长桌宴"], "长桌宴"),
        (["茶席"], "茶席茶寮"),
        (["阅读"], "阅读区布置"),
        (["图书"], "图书配置"),
        (["二维码"], "音频二维码点位"),
    ]
    for roots, standard_item in mapping:
        if any(root in candidate for root in roots):
            return standard_item
    return "人工确认"


def normalize_feedback_df(feedback_df: pd.DataFrame) -> pd.DataFrame:
    if feedback_df.empty:
        return pd.DataFrame(columns=FEEDBACK_COLUMNS)

    normalized_df = feedback_df.copy()
    if "是否处理" not in normalized_df.columns:
        normalized_df["是否处理"] = normalized_df.get("是否加入规则库", False)
    if "选择标准项目" not in normalized_df.columns:
        normalized_df["选择标准项目"] = ""

    for column in FEEDBACK_COLUMNS:
        if column not in normalized_df.columns:
            normalized_df[column] = ""

    normalized_df["是否处理"] = normalized_df["是否处理"].fillna(False).astype(bool)
    normalized_df["候选词"] = normalized_df["候选词"].astype(str)
    normalized_df["可能归类"] = normalized_df["可能归类"].astype(str)
    normalized_df["上下文摘要"] = normalized_df["上下文摘要"].map(lambda value: summarize_text(value, 40))
    normalized_df["处理方式"] = normalized_df["处理方式"].replace("", "暂不处理").fillna("暂不处理")
    valid_actions = {"暂不处理", "加入已有标准项目别名", "新建标准项目", "标记为无需报价"}
    normalized_df.loc[~normalized_df["处理方式"].isin(valid_actions), "处理方式"] = "暂不处理"

    def _default_standard_item(row: pd.Series) -> str:
        selected_item = str(row.get("选择标准项目", "")).strip()
        if selected_item in rules:
            return selected_item
        possible_item = str(row.get("可能归类", "")).strip()
        return possible_item if possible_item in rules else ""

    normalized_df["选择标准项目"] = normalized_df.apply(_default_standard_item, axis=1)
    return normalized_df[FEEDBACK_COLUMNS]


def make_activity_section_df(activity_sections: list[dict]) -> pd.DataFrame:
    rows = []
    seen_names: set[str] = set()
    for section in activity_sections:
        confidence = str(section.get("section_confidence", ""))
        section_level = str(section.get("section_level", "main" if confidence == "strong" else "sub"))
        if str(section.get("candidate_type", "activity")) != "activity":
            continue
        if section_level != "main" or confidence == "noise":
            continue
        standard_name = str(section.get("normalized_name", section.get("name", "")))
        if not standard_name or standard_name in seen_names:
            continue
        seen_names.add(standard_name)
        rows.append(
            {
                "是否作为报价板块": bool(section.get("selected", section_level == "main")),
                "层级": "主活动板块" if section_level == "main" else "疑似子活动",
                "标准板块名": standard_name,
                "原始识别文本": summarize_text(section.get("raw_name", section.get("name", "")), 40),
                "置信度": confidence,
                "判断原因": str(section.get("reason", "")),
                "_original_name": standard_name,
            }
        )
    if not rows:
        return pd.DataFrame(columns=SECTION_CONFIRM_COLUMNS)
    df = pd.DataFrame(rows)
    return df.set_index("_original_name", drop=True)


def make_sub_activity_section_df(activity_sections: list[dict]) -> pd.DataFrame:
    rows = []
    seen_names: set[tuple[str, str]] = set()
    for section in activity_sections:
        section_level = str(section.get("section_level", ""))
        if section_level != "sub" and str(section.get("candidate_type", "")) != "sub_activity":
            continue
        standard_name = str(section.get("normalized_name", section.get("name", ""))).strip()
        parent = str(section.get("parent", "")).strip()
        if not standard_name:
            continue
        key = (parent, standard_name)
        if key in seen_names:
            continue
        seen_names.add(key)
        rows.append(
            {
                "父级板块": parent or "待确认",
                "疑似子活动": standard_name,
                "原始识别文本": summarize_text(section.get("raw_name", section.get("name", "")), 40),
                "判断原因": str(section.get("reason", "")),
            }
        )
    return pd.DataFrame(rows, columns=["父级板块", "疑似子活动", "原始识别文本", "判断原因"])


def merge_activity_section_selection(
    section_df: pd.DataFrame,
    activity_sections: list[dict],
) -> tuple[pd.DataFrame, list[dict], list[dict]]:
    selection_map: dict[str, bool] = {}
    rename_map: dict[str, str] = {}
    if not section_df.empty:
        for index, row in section_df.iterrows():
            original_name = str(row.get("_original_name", index)).strip()
            renamed_name = str(row.get("标准板块名", "")).strip() or original_name
            selection_map[original_name] = bool(row.get("是否作为报价板块", False))
            rename_map[original_name] = renamed_name

    updated_sections = []
    for section in activity_sections:
        if str(section.get("candidate_type", "activity")) not in {"activity", "sub_activity"}:
            continue
        key = str(section.get("normalized_name", section.get("name", "")))
        section_level = str(section.get("section_level", "main" if section.get("section_confidence") == "strong" else "sub"))
        updated_section = dict(section)
        if section_level == "main" and str(section.get("section_confidence", "")) != "noise":
            selected = selection_map.get(key, bool(section.get("selected", True)))
            updated_section["selected"] = selected
            updated_section["name"] = rename_map.get(key, key)
            updated_section["normalized_name"] = rename_map.get(key, key)
        else:
            updated_section["selected"] = False
        updated_section["section_level"] = section_level
        updated_sections.append(updated_section)

    updated_df = make_activity_section_df(updated_sections)
    confirmed_sections = prepare_confirmed_activity_sections(updated_sections)
    return updated_df, updated_sections, confirmed_sections


def apply_activity_section_changes(
    quote_df: pd.DataFrame,
    section_df: pd.DataFrame,
    activity_sections_all: list[dict],
) -> tuple[pd.DataFrame, pd.DataFrame, list[dict], list[dict]]:
    raw_rename_map = {
        str(index): str(row.get("标准板块名", "")).strip() or str(index)
        for index, row in section_df.iterrows()
    }
    if quote_df.empty:
        updated_df, updated_sections, confirmed_sections = merge_activity_section_selection(section_df, activity_sections_all)
        return quote_df.copy(), updated_df, updated_sections, confirmed_sections

    updated_section_df, updated_sections, confirmed_sections = merge_activity_section_selection(section_df, activity_sections_all)
    confirmed_names = {str(section.get("normalized_name", section.get("name", ""))) for section in confirmed_sections}

    updated_quote_df = quote_df.copy()
    for index, row in updated_quote_df.iterrows():
        current_section = str(row.get("quote_section", row.get("项目分类", ""))).strip()
        renamed_section = raw_rename_map.get(current_section, current_section)
        remark = str(row.get("备注", "")).replace("；原归属板块已取消，请人工确认归属", "").replace("原归属板块已取消，请人工确认归属", "").strip("；")

        if current_section in raw_rename_map:
            updated_quote_df.at[index, "quote_section"] = renamed_section
            updated_quote_df.at[index, "项目分类"] = renamed_section
            current_section = renamed_section

        if current_section and current_section not in {"活动宣传", "美陈搭建类", "其他搭建类", "未归属板块", "人员类及其他"} and current_section not in confirmed_names:
            updated_quote_df.at[index, "quote_section"] = "未归属板块"
            updated_quote_df.at[index, "项目分类"] = "未归属板块"
            addition = "原归属板块已取消，请人工确认归属"
            updated_quote_df.at[index, "备注"] = f"{remark}；{addition}" if remark else addition
        else:
            updated_quote_df.at[index, "备注"] = remark

    updated_quote_df = sort_by_quote_section(dedupe_final_quote_items(updated_quote_df), confirmed_sections)
    return updated_quote_df, updated_section_df, updated_sections, confirmed_sections


def filter_candidate_df(candidate_df: pd.DataFrame) -> pd.DataFrame:
    if candidate_df.empty:
        return pd.DataFrame(columns=FEEDBACK_COLUMNS)

    rows = []
    for _, row in candidate_df.iterrows():
        candidate = str(row.get("候选词", ""))
        if not any(root in candidate for root in QUOTE_CANDIDATE_ROOTS):
            continue
        possible_item = infer_candidate_item(candidate)
        has_standard_item = possible_item in rules
        rows.append(
            {
                "是否处理": False,
                "候选词": candidate,
                "可能归类": possible_item,
                "上下文摘要": summarize_text(row.get("所在上下文", "")),
                "处理方式": "加入已有标准项目别名" if has_standard_item else "暂不处理",
                "选择标准项目": possible_item if has_standard_item else "",
                "备注": "",
            }
        )
    return pd.DataFrame(rows, columns=FEEDBACK_COLUMNS)


def confirmation_status(row: pd.Series) -> tuple[str, str]:
    source_status = str(row.get("来源状态", ""))
    quote_type = str(row.get("报价类型", ""))
    remark = str(row.get("备注", ""))
    item = str(row.get("标准项目", ""))
    category = str(row.get("项目分类", ""))

    if source_status == "系统推算":
        return "系统建议", "系统补全，需确认是否需要"
    if "未明确数量" in remark or "数量" in remark:
        return "需确认数量", "缺数量"
    if any(key in item for key in ["人员", "主持", "节目演出", "赛事"]) or category == "人员执行类":
        return "需确认数量", "缺人数/场次"
    if any(key in item for key in ["点位", "装置", "展板", "导视"]):
        return "需确认规格", "缺点位数量"
    if source_status == "需确认" or quote_type in {"档位报价", "模糊报价"}:
        return "需确认规格", "缺尺寸/材质"
    return "已确认", ""


def source_basis(row: pd.Series | dict) -> str:
    evidence_type = str(row.get("evidence_type", "")).strip()
    evidence_text = str(row.get("evidence_text", "")).strip()
    trigger_module = str(row.get("trigger_module", "")).strip()
    if evidence_type == "explicit_text":
        return f"原文命中：{summarize_text(evidence_text or row.get('原始命中词', ''), 32)}"
    if evidence_type == "module_completion":
        return f"模块补全：{trigger_module or evidence_text.replace('由【', '').replace('】模块补全', '')}"
    if evidence_type == "user_selected_suggestion":
        return "用户补充：建议补充项"
    return f"待确认来源：{evidence_type or 'unknown'}"


def filter_official_quote_rows(full_df: pd.DataFrame) -> pd.DataFrame:
    if full_df.empty:
        return full_df.copy()
    valid_evidence = {"explicit_text", "module_completion", "user_selected_suggestion"}
    filtered_df = full_df[full_df["是否保留"].astype(bool)].copy()
    if "evidence_type" not in filtered_df.columns:
        return filtered_df.iloc[0:0].copy()
    filtered_df = filtered_df[filtered_df["evidence_type"].isin(valid_evidence)]
    evidence_text = filtered_df.get("evidence_text", pd.Series("", index=filtered_df.index)).astype(str).str.strip()
    filtered_df = filtered_df[filtered_df["evidence_type"].ne("explicit_text") | evidence_text.ne("")]
    confirmed_modules = {
        str(section.get("name", "")).strip()
        for section in st.session_state.get("activity_sections", [])
        if str(section.get("name", "")).strip()
    }
    if "trigger_module" not in filtered_df.columns:
        return filtered_df[filtered_df["evidence_type"].ne("module_completion")]
    module_ok = (
        filtered_df["trigger_module"].astype(str).str.strip().isin(confirmed_modules)
        & filtered_df.get("evidence_text", pd.Series("", index=filtered_df.index)).astype(str).str.strip().ne("")
    )
    filtered_df = filtered_df[filtered_df["evidence_type"].ne("module_completion") | module_ok]
    return filtered_df


def make_judgement_df(full_df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    official_df = filter_official_quote_rows(full_df)
    for index, row in sort_by_quote_section(official_df, st.session_state.get("activity_sections", [])).iterrows():
        status, need = confirmation_status(row)
        rows.append(
            {
                "_row_id": index,
                "是否加入": bool(row.get("是否保留", True)),
                "项目分类": row.get("项目分类", ""),
                "项目": row.get("标准项目", ""),
                "内容/尺寸/工艺": sanitize_quote_content_text(row.get("内容/尺寸/工艺", ""), row.get("标准项目", "")),
                "数量": row.get("数量", 1),
                "单位": row.get("单位", ""),
                "确认状态": status,
                "需要确认什么": need,
                "来源依据": source_basis(row),
            }
        )
    if not rows:
        return pd.DataFrame()
    return pd.DataFrame(rows, columns=["_row_id", "是否加入", "项目分类", "项目", "内容/尺寸/工艺", "数量", "单位", "确认状态", "需要确认什么", "来源依据"]).set_index("_row_id")


def merge_judgement_edits(full_df: pd.DataFrame, edited_df: pd.DataFrame) -> pd.DataFrame:
    full_df = full_df.copy()
    for index, row in edited_df.iterrows():
        if index not in full_df.index:
            continue
        full_df.at[index, "是否保留"] = bool(row.get("是否加入", True))
        full_df.at[index, "标准项目"] = row.get("项目", "")
        full_df.at[index, "项目分类"] = row.get("项目分类", "")
        full_df.at[index, "quote_section"] = row.get("项目分类", "")
        full_df.at[index, "内容/尺寸/工艺"] = sanitize_quote_content_text(row.get("内容/尺寸/工艺", ""), row.get("项目", ""))
        full_df.at[index, "数量"] = row.get("数量", 1)
        full_df.at[index, "单位"] = row.get("单位", "")
    return dedupe_final_quote_items(recalculate_totals(full_df))


def make_final_quote_df(full_df: pd.DataFrame) -> pd.DataFrame:
    if full_df.empty:
        return pd.DataFrame(columns=FINAL_QUOTE_COLUMNS)
    filtered_df = filter_official_quote_rows(full_df)
    if filtered_df.empty:
        return pd.DataFrame(columns=FINAL_QUOTE_COLUMNS)
    kept_df = sort_by_quote_section(filtered_df, st.session_state.get("activity_sections", []))
    final_df = kept_df[["项目分类", "标准项目", "内容/尺寸/工艺", "数量", "单位", "单价", "合计", "备注"]].rename(columns={"标准项目": "项目"})
    return dedupe_final_quote_items(final_df)


def merge_final_quote_edits(full_df: pd.DataFrame, edited_df: pd.DataFrame) -> pd.DataFrame:
    full_df = full_df.copy()
    for index, row in edited_df.iterrows():
        if index not in full_df.index:
            continue
        full_df.at[index, "项目分类"] = row.get("项目分类", "")
        full_df.at[index, "quote_section"] = row.get("项目分类", "")
        full_df.at[index, "标准项目"] = row.get("项目", "")
        full_df.at[index, "内容/尺寸/工艺"] = sanitize_quote_content_text(row.get("内容/尺寸/工艺", ""), row.get("项目", ""))
        full_df.at[index, "数量"] = row.get("数量", 1)
        full_df.at[index, "单位"] = row.get("单位", "")
        full_df.at[index, "单价"] = row.get("单价", 0)
        full_df.at[index, "备注"] = row.get("备注", "")
    return dedupe_final_quote_items(recalculate_totals(full_df))


def split_suggestions(suggested_df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    if suggested_df.empty:
        return suggested_df, suggested_df
    item_col = suggested_df["建议补充项"].astype(str)
    key_df = suggested_df[item_col.isin(KEY_SUGGESTION_ITEMS)].copy()
    optional_df = suggested_df[~item_col.isin(KEY_SUGGESTION_ITEMS)].copy()
    return key_df, optional_df


def suggestion_display(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=["项目", "项目分类", "需要确认什么"])
    return pd.DataFrame(
        {
            "项目": df["建议补充项"],
            "项目分类": df["项目分类"],
            "需要确认什么": df["备注"].map(lambda value: summarize_text(value, 28)),
        }
    )


def make_suggestion_editor_df(suggested_df: pd.DataFrame) -> pd.DataFrame:
    if suggested_df.empty:
        return pd.DataFrame(columns=SUGGESTION_EDITOR_COLUMNS)

    rows = []
    for _, row in suggested_df.iterrows():
        standard_item = str(row.get("建议补充项", row.get("建议项目", "")))
        rule = rules.get(standard_item, {})
        rows.append(
            {
                "是否加入报价单": bool(row.get("是否加入报价单", False)),
                "建议项目": standard_item,
                "项目分类": row.get("项目分类", ""),
                "内容/尺寸/工艺": sanitize_quote_content_text(row.get("备注", "") or rule.get("default_desc", ""), standard_item),
                "数量": row.get("数量", 1) or 1,
                "单位": row.get("单位", "") or rule.get("default_unit", "项"),
                "报价类型": row.get("报价类型", "") or rule.get("quote_type", ""),
                "来源说明": row.get("建议原因", ""),
                "备注": row.get("备注", ""),
            }
        )
    return pd.DataFrame(rows, columns=SUGGESTION_EDITOR_COLUMNS)


def final_quote_export_df(final_quote_df: pd.DataFrame) -> pd.DataFrame:
    if final_quote_df.empty:
        return pd.DataFrame(columns=QUOTE_COLUMNS)

    export_df = final_quote_df.copy()
    if "是否加入" in export_df.columns:
        export_df = export_df[export_df["是否加入"].astype(bool)]
    if "是否保留" not in export_df.columns:
        export_df["是否保留"] = True
    if "标准项目" not in export_df.columns and "项目" in export_df.columns:
        export_df["标准项目"] = export_df["项目"]
    if "quote_section" not in export_df.columns:
        export_df["quote_section"] = export_df["项目分类"].astype(str)
    for column in QUOTE_COLUMNS:
        if column not in export_df.columns:
            export_df[column] = ""
    export_df = dedupe_final_quote_items(export_df)
    export_df = recalculate_totals(export_df)
    return sort_by_quote_section(export_df, st.session_state.get("activity_sections", []))


def recalculate_final_quote_df(final_quote_df: pd.DataFrame) -> pd.DataFrame:
    if final_quote_df.empty:
        return final_quote_df.copy()
    recalculated = final_quote_df.copy()
    for index, row in recalculated.iterrows():
        try:
            quantity = float(row.get("数量", 0))
            unit_price = float(row.get("单价", 0))
        except (TypeError, ValueError):
            recalculated.at[index, "合计"] = 0
            continue
        recalculated.at[index, "合计"] = quantity * unit_price
    return dedupe_final_quote_items(recalculated)


def quote_display_dedupe_key(row: pd.Series | dict) -> str:
    item = str(row.get("标准项目", row.get("项目", ""))).strip()
    section = str(row.get("quote_section", row.get("项目分类", ""))).strip()
    key = make_quote_item_key({"标准项目": item, "项目分类": section, "quote_section": section})
    if key.startswith("PUBLIC::"):
        return key
    return f"{section}::{item}"


def build_suggestion_quote_row(row: pd.Series, activity_sections: list[dict]) -> dict:
    standard_item = str(row.get("建议项目", "")).strip()
    rule = rules.get(standard_item, {})
    quote_section = str(row.get("项目分类", "")).strip()
    item = {
        "是否保留": True,
        "项目分类": quote_section or rule.get("category", ""),
        "标准项目": standard_item,
        "原始命中词": "建议补充项",
        "内容/尺寸/工艺": sanitize_quote_content_text(row.get("内容/尺寸/工艺", "") or rule.get("default_desc", ""), standard_item),
        "数量": row.get("数量", 1) or 1,
        "单位": row.get("单位", "") or rule.get("default_unit", "项"),
        "单价": 0,
        "合计": 0,
        "报价类型": row.get("报价类型", "") or rule.get("quote_type", ""),
        "来源状态": "需确认",
        "备注": row.get("备注", ""),
        "internal_category": "",
        "quote_section": quote_section,
        "quote_section_order": "",
        "匹配位置": [],
        "命中模块": [],
        "source_context_start": "",
        "source_context_text": "",
        "matched_section_name": "",
        "evidence_type": "user_selected_suggestion",
        "evidence_text": row.get("来源说明", ""),
        "trigger_module": "",
        "trigger_section": quote_section,
    }
    if not item["quote_section"]:
        item["quote_section"] = assign_quote_section(item, None, "", activity_sections)
        item["项目分类"] = item["quote_section"] or "未归属板块"
        item["trigger_section"] = item["quote_section"]
    if not item["quote_section"]:
        item["quote_section"] = "未归属板块"
        item["项目分类"] = "未归属板块"
        item["trigger_section"] = "未归属板块"
    return item


def apply_selected_suggestions(suggestion_editor_df: pd.DataFrame) -> tuple[int, list[str]]:
    if suggestion_editor_df.empty:
        return 0, []

    selected_df = suggestion_editor_df[suggestion_editor_df["是否加入报价单"].astype(bool)].copy()
    if selected_df.empty:
        return 0, []

    activity_sections = st.session_state.get("activity_sections", [])
    working_df = st.session_state.get("working_quote_df", st.session_state.quote_df).copy()
    final_df = st.session_state.get("final_quote_df", make_final_quote_df(working_df)).copy()
    existing_keys = {quote_display_dedupe_key(row) for _, row in final_df.iterrows()}

    rows_to_add = []
    duplicates = []
    for _, suggestion in selected_df.iterrows():
        quote_row = build_suggestion_quote_row(suggestion, activity_sections)
        key = quote_display_dedupe_key(quote_row)
        item_name = str(quote_row.get("标准项目", ""))
        if key in existing_keys:
            duplicates.append(item_name)
            continue
        existing_keys.add(key)
        rows_to_add.append(quote_row)

    if rows_to_add:
        add_df = pd.DataFrame(rows_to_add)
        for column in QUOTE_COLUMNS:
            if column not in add_df.columns:
                add_df[column] = ""
        working_df = pd.concat([working_df, add_df], ignore_index=True)
        working_df = sort_by_quote_section(recalculate_totals(working_df), activity_sections)
        refreshed_suggestions = suggestion_editor_df.copy()
        refreshed_suggestions["是否加入报价单"] = False
        st.session_state["suggestion_editor_df"] = refreshed_suggestions
        sync_working_quote_state(working_df, activity_sections)
        st.session_state["quote_editor_version"] = st.session_state.get("quote_editor_version", 0) + 1
        reset_editor_widget_state()
        st.session_state["excel_bytes"] = None

    return len(rows_to_add), duplicates


def run_recognition(current_text: str) -> None:
    reset_editor_widget_state()
    current_rules = load_rules(RULES_PATH)
    ignored_terms = load_ignored_terms(IGNORED_TERMS_PATH)
    extracted_rows = extract_quote_items(current_text, current_rules)
    activity_content_diagnostics = diagnose_activity_content_ranges(current_text)
    section_candidate_diagnostics = pd.DataFrame(diagnose_section_candidates(current_text))
    all_activity_sections = extract_activity_sections(current_text, extracted_rows)
    activity_section_df = make_activity_section_df(all_activity_sections)
    activity_section_df, all_activity_sections, activity_sections = merge_activity_section_selection(activity_section_df, all_activity_sections)
    quote_rows = build_quote_rows(extracted_rows, PRICE_DB_PATH, current_text, activity_sections=all_activity_sections)
    activity_types = classify_activity_types(current_text)
    suggested_rows = build_suggested_items(activity_types, quote_rows, current_rules, current_text)
    candidate_rows = extract_unrecognized_candidates(current_text, quote_rows, current_rules, ignored_terms)

    quote_df = sort_by_quote_section(pd.DataFrame(quote_rows), activity_sections)
    candidate_df = pd.DataFrame(candidate_rows, columns=["候选词", "所在上下文", "建议动作", "是否加入规则库"])
    candidate_display_df = filter_candidate_df(candidate_df)
    activity_type_df = pd.DataFrame(activity_types).drop(columns=["focus_items"], errors="ignore")
    suggested_df = sort_by_quote_section(pd.DataFrame(
        suggested_rows,
        columns=["是否加入报价单", "建议补充项", "项目分类", "报价类型", "来源状态", "建议原因", "备注"],
    ), activity_sections)
    coverage = calculate_coverage(len(quote_df), len(candidate_display_df))

    st.session_state.raw_quote_df = quote_df.copy()
    st.session_state.candidate_df = candidate_df.copy()
    st.session_state.candidate_display_df = candidate_display_df.copy()
    st.session_state.feedback_df = candidate_display_df.copy()
    st.session_state.activity_type_df = activity_type_df.copy()
    st.session_state.activity_sections = activity_sections
    st.session_state.activity_sections_all = all_activity_sections
    st.session_state.activity_content_diagnostics = activity_content_diagnostics
    st.session_state.section_candidate_diagnostics = section_candidate_diagnostics
    st.session_state.activity_section_df = activity_section_df.copy()
    st.session_state.section_editor_df = activity_section_df.copy()
    st.session_state.suggested_df = suggested_df.copy()
    st.session_state.suggestion_editor_df = make_suggestion_editor_df(suggested_df)
    st.session_state.coverage = coverage
    st.session_state.excel_bytes = None
    st.session_state.last_text_hash = get_text_hash(current_text)
    st.session_state.raw_text = current_text
    st.session_state.confirmed_sections = [section.get("name", "") for section in activity_sections]
    sync_working_quote_state(quote_df, activity_sections)
    st.session_state["quote_editor_version"] = st.session_state.get("quote_editor_version", 0) + 1
