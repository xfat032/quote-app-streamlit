from __future__ import annotations

import hashlib
import json
from pathlib import Path

import pandas as pd
import streamlit as st

from core.activity_classifier import build_suggested_items, classify_activity_types
from core.excel_exporter import export_quote_to_excel
from core.extractor import calculate_coverage, extract_quote_items, extract_unrecognized_candidates
from core.normalizer import QUOTE_TYPES, load_rules
from core.quote_builder import (
    QUOTE_COLUMNS,
    assign_quote_section,
    build_quote_rows,
    diagnose_activity_content_ranges,
    diagnose_section_candidates,
    dedupe_final_quote_items,
    ensure_price_db,
    extract_activity_sections,
    make_quote_item_key,
    prepare_confirmed_activity_sections,
    recalculate_totals,
    reassign_quote_sections,
    sort_by_quote_section,
    sort_quote_items,
)
from core.rule_feedback import apply_feedback_rows, ensure_ignored_terms, load_ignored_terms
from core.text_reader import read_text_from_upload


BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
OUTPUT_DIR = BASE_DIR / "output"
RULES_PATH = DATA_DIR / "rules_config.json"
PRICE_DB_PATH = DATA_DIR / "price_db.xlsx"
IGNORED_TERMS_PATH = DATA_DIR / "ignored_terms.json"
EXPORT_PATH = OUTPUT_DIR / "报价单.xlsx"

NO_QUOTE_TERMS = [
    "活动背景",
    "活动概述",
    "活动名称",
    "活动时间",
    "活动地点",
    "主会场",
    "分会场",
    "指导单位",
    "主办单位",
    "承办单位",
    "协办单位",
    "活动目标",
    "参与对象",
    "预计参与人数",
    "项目背景",
    "项目分析",
    "品牌定位",
    "传播目标",
    "经济目标",
    "文化目标",
    "感谢观看",
    "目录",
    "章节标题",
    "PART",
    "CONTENTS",
]

QUOTE_CANDIDATE_ROOTS = [
    "舞台",
    "灯光",
    "音响",
    "视频",
    "摄影",
    "直播",
    "海报",
    "推文",
    "媒体",
    "达人",
    "KOL",
    "市集",
    "摊位",
    "帐篷",
    "展板",
    "展览",
    "装置",
    "打卡",
    "导视",
    "指引",
    "互动",
    "游戏",
    "道具",
    "礼品",
    "奖品",
    "人员",
    "主持",
    "安保",
    "医疗",
    "志愿者",
    "交通",
    "停车",
    "电力",
    "发电",
    "围挡",
    "铁马",
    "餐费",
    "饮用水",
    "运输",
    "手作",
    "体验",
    "花灯",
    "灯笼",
    "河灯",
    "巡游",
    "NPC",
    "服饰",
    "妆造",
    "长桌宴",
    "茶席",
    "阅读",
    "图书",
    "二维码",
]

KEY_SUGGESTION_ITEMS = {
    "电力保障",
    "演艺服化道",
    "NPC互动服务",
    "美陈装置",
    "互动道具",
    "安保人员",
    "医疗保障",
    "交通停车引导",
    "交通引导",
    "灯光艺术装置",
    "启动仪式道具",
    "技术人员",
    "救生员",
    "铁马围挡",
    "应急物资",
}

OPTIONAL_SUGGESTION_ITEMS = {"餐费", "饮用水", "运输费", "工作证", "对讲机", "执行人员", "志愿者"}
FINAL_QUOTE_COLUMNS = ["项目分类", "项目", "内容/尺寸/工艺", "数量", "单位", "单价", "合计", "备注"]
FEEDBACK_COLUMNS = ["是否处理", "候选词", "可能归类", "上下文摘要", "处理方式", "选择标准项目", "备注"]
SECTION_CONFIRM_COLUMNS = ["是否作为报价板块", "层级", "标准板块名", "原始识别文本", "置信度", "判断原因"]
SUGGESTION_EDITOR_COLUMNS = ["是否加入报价单", "建议项目", "项目分类", "内容/尺寸/工艺", "数量", "单位", "报价类型", "来源说明", "备注"]


st.set_page_config(page_title="活动方案报价单生成器", page_icon="📋", layout="wide")

DATA_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
ensure_price_db(PRICE_DB_PATH)
ensure_ignored_terms(IGNORED_TERMS_PATH)


def append_default_ignored_terms() -> None:
    existing = load_ignored_terms(IGNORED_TERMS_PATH)
    updated = list(dict.fromkeys([*existing, *NO_QUOTE_TERMS]))
    if updated != existing:
        IGNORED_TERMS_PATH.write_text(json.dumps(updated, ensure_ascii=False, indent=2), encoding="utf-8")


append_default_ignored_terms()
rules = load_rules(RULES_PATH)

st.title("活动方案报价单生成器")
st.caption("输入方案 → 识别报价项 → 人工判断 → 编辑报价单 → 导出 Excel")
if st.session_state.get("feedback_rerun_notice"):
    st.success(st.session_state.feedback_rerun_notice)
    st.session_state.feedback_rerun_notice = ""

if "plan_text_input" not in st.session_state:
    st.session_state["plan_text_input"] = ""
if "plan_text_value" not in st.session_state:
    st.session_state["plan_text_value"] = ""
if "quote_df" not in st.session_state:
    st.session_state.quote_df = pd.DataFrame(columns=QUOTE_COLUMNS)
if "raw_quote_df" not in st.session_state:
    st.session_state.raw_quote_df = pd.DataFrame(columns=QUOTE_COLUMNS)
if "working_quote_df" not in st.session_state:
    st.session_state.working_quote_df = pd.DataFrame(columns=QUOTE_COLUMNS)
if "review_quote_df" not in st.session_state:
    st.session_state.review_quote_df = pd.DataFrame(columns=QUOTE_COLUMNS)
if "recognized_df" not in st.session_state:
    st.session_state.recognized_df = pd.DataFrame(columns=QUOTE_COLUMNS)
if "candidate_df" not in st.session_state:
    st.session_state.candidate_df = pd.DataFrame(columns=["候选词", "所在上下文", "建议动作", "是否加入规则库"])
if "candidate_display_df" not in st.session_state:
    st.session_state.candidate_display_df = pd.DataFrame(columns=FEEDBACK_COLUMNS)
if "feedback_df" not in st.session_state:
    st.session_state.feedback_df = pd.DataFrame(columns=FEEDBACK_COLUMNS)
if "activity_type_df" not in st.session_state:
    st.session_state.activity_type_df = pd.DataFrame(columns=["活动类型", "命中关键词", "重点规则"])
if "activity_sections" not in st.session_state:
    st.session_state.activity_sections = []
if "activity_sections_all" not in st.session_state:
    st.session_state.activity_sections_all = []
if "activity_content_diagnostics" not in st.session_state:
    st.session_state.activity_content_diagnostics = {}
if "activity_section_df" not in st.session_state:
    st.session_state.activity_section_df = pd.DataFrame(columns=SECTION_CONFIRM_COLUMNS)
if "suggested_df" not in st.session_state:
    st.session_state.suggested_df = pd.DataFrame(columns=["是否加入报价单", "建议补充项", "项目分类", "报价类型", "来源状态", "建议原因", "备注"])
if "suggestion_editor_df" not in st.session_state:
    st.session_state.suggestion_editor_df = pd.DataFrame(columns=SUGGESTION_EDITOR_COLUMNS)
if "coverage" not in st.session_state:
    st.session_state.coverage = None
if "current_file_signature" not in st.session_state:
    st.session_state.current_file_signature = ""
if "last_text_hash" not in st.session_state:
    st.session_state.last_text_hash = ""
if "excel_bytes" not in st.session_state:
    st.session_state.excel_bytes = None
if "upload_notice" not in st.session_state:
    st.session_state.upload_notice = ""
if "quote_editor_version" not in st.session_state:
    st.session_state.quote_editor_version = 0


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
                "内容/尺寸/工艺": row.get("内容/尺寸/工艺", ""),
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
        full_df.at[index, "内容/尺寸/工艺"] = row.get("内容/尺寸/工艺", "")
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
        full_df.at[index, "内容/尺寸/工艺"] = row.get("内容/尺寸/工艺", "")
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
                "内容/尺寸/工艺": row.get("备注", "") or rule.get("default_desc", ""),
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
        "内容/尺寸/工艺": row.get("内容/尺寸/工艺", "") or rule.get("default_desc", ""),
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


st.header("1. 方案输入区")
col_a, col_b, col_c = st.columns(3)
with col_a:
    activity_name = st.text_input("活动名称", placeholder="例如：农野开心乐")
with col_b:
    activity_time = st.text_input("活动时间", placeholder="例如：2026年6月")
with col_c:
    client_name = st.text_input("单位名称", placeholder="例如：某某单位")

uploaded_file = st.file_uploader("上传方案文件（支持 .txt / .docx / .pdf）", type=["txt", "docx", "pdf"])
if uploaded_file is not None:
    new_signature = get_file_signature(uploaded_file)
    if st.session_state.get("current_file_signature") != new_signature:
        try:
            clear_result_state()
            extracted_text = read_text_from_upload(uploaded_file)
            st.session_state["plan_text_input"] = extracted_text
            st.session_state["plan_text_value"] = extracted_text
            st.session_state["current_file_signature"] = new_signature
            st.session_state["last_text_hash"] = get_text_hash(extracted_text)
            st.session_state["upload_notice"] = "已读取上传文件，并填入方案文本区域。"
            st.rerun()
        except Exception as exc:
            st.error(f"文件读取失败：{exc}")

if st.session_state.get("upload_notice"):
    st.success(st.session_state["upload_notice"])
    st.session_state["upload_notice"] = ""

plan_text = st.text_area("方案文本", key="plan_text_input", height=220, placeholder="粘贴活动方案文本，或上传文件自动读取。")
st.session_state["plan_text_value"] = plan_text
col_recognize, col_rerun = st.columns([1, 1])
with col_recognize:
    recognize_clicked = st.button("识别报价项", type="primary")
with col_rerun:
    rerun_clicked = st.button("重新识别当前方案")

if recognize_clicked or rerun_clicked:
    plan_text = st.session_state.get("plan_text_value", "")
    if not plan_text.strip():
        st.warning("请先输入或上传方案文本。")
    else:
        current_text_hash = get_text_hash(plan_text)
        if current_text_hash != st.session_state.get("last_text_hash", ""):
            clear_result_state()
            st.session_state["plan_text_value"] = plan_text
            st.session_state.last_text_hash = current_text_hash
        run_recognition(plan_text)
        st.success("已完成识别。" if recognize_clicked else "已使用最新规则重新识别当前方案。")

if st.session_state.coverage is not None:
    st.header("2. 识别结果摘要")
    working_quote_df = st.session_state.get("working_quote_df", st.session_state.quote_df)
    judgement_df = make_judgement_df(working_quote_df)
    confirm_count = 0 if judgement_df.empty else int((judgement_df["确认状态"] != "已确认").sum())
    candidate_count = len(st.session_state.candidate_display_df)

    metric_a, metric_b, metric_c, metric_d = st.columns(4)
    metric_a.metric("已识别报价项", f"{len(working_quote_df)} 项")
    metric_b.metric("需人工确认", f"{confirm_count} 项")
    metric_c.metric("建议补充项", f"{len(st.session_state.suggested_df)} 项")
    metric_d.metric("未识别候选", f"{candidate_count} 项")

    if st.session_state.activity_type_df.empty:
        st.write("识别到的活动类型：暂未命中明确模板")
    else:
        activity_types = " / ".join(st.session_state.activity_type_df["活动类型"].astype(str).tolist())
        st.write(f"识别到的活动类型：{activity_types}")

    diagnostics = st.session_state.get("activity_content_diagnostics", {})
    with st.expander("活动内容区域调试信息", expanded=False):
        ranges = diagnostics.get("ranges", []) if isinstance(diagnostics, dict) else []
        if not ranges:
            st.warning("未定位到活动内容区域，已启用全文强标题、白名单、编号标题、连续短标题和报价项反推兜底。")
        else:
            debug_rows = [
                {
                    "区域": item.get("index", index + 1),
                    "起点": item.get("start", 0),
                    "终点": item.get("end", 0),
                    "截取文本长度": item.get("length", 0),
                    "疑似误用了目录页": "是" if item.get("directory_like") else "否",
                    "提示": "；".join(item.get("warnings", [])),
                }
                for index, item in enumerate(ranges)
            ]
            st.dataframe(pd.DataFrame(debug_rows), use_container_width=True, hide_index=True)
            if bool(diagnostics.get("has_short_range")):
                st.warning("活动内容区域可能切片错误")
            if any(bool(item.get("directory_like")) for item in ranges):
                st.warning("活动内容区域疑似误用了目录页")
    filtered_candidate_df = st.session_state.get("section_candidate_diagnostics", pd.DataFrame())
    if isinstance(filtered_candidate_df, pd.DataFrame) and not filtered_candidate_df.empty:
        with st.expander("被过滤的活动板块候选", expanded=False):
            st.dataframe(filtered_candidate_df, use_container_width=True, hide_index=True)

    section_df = st.session_state.activity_section_df.copy()
    sub_section_df = make_sub_activity_section_df(st.session_state.get("activity_sections_all", []))
    if not section_df.empty and "层级" not in section_df.columns:
        section_df["层级"] = section_df["置信度"].map(lambda value: "主活动板块" if str(value) == "strong" else "疑似子活动")
    strong_sections = []
    section_candidate_count = 0
    if not section_df.empty:
        strong_sections = section_df[
            (section_df["层级"] == "主活动板块") & (section_df["是否作为报价板块"].astype(bool))
        ]["标准板块名"].astype(str).tolist()
        section_candidate_count = len(sub_section_df)

    if strong_sections:
        lines = [f"{index}. {section}" for index, section in enumerate(strong_sections, start=1)]
        st.write("识别到的活动板块：")
        for line in lines:
            st.write(line)
        if section_candidate_count:
            st.caption(f"疑似活动板块：{section_candidate_count} 个，点击展开确认")
    else:
        st.info("未识别到明确活动板块，已根据项目内容反推报价板块，请人工检查“未归属板块”。")

    if not section_df.empty:
        with st.expander("活动板块确认区", expanded=False):
            if "section_editor_df" not in st.session_state:
                st.session_state["section_editor_df"] = section_df.copy()
            editor_source_df = st.session_state.get("section_editor_df", section_df.copy())
            if isinstance(editor_source_df, pd.DataFrame) and not editor_source_df.empty:
                editor_source_df = editor_source_df.copy()
                for column in SECTION_CONFIRM_COLUMNS:
                    if column not in editor_source_df.columns:
                        editor_source_df[column] = ""
            else:
                editor_source_df = section_df.copy()

            main_preview = editor_source_df[editor_source_df["层级"] == "主活动板块"]["标准板块名"].astype(str).tolist()
            if main_preview:
                st.write("主活动板块：")
                for section_name in main_preview:
                    st.write(f"- {section_name}")
            if not sub_section_df.empty:
                with st.expander("疑似子活动（默认不作为报价板块）", expanded=False):
                    st.dataframe(sub_section_df, use_container_width=True, hide_index=True)

            with st.form("activity_section_form"):
                edited_section_df = st.data_editor(
                    editor_source_df[SECTION_CONFIRM_COLUMNS],
                    use_container_width=True,
                    hide_index=True,
                    num_rows="fixed",
                    disabled=["层级", "原始识别文本", "置信度", "判断原因"],
                    column_config={
                        "是否作为报价板块": st.column_config.CheckboxColumn("是否作为报价板块", default=False),
                    },
                    key="activity_section_editor",
                )
                apply_sections = st.form_submit_button("应用活动板块修改")

            if apply_sections:
                st.session_state["section_editor_df"] = edited_section_df.copy()
                base_quote_df = st.session_state.get("working_quote_df", st.session_state.get("raw_quote_df", st.session_state.quote_df)).copy()
                updated_quote_df, updated_section_df, updated_sections, confirmed_sections = apply_activity_section_changes(
                    base_quote_df,
                    edited_section_df,
                    st.session_state.get("activity_sections_all", []),
                )
                st.session_state.activity_section_df = updated_section_df.copy()
                st.session_state.section_editor_df = updated_section_df.copy()
                st.session_state.activity_sections_all = updated_sections
                st.session_state.activity_sections = confirmed_sections
                st.session_state.confirmed_sections = [section.get("name", "") for section in confirmed_sections]
                st.session_state.suggested_df = sort_by_quote_section(st.session_state.suggested_df, confirmed_sections)
                sync_working_quote_state(updated_quote_df, confirmed_sections)
                st.session_state["quote_editor_version"] = st.session_state.get("quote_editor_version", 0) + 1
                st.session_state.excel_bytes = None
                reset_editor_widget_state()
                st.success("已应用活动板块修改，最终报价单已更新。")
                st.rerun()

    if st.session_state.coverage < 0.6:
        st.warning("当前方案存在较多疑似报价内容未进入报价单，请优先检查：1. 未识别候选清单 2. 建议补充项 3. 需确认报价项")

    st.header("3. 需要你判断的项目")
    working_quote_df = st.session_state.get("working_quote_df", st.session_state.quote_df)
    if working_quote_df.empty:
        st.info("暂未识别到报价项，可以在规则配置中补充同义词后重试。")
    else:
        st.subheader("报价项保留与确认")
        review_editor_df = st.session_state.get("review_quote_df", judgement_df.copy())
        with st.form("review_quote_form"):
            edited_judgement_df = st.data_editor(
                review_editor_df,
                use_container_width=True,
                hide_index=True,
                num_rows="fixed",
                disabled=["来源依据"],
                column_config={
                    "是否加入": st.column_config.CheckboxColumn("是否加入", default=True),
                    "数量": st.column_config.NumberColumn("数量", min_value=0.0, step=1.0),
                    "确认状态": st.column_config.SelectboxColumn("确认状态", options=["已确认", "需确认数量", "需确认规格", "需确认是否报价", "系统建议"]),
                },
                key="review_quote_editor",
            )
            review_submitted = st.form_submit_button("应用本表修改")
        if review_submitted:
            st.session_state["review_quote_df"] = edited_judgement_df.copy()
            updated_working_df = merge_judgement_edits(working_quote_df, edited_judgement_df)
            st.session_state["working_quote_df"] = updated_working_df.copy()
            st.session_state["quote_df"] = updated_working_df.copy()
            st.session_state["recognized_df"] = updated_working_df.copy()
            updated_final_df = make_final_quote_df(updated_working_df)
            st.session_state["final_quote_df"] = updated_final_df.copy()
            st.session_state["confirmed_quote_df"] = updated_final_df.copy()
            st.session_state["excel_bytes"] = None

    st.subheader("建议补充项")
    if st.session_state.suggested_df.empty:
        st.caption("暂无建议补充项。")
    else:
        suggestion_editor_source = st.session_state.get("suggestion_editor_df")
        if not isinstance(suggestion_editor_source, pd.DataFrame) or suggestion_editor_source.empty:
            suggestion_editor_source = make_suggestion_editor_df(st.session_state.suggested_df)
        edited_suggestion_df = st.data_editor(
            suggestion_editor_source,
            use_container_width=True,
            hide_index=True,
            num_rows="fixed",
            column_config={
                "是否加入报价单": st.column_config.CheckboxColumn("是否加入报价单", default=False),
                "数量": st.column_config.NumberColumn("数量", min_value=0.0, step=1.0),
                "报价类型": st.column_config.SelectboxColumn("报价类型", options=list(QUOTE_TYPES)),
            },
            key="suggestion_quote_editor",
        )
        st.session_state.suggestion_editor_df = edited_suggestion_df.copy()
        if st.button("应用建议补充项"):
            added_count, duplicate_items = apply_selected_suggestions(edited_suggestion_df)
            if duplicate_items:
                st.warning(f"已存在 {len(duplicate_items)} 项，未重复添加：{'、'.join(dict.fromkeys(duplicate_items))}")
            if added_count:
                st.success(f"已加入 {added_count} 个建议补充项。")
                st.rerun()
            elif not duplicate_items:
                st.info("请先勾选需要加入报价单的建议补充项。")

    st.header("4. 最终报价单编辑区")
    working_quote_df = st.session_state.get("working_quote_df", st.session_state.quote_df)
    final_quote_df = st.session_state.get("final_quote_df")
    if final_quote_df is None or final_quote_df.empty and not working_quote_df.empty:
        final_quote_df = make_final_quote_df(working_quote_df)
        st.session_state.final_quote_df = final_quote_df.copy()
        st.session_state.confirmed_quote_df = final_quote_df.copy()
    elif final_quote_df is not None and not final_quote_df.empty:
        final_quote_df = recalculate_final_quote_df(final_quote_df)
        st.session_state.final_quote_df = final_quote_df.copy()
        st.session_state.confirmed_quote_df = final_quote_df.copy()
    if final_quote_df.empty:
        st.info("当前没有勾选进入报价单的项目。")
    else:
        with st.form("final_quote_form"):
            edited_final_df = st.data_editor(
                final_quote_df,
                use_container_width=True,
                hide_index=True,
                num_rows="fixed",
                disabled=["合计"],
                column_config={
                    "数量": st.column_config.NumberColumn("数量", min_value=0.0, step=1.0),
                    "单价": st.column_config.NumberColumn("单价", min_value=0.0, step=1.0, format="%.2f"),
                    "合计": st.column_config.NumberColumn("合计", disabled=True, format="%.2f"),
                },
                key="final_quote_editor",
            )
            final_submitted = st.form_submit_button("保存报价单编辑")
        if final_submitted:
            edited_final_df = recalculate_final_quote_df(edited_final_df)
            st.session_state["final_quote_df"] = edited_final_df.copy()
            updated_working_df = merge_final_quote_edits(working_quote_df, edited_final_df)
            st.session_state["working_quote_df"] = updated_working_df.copy()
            st.session_state["quote_df"] = updated_working_df.copy()
            st.session_state["recognized_df"] = updated_working_df.copy()
            st.session_state["confirmed_quote_df"] = edited_final_df.copy()
            st.session_state["excel_bytes"] = None
        st.caption("这里就是最终导出的报价单明细。单价为 0 时，合计保持 0。")

    st.header("5. 导出区")
    final_export_source_df = st.session_state.get("final_quote_df", pd.DataFrame(columns=FINAL_QUOTE_COLUMNS))
    can_export = not final_export_source_df.empty
    if st.button("生成 Excel 报价单", disabled=not can_export):
        final_export_source_df = recalculate_final_quote_df(st.session_state.get("final_quote_df", pd.DataFrame(columns=FINAL_QUOTE_COLUMNS)))
        st.session_state["final_quote_df"] = final_export_source_df.copy()
        export_df = final_quote_export_df(final_export_source_df)
        excel_bytes = export_quote_to_excel(
            export_df,
            activity_name=activity_name,
            activity_time=activity_time,
            client_name=client_name,
            output_path=EXPORT_PATH,
        )
        st.session_state.excel_bytes = excel_bytes
        st.success("Excel 报价单已生成。")

    if st.session_state.get("excel_bytes"):
        st.download_button(
            "下载报价单.xlsx",
            data=st.session_state.excel_bytes,
            file_name="报价单.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    st.header("6. 未识别候选处理区")
    with st.expander(f"未识别候选：发现 {candidate_count} 个可能漏项，点击展开处理", expanded=False):
        if st.session_state.candidate_display_df.empty:
            st.caption("暂无需要处理的未识别候选。")
        else:
            processing_options = ["暂不处理", "加入已有标准项目别名", "新建标准项目", "标记为无需报价"]
            standard_item_options = ["", *sorted(rules.keys())]
            editor_source_df = normalize_feedback_df(st.session_state.feedback_df)
            feedback_df = st.data_editor(
                editor_source_df,
                use_container_width=True,
                hide_index=True,
                num_rows="fixed",
                disabled=["候选词", "可能归类", "上下文摘要"],
                column_config={
                    "是否处理": st.column_config.CheckboxColumn("是否处理", default=False),
                    "处理方式": st.column_config.SelectboxColumn("处理方式", options=processing_options),
                    "选择标准项目": st.column_config.SelectboxColumn("选择标准项目", options=standard_item_options),
                },
                key="candidate_feedback_editor",
            )
            feedback_df = normalize_feedback_df(feedback_df)
            st.session_state.feedback_df = feedback_df

            new_category = "待分类"
            new_standard_item = ""
            new_quote_type = "模糊报价"
            new_default_unit = "项"
            new_default_desc = "待补充说明"
            needs_new_item_fields = bool(feedback_df["处理方式"].eq("新建标准项目").any())

            if needs_new_item_fields:
                with st.expander("新建标准项目字段", expanded=True):
                    new_col_a, new_col_b, new_col_c = st.columns(3)
                    with new_col_a:
                        new_category = st.text_input("项目分类", value=new_category)
                        new_standard_item = st.text_input("标准项目", placeholder="默认使用候选词")
                    with new_col_b:
                        new_quote_type = st.selectbox("报价类型", options=list(QUOTE_TYPES), index=2)
                        new_default_unit = st.text_input("默认单位", value=new_default_unit)
                    with new_col_c:
                        new_default_desc = st.text_area("默认说明", value=new_default_desc, height=100)

            if st.button("保存规则更新"):
                result = apply_feedback_rows(
                    RULES_PATH,
                    IGNORED_TERMS_PATH,
                    feedback_df.to_dict("records"),
                    {
                        "项目分类": new_category,
                        "标准项目": new_standard_item,
                        "报价类型": new_quote_type,
                        "默认单位": new_default_unit,
                        "默认说明": new_default_desc,
                    },
                )
                messages = result.get("messages", [])
                if messages:
                    with st.expander("保存明细", expanded=False):
                        for message in messages:
                            st.write(message)
                    st.success(f"已更新 {result.get('rules_updated', 0)} 条规则，已忽略 {result.get('ignored_count', 0)} 条候选词")
                    if result.get("rules_updated", 0) or result.get("ignored_count", 0):
                        st.success("规则已更新，下次识别将自动生效。")
                else:
                    st.info("没有需要保存的规则更新。")

            if st.button("重新识别当前方案", key="rerun_after_feedback"):
                if plan_text.strip():
                    run_recognition(plan_text)
                    st.session_state.feedback_rerun_notice = "已按最新规则重新识别当前方案。"
                    st.rerun()
                else:
                    st.warning("当前没有可重新识别的方案文本。")

    st.header("7. 高级信息折叠区")
    with st.expander("高级信息：活动类型识别详情", expanded=False):
        if st.session_state.activity_type_df.empty:
            st.caption("暂无活动类型详情。")
        else:
            st.dataframe(st.session_state.activity_type_df, use_container_width=True, hide_index=True)

    with st.expander("高级信息：完整识别数据", expanded=False):
        st.dataframe(st.session_state.recognized_df, use_container_width=True, hide_index=True)

    with st.expander("高级信息：完整未识别候选", expanded=False):
        if st.session_state.candidate_df.empty:
            st.caption("暂无完整候选数据。")
        else:
            st.dataframe(st.session_state.candidate_df, use_container_width=True, hide_index=True)

    with st.expander("高级信息：规则库调试信息", expanded=False):
        st.write(f"规则库路径：{RULES_PATH}")
        st.write(f"忽略词路径：{IGNORED_TERMS_PATH}")
        st.write(f"标准项目数量：{len(rules)}")
        st.write(f"忽略词数量：{len(load_ignored_terms(IGNORED_TERMS_PATH))}")
else:
    st.info("请先输入或上传方案文本，然后点击“识别报价项”。")
