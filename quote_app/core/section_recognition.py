"""Recognize activity sections and assign quote rows to quote sections."""

from __future__ import annotations

import re
from typing import Any

from .quote_constants import *
from .quote_utils import (
    _matches_keywords,
    _merge_text_values,
    _normalize_item_name,
    _safe_first_match_position,
)

def _clean_section_heading(line: str) -> str:
    raw = str(line).strip()
    if _is_date_or_time_section(raw):
        return raw
    cleaned = SECTION_PREFIX_RE.sub("", raw)
    cleaned = SECTION_SPACE_NUMBER_PREFIX_RE.sub("", cleaned)
    return cleaned.strip("：:  ")


def _range_heading(line: str) -> str:
    cleaned = _clean_section_heading(_normalize_section_line(line))
    cleaned = re.sub(r"^延展活动[：:].*$", "延展活动", cleaned)
    cleaned = cleaned.strip(" ：:")
    return cleaned


def _is_activity_range_start(line: str) -> bool:
    heading = _range_heading(line)
    return (
        heading in ACTIVITY_CONTENT_START_TITLES
        or any(heading.startswith(title) for title in ACTIVITY_CONTENT_START_TITLES)
        or bool(SECTION_EXTEND_INLINE_RE.match(_normalize_section_line(line)))
    )


def _is_activity_range_stop(line: str) -> bool:
    heading = _range_heading(line)
    return heading in ACTIVITY_CONTENT_STOP_TITLES or any(heading.startswith(title) for title in ACTIVITY_CONTENT_STOP_TITLES)


def extract_activity_content_ranges(text: str) -> list[tuple[int, int]]:
    ranges: list[tuple[int, int]] = []
    active_start: int | None = None
    cursor = 0

    for raw_line in text.splitlines(keepends=True):
        if _is_activity_range_stop(raw_line):
            if active_start is not None and active_start < cursor:
                ranges.append((active_start, cursor))
                active_start = None
            cursor += len(raw_line)
            continue

        if _is_activity_range_start(raw_line):
            if active_start is not None and active_start < cursor:
                ranges.append((active_start, cursor))
            active_start = cursor

        cursor += len(raw_line)

    if active_start is not None and active_start < len(text):
        ranges.append((active_start, len(text)))

    return ranges


def diagnose_activity_content_ranges(text: str) -> dict[str, Any]:
    ranges = extract_activity_content_ranges(text)
    diagnostics: list[dict[str, Any]] = []
    for index, (start, end) in enumerate(ranges, start=1):
        excerpt = text[start:end]
        lines = [line.strip() for line in excerpt.splitlines() if line.strip()]
        directory_like = any(token in excerpt[:500] for token in ["目录", "CONTENTS", "PART"])
        warnings: list[str] = []
        if len(excerpt) < 500:
            warnings.append("活动内容区域可能切片错误")
        if directory_like:
            warnings.append("疑似误用了目录页")
        diagnostics.append(
            {
                "index": index,
                "start": start,
                "end": end,
                "length": len(excerpt),
                "directory_like": directory_like,
                "warnings": warnings,
                "preview": excerpt[:160].replace("\n", " / "),
            }
        )

    total_length = sum(item["length"] for item in diagnostics)
    return {
        "range_count": len(ranges),
        "ranges": diagnostics,
        "total_length": total_length,
        "has_short_range": bool(ranges) and total_length < 500,
        "directory_like": any(item["directory_like"] for item in diagnostics),
    }


def diagnose_section_candidates(text: str) -> list[dict[str, Any]]:
    candidates: list[dict[str, Any]] = []
    seen: set[tuple[str, str]] = set()
    cursor = 0
    for raw_line in text.splitlines(keepends=True):
        normalized_line = _normalize_section_line(raw_line)
        cleaned = _clean_section_heading(normalized_line)
        if not cleaned:
            cursor += len(raw_line)
            continue

        structured_names, structured_reason = _extract_structured_section_names(normalized_line, True)
        names = structured_names or [cleaned]
        for name in names:
            candidate_type = classify_section_candidate(name, f"{normalized_line} {structured_reason}")
            if candidate_type in {"activity", "sub_activity"}:
                continue
            reason = "容器标题" if candidate_type == "container" else f"候选分类：{candidate_type}"
            key = (name, reason)
            if key in seen:
                continue
            seen.add(key)
            candidates.append(
                {
                    "候选标题": name,
                    "分类": candidate_type,
                    "过滤原因": reason,
                    "位置": cursor,
                    "原始文本": normalized_line,
                }
            )
        cursor += len(raw_line)
    return candidates


def _is_continuous_short_activity_title(name: str) -> bool:
    cleaned = str(name or "").strip()
    if not (3 <= len(cleaned) <= 18):
        return False
    if TIME_LINE_RE.search(cleaned) or re.search(r"\d{1,2}[:：]\d{2}", cleaned) or "月" in cleaned or "日" in cleaned:
        return False
    if any(word in cleaned for word in SHORT_TITLE_BLOCK_WORDS):
        return False
    return any(keyword in cleaned for keyword in CONTINUOUS_SHORT_TITLE_KEYWORDS) or cleaned in ACTIVITY_SECTION_CANDIDATES


def _is_date_or_time_section(text: str) -> bool:
    cleaned = str(text or "").strip()
    return bool(
        TIME_LINE_RE.search(cleaned)
        or re.fullmatch(r"\d{4}", cleaned)
        or re.fullmatch(r"\d{1,2}\s+\d{1,2}\.\d{1,2}", cleaned)
        or re.fullmatch(r"\d{1,2}(?:\.\d{1,2}){1,2}", cleaned)
        or re.fullmatch(r"\d{1,2}\.\d{1,2}\s*[-—~～]\s*\d{1,2}(?:\.\d{1,2})?", cleaned)
        or re.fullmatch(r"\d{1,2}[:：]\d{2}\s*[-—~～]\s*\d{1,2}[:：]\d{2}", cleaned)
    )


def _starts_with_any(text: str, values: set[str] | tuple[str, ...]) -> bool:
    return any(text == value or text.startswith(f"{value}-") or text.startswith(f"{value}－") or text.startswith(f"{value} ") for value in values)


def classify_section_candidate(name: str, context: str) -> str:
    cleaned = str(name or "").strip()
    source = f"{cleaned} {context or ''}"
    filter_source = cleaned
    if not cleaned:
        return "noise"
    if _is_date_or_time_section(cleaned):
        return "noise"
    if re.fullmatch(r"(?:第)?[一二三四五六七八九十\d]+场", cleaned):
        return "process"

    if any(term in filter_source for term in LOCATION_SECTION_TERMS):
        return "location"
    if any(term in filter_source for term in PROMO_SECTION_TERMS):
        return "promo"
    if any(term in filter_source for term in PROCESS_SECTION_TERMS):
        return "process"
    if any(term in filter_source for term in MATERIAL_SECTION_TERMS):
        return "material"
    if cleaned in CONTAINER_SECTION_TITLES or (cleaned in GENERIC_ACTIVITY_TERMS and cleaned not in GENERIC_ACTIVITY_EXCEPTIONS) or cleaned in SECTION_IGNORE_TITLES:
        return "container"

    if cleaned == "环境装置" and any(term in source for term in ["展览", "展区", "展览内容", "中文括号活动/展区标题", "板块标题"]):
        return "activity"

    if "示意图" in filter_source:
        return "build"
    if any(term in filter_source for term in BUILD_SECTION_TERMS):
        if cleaned not in BUILD_ACTIVITY_EXCEPTIONS:
            return "build"
        if cleaned == "环境装置" and not any(term in source for term in ["展览", "展区", "展览内容", "中文括号活动/展区标题"]):
            return "build"

    if cleaned in SUB_SECTION_TITLES or _starts_with_any(cleaned, SUB_SECTION_TITLES):
        return "sub_activity"

    if cleaned in MAIN_SECTION_TITLES:
        return "activity"
    if cleaned in ACTIVITY_SECTION_CANDIDATES and (cleaned not in GENERIC_ACTIVITY_TERMS or cleaned in GENERIC_ACTIVITY_EXCEPTIONS):
        return "activity"
    if any(term in cleaned for term in ACTIVITY_FORM_TERMS):
        return "activity"
    structured_context = (
        "活动内容区域识别" in source
        or "编号活动标题" in source
        or "中文括号活动/展区标题" in source
        or "区域标题" in source
        or "课程/环节标题" in source
        or ("展区标题" in source and "中文括号活动/展区标题" not in source)
    )
    if structured_context and len(cleaned) <= 25:
        return "activity"

    return "noise"


def classify_section_level(name: str, reason: str = "", confidence: str = "") -> str:
    cleaned = str(name or "").strip()
    if not cleaned:
        return "noise"

    if _is_date_or_time_section(cleaned):
        return "noise"
    if reason == "报价项反推活动板块" and cleaned not in PUBLIC_QUOTE_SECTIONS and cleaned != "人员类及其他":
        return "main"
    if cleaned in CONTAINER_SECTION_TITLES or (cleaned in GENERIC_ACTIVITY_TERMS and cleaned not in GENERIC_ACTIVITY_EXCEPTIONS) or cleaned in SECTION_IGNORE_TITLES:
        return "noise"
    candidate_type = classify_section_candidate(cleaned, reason)
    if candidate_type not in {"activity", "sub_activity"}:
        return "noise"
    if cleaned in MATERIAL_SECTION_TERMS:
        return "noise"
    if cleaned in PROMOTION_ACTION_TITLES or cleaned.startswith(PROMOTION_SENTENCE_PREFIXES):
        return "noise"
    if len(cleaned) > 18 and cleaned not in MAIN_SECTION_TITLES and cleaned not in ACTIVITY_SECTION_CANDIDATES:
        return "noise"
    is_structured_space_title = any(term in reason for term in ["区域标题", "展区标题"])
    noise_words = [word for word in SECTION_NOISE_SENTENCE_WORDS if word != "体验"]
    experience_as_activity_title = (
        "体验" in cleaned
        and len(cleaned) <= 18
        and candidate_type == "activity"
        and not any(word in cleaned for word in noise_words)
        and not cleaned.startswith(SECTION_STATEMENT_PREFIXES)
    )
    if (
        any(word in cleaned for word in noise_words)
        or ("体验" in cleaned and not experience_as_activity_title)
    ) and cleaned not in MAIN_SECTION_TITLES and not is_structured_space_title:
        return "noise"

    if candidate_type == "sub_activity" or _starts_with_any(cleaned, SUB_SECTION_TITLES):
        return "sub"
    if confidence == "candidate" or cleaned in SECTION_CHILD_MODULES:
        return "sub"

    if cleaned in MAIN_SECTION_TITLES:
        return "main"
    if reason in {"板块标题", "延展活动标题"}:
        return "main"
    if cleaned in ACTIVITY_SECTION_CANDIDATES and (cleaned not in GENERIC_ACTIVITY_TERMS or cleaned in GENERIC_ACTIVITY_EXCEPTIONS):
        return "main"
    if confidence == "strong":
        return "main"
    return "sub"


def _is_forbidden_section_name(name: str) -> bool:
    cleaned = str(name or "").strip()
    if not cleaned:
        return True
    if cleaned in SECTION_FORBIDDEN_TITLES:
        return True
    if len(cleaned) > 18 and cleaned not in ACTIVITY_SECTION_CANDIDATES:
        return True
    for word in SECTION_FORBIDDEN_WORDS:
        if word not in cleaned:
            continue
        if word == "品牌" and any(term in cleaned for term in ACTIVITY_FORM_TERMS):
            continue
        return True
    if cleaned.startswith(SECTION_STATEMENT_PREFIXES):
        return True
    return False


def _clean_explicit_section_name(raw_name: str) -> str | None:
    cleaned = str(raw_name or "").strip()
    cleaned = re.sub(r"^[（(][一二三四五六七八九十]+[)）]\s*", "", cleaned)
    cleaned = re.sub(r"^板块[一二三四五六七八九十]+[：:]\s*", "", cleaned)
    cleaned = re.sub(r"^[一二三四五六七八九十]+、\s*延展活动[：:]\s*", "", cleaned)
    cleaned = SECTION_PREFIX_RE.sub("", cleaned)
    named_match = SECTION_NAMED_CONTAINER_PREFIX_RE.match(cleaned)
    if named_match:
        cleaned = named_match.group(1)
    elif re.search(r"[：:]", cleaned):
        before, after = re.split(r"[：:]", cleaned, maxsplit=1)
        if 2 <= len(before.strip()) <= 18 and len(after.strip()) >= 4:
            cleaned = before
    cleaned = cleaned.strip(" ：:")
    cleaned = re.sub(r"\s+", " ", cleaned)

    if not cleaned:
        return None
    if _is_date_or_time_section(cleaned) or re.fullmatch(r"[\d.\-—~～]+", cleaned):
        return None
    if any(word in cleaned for word in SECTION_EXPLICIT_NOISE_WORDS):
        return None
    if _is_forbidden_section_name(cleaned):
        return None
    if len(cleaned) < 2 or len(cleaned) > 30:
        return None
    return cleaned


def _has_explanatory_colon_tail(line: str, name: str) -> bool:
    if not re.search(r"[：:]", line):
        return False
    before, after = re.split(r"[：:]", line, maxsplit=1)
    return name in before and len(after.strip()) >= 4


def _extract_structured_section_names(line: str, container_active: bool) -> tuple[list[str], str]:
    normalized = _normalize_section_line(line)
    if _is_date_or_time_section(normalized):
        return [], ""
    for reason, pattern in STRUCTURE_PATTERNS:
        match = pattern.match(normalized)
        if match:
            name = _clean_explicit_section_name(match.group(1))
            if not name:
                return [], ""
            if reason in {"板块标题", "延展活动标题", "展区标题", "区域标题", "课程/环节标题"}:
                return [name], reason
            if reason == "中文括号活动/展区标题":
                without_prefix = SECTION_PREFIX_RE.sub("", normalized).strip()
                if SECTION_NAMED_CONTAINER_PREFIX_RE.match(without_prefix):
                    return [name], reason
                candidate_type = classify_section_candidate(name, reason)
                if candidate_type in {"activity", "sub_activity"}:
                    return [name], reason
                if container_active and (_contains_activity_keyword(name) or any(keyword in name for keyword in ACTIVITY_SECTION_KEYWORDS)):
                    return [name], reason
                return [], ""
            if (
                reason == "数字编号活动标题"
                and _has_explanatory_colon_tail(normalized, name)
                and name not in MAIN_SECTION_TITLES
                and name not in ACTIVITY_SECTION_CANDIDATES
                and name not in SECTION_CHILD_MODULES
                and not _contains_activity_keyword(name)
                and not any(term in name for term in ACTIVITY_FORM_TERMS)
            ):
                return [], ""
            if (
                container_active
                or name in SECTION_CHILD_MODULES
                or _is_continuous_short_activity_title(name)
                or _contains_activity_keyword(name)
                or any(keyword in name for keyword in ACTIVITY_SECTION_KEYWORDS)
            ):
                return [name], reason
            candidate_type = classify_section_candidate(name, reason)
            if candidate_type in {"activity", "sub_activity"}:
                return [name], reason
            return [], ""

    match = SECTION_CN_NUMBER_RE.match(normalized)
    if match:
        name = _clean_explicit_section_name(match.group(1))
        if name and (
            container_active
            or name in SECTION_CHILD_MODULES
            or _is_continuous_short_activity_title(name)
            or _contains_activity_keyword(name)
            or any(keyword in name for keyword in ACTIVITY_SECTION_KEYWORDS)
        ):
            return [name], "数字编号活动标题"

    return [], ""


def _best_section_name(text: str) -> str:
    cleaned = str(text).strip()
    if cleaned in SHORT_SECTION_TITLES or cleaned in ACTIVITY_SECTION_CANDIDATES:
        return cleaned
    if len(cleaned) <= 30 and any(keyword in cleaned for keyword in ACTIVITY_SECTION_KEYWORDS):
        return cleaned
    for candidate in sorted(ACTIVITY_SECTION_CANDIDATES, key=len, reverse=True):
        if candidate in cleaned:
            return cleaned
    return cleaned


def _strip_common_ppt_prefix(line: str) -> str:
    return re.sub(r"^(?:PART|CONTENTS)\s*\d*\s*", "", str(line).strip(), flags=re.IGNORECASE).strip()


def _normalize_section_line(line: str) -> str:
    return re.sub(r"\s+", " ", _strip_common_ppt_prefix(str(line))).strip()


def _canonicalize_section_text(text: str) -> str:
    cleaned = str(text or "")
    cleaned = re.sub(r"[\\|]", "/", cleaned)
    cleaned = re.sub(r"[：:；;，,。.!！?？【】\[\]（）()]+", " ", cleaned)
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    return cleaned


def split_compound_section_name(raw_name: str) -> list[str]:
    text = _canonicalize_section_text(raw_name)
    if not text:
        return []

    normalized_hits = _scan_specific_section_names(text)
    if normalized_hits:
        return normalized_hits

    return [text]


def _trim_section_edges(text: str) -> str:
    cleaned = text.strip()
    for word in SECTION_TRAILING_NOISE_WORDS:
        cleaned = re.sub(rf"^(?:{re.escape(word)})+", "", cleaned)
        cleaned = re.sub(rf"(?:{re.escape(word)})+$", "", cleaned)
    cleaned = cleaned.strip(" /")
    cleaned = re.sub(r"\s+", " ", cleaned)
    return cleaned.strip()


def _find_section_alias_rule(name: str) -> dict[str, Any] | None:
    source_name = _canonicalize_section_text(name)
    for rule in SECTION_ALIAS_MAP:
        include = rule["include"]
        exclude = rule.get("exclude", [])
        if any(keyword in source_name for keyword in include) and not any(keyword in source_name for keyword in exclude):
            return rule
    return None


def _map_section_alias(name: str) -> str | None:
    rule = _find_section_alias_rule(name)
    if rule:
        return str(rule["name"])
    return None


def _scan_specific_section_names(text: str) -> list[str]:
    canonical_text = _canonicalize_section_text(text)
    if not canonical_text:
        return []

    hits: list[tuple[int, int, str]] = []
    for rule in SECTION_ALIAS_MAP:
        exclude = rule.get("exclude", [])
        if any(keyword in canonical_text for keyword in exclude):
            continue
        source_terms = rule.get("source_terms", rule.get("include", []))
        matched_positions = [canonical_text.find(term) for term in source_terms if term in canonical_text]
        if matched_positions:
            hits.append((min(matched_positions), -len(str(rule["name"])), str(rule["name"])))

    specific_candidates = [
        candidate
        for candidate in ACTIVITY_SECTION_CANDIDATES
        if (candidate not in GENERIC_ACTIVITY_TERMS or candidate in GENERIC_ACTIVITY_EXCEPTIONS) and candidate not in CONTAINER_SECTION_TITLES and len(candidate) >= 4
    ]
    for candidate in specific_candidates:
        position = canonical_text.find(candidate)
        if position != -1:
            hits.append((position, -len(candidate), candidate))

    ordered_names: list[str] = []
    seen_names: set[str] = set()
    for _, _, name in sorted(hits):
        if any(name != existing and name in existing for existing in ordered_names):
            continue
        if name not in seen_names:
            ordered_names.append(name)
            seen_names.add(name)
    return ordered_names


def validate_section_source(section: dict[str, Any], full_text: str) -> bool:
    normalized_name = str(section.get("normalized_name") or section.get("name") or "").strip()
    raw_name = str(section.get("raw_name", "")).strip()
    source_text = str(section.get("source_text") or section.get("source_context") or "").strip()

    if normalized_name in PUBLIC_QUOTE_SECTIONS or normalized_name == "未归属板块":
        return True

    if not source_text:
        return False

    if normalized_name and normalized_name in full_text:
        return True
    if raw_name and raw_name in full_text:
        return True

    alias_rule = _find_section_alias_rule(raw_name) or _find_section_alias_rule(source_text) or _find_section_alias_rule(normalized_name)
    if alias_rule:
        source_terms = alias_rule.get("source_terms", alias_rule.get("include", []))
        return any(term in source_text for term in source_terms)

    return False


def normalize_section_name(raw_name: str) -> str | None:
    raw_text = str(raw_name or "").strip()
    named_match = SECTION_NAMED_CONTAINER_PREFIX_RE.match(raw_text)
    if named_match:
        raw_text = named_match.group(1)
    cleaned = _canonicalize_section_text(raw_text)
    named_space_match = SECTION_NAMED_CONTAINER_PREFIX_SPACE_RE.match(cleaned)
    if named_space_match:
        cleaned = named_space_match.group(1).strip()
    cleaned = SECTION_SPACE_NUMBER_PREFIX_RE.sub("", cleaned)
    cleaned = _trim_section_edges(cleaned)
    if not cleaned:
        return None

    if "手艺启元礼" in raw_name and "启动仪式" in raw_name:
        return "手艺启元礼.启动仪式"

    if cleaned in MAIN_SECTION_TITLES:
        return cleaned

    for sub_title in sorted(SUB_SECTION_TITLES, key=len, reverse=True):
        if cleaned == sub_title or cleaned.startswith(f"{sub_title} ") or cleaned.startswith(f"{sub_title}-"):
            return sub_title

    mapped_name = _map_section_alias(raw_name) or _map_section_alias(cleaned)
    if mapped_name:
        return mapped_name

    cleaned = re.sub(r"(走秀)\s+\1$", r"\1", cleaned)
    cleaned = re.sub(r"(活动)\s+\1$", r"\1", cleaned)
    cleaned = re.sub(r"(体验)\s+\1$", r"\1", cleaned)
    cleaned = re.sub(r"^(走秀|活动|体验)$", "", cleaned)
    cleaned = _trim_section_edges(cleaned)
    if not cleaned:
        return None

    if TIME_LINE_RE.search(cleaned) and not _contains_activity_keyword(cleaned):
        return None

    if len(cleaned) > 16 and any(word in cleaned for word in SECTION_NOISE_SENTENCE_WORDS):
        return None

    if re.fullmatch(r"(.+)\s+\1", cleaned):
        return None

    if len(cleaned) < 3:
        return None

    if len(cleaned) > 18 and cleaned not in ACTIVITY_SECTION_CANDIDATES:
        return None

    if (cleaned in GENERIC_ACTIVITY_TERMS and cleaned not in GENERIC_ACTIVITY_EXCEPTIONS) or cleaned in SECTION_IGNORE_TITLES:
        return None

    if _is_forbidden_section_name(cleaned):
        return None

    return cleaned


def _contains_activity_keyword(text: str) -> bool:
    return any(keyword in text for keyword in STRONG_ACTIVITY_KEYWORDS) or any(name == text for name in ACTIVITY_SECTION_CANDIDATES)


def _is_noise_line(text: str) -> tuple[bool, str]:
    if not text:
        return True, "空行"
    if _is_date_or_time_section(text):
        return True, "日期/时间"
    if _is_forbidden_section_name(text):
        return True, "非活动板块章节/说明"
    if text in SECTION_IGNORE_TITLES or text in NOISE_INFO_TERMS:
        return True, "基础信息/说明标题"
    if text in GENERIC_ACTIVITY_TERMS and text not in GENERIC_ACTIVITY_EXCEPTIONS:
        return True, "泛词容器标题"
    candidate_type = classify_section_candidate(text, text)
    if candidate_type != "activity":
        return True, f"候选分类：{candidate_type}"
    if text.endswith("活动方案"):
        return True, "方案标题"
    if "、" in text or "，" in text or "," in text:
        generic_hits = [term for term in ["市集", "美食", "巡游", "展览", "音乐", "灯展", "游园"] if term in text]
        if len(generic_hits) >= 2:
            return True, "框架词列举"
    if len(text) > 18 and text not in MAIN_SECTION_TITLES and text not in ACTIVITY_SECTION_CANDIDATES:
        return True, "长度过长，像说明句"
    if re.search(r"[，。；;,.!?！？]", text):
        return True, "包含句读符号，像正文句"
    if any(term in text for term in NOISE_SENTENCE_TERMS) and len(text) > 8:
        return True, "说明句/残句"
    if any(term in text for term in ["布置", "氛围布置", "示意图", "场地", "点位", "平面图", "动线", "区域"]):
        return True, "布置/点位类内容"
    if text in {"项互动体验", "体验动", "场地", "利面", "舞台 停"}:
        return True, "断行残句"
    return False, ""


def _is_container_heading(text: str) -> bool:
    clean = _clean_section_heading(text)
    return (
        clean in CONTAINER_SECTION_TITLES
        or any(clean.startswith(title) for title in CONTAINER_SECTION_TITLES)
        or any(hint in clean for hint in PROMOTION_CONTAINER_HINTS)
    )


def _is_numbered_heading(raw_line: str) -> bool:
    return bool(SECTION_PREFIX_RE.match(raw_line) or SECTION_SPACE_NUMBER_PREFIX_RE.match(raw_line))


def _classify_section_line(
    raw_line: str,
    container_active: bool,
) -> tuple[str, str]:
    normalized = _normalize_section_line(raw_line)
    cleaned = _clean_section_heading(normalized)
    looks_like_heading = _looks_like_section_heading(normalized)
    inline_names = _scan_specific_section_names(normalized) if looks_like_heading else []

    if _is_container_heading(normalized) or _is_container_heading(cleaned):
        return "container", "容器标题"

    if cleaned in GENERIC_ACTIVITY_TERMS and cleaned not in GENERIC_ACTIVITY_EXCEPTIONS:
        return "noise", "通用框架词"

    if _starts_with_any(cleaned, SUB_SECTION_TITLES):
        return "candidate", "子活动/流程环节"

    alias_hit = _map_section_alias(cleaned)
    if alias_hit and looks_like_heading:
        return "strong", f"命中标准板块映射：{alias_hit}"

    if len(inline_names) >= 2:
        return "strong", "并列活动名称清单"

    if container_active and inline_names:
        return "strong", "容器标题下的活动内容清单"

    if cleaned in ACTIVITY_SECTION_CANDIDATES:
        return "strong", "白名单活动名称"

    is_noise, noise_reason = _is_noise_line(cleaned)
    if is_noise:
        return "noise", noise_reason

    numbered = _is_numbered_heading(normalized)
    short_title = 4 <= len(cleaned) <= 18
    activity_like = _contains_activity_keyword(cleaned)
    continuous_short_title = container_active and _is_continuous_short_activity_title(cleaned)

    if numbered and short_title and activity_like:
        return "strong", "明确编号标题"

    if container_active and numbered and short_title and (activity_like or continuous_short_title):
        return "strong", "容器标题下的编号活动"

    if short_title and (activity_like or continuous_short_title) and cleaned == normalized:
        return "strong", "独立短标题且含活动关键词"

    if len(cleaned) <= 18 and (activity_like or continuous_short_title):
        return "candidate", "疑似活动标题"

    return "noise", "不满足活动板块特征"


def _looks_like_section_heading(line: str) -> bool:
    raw_line = _strip_common_ppt_prefix(line)
    cleaned = _clean_section_heading(raw_line)
    if not cleaned or len(cleaned) > 30:
        return False
    if cleaned in SECTION_IGNORE_TITLES:
        return False
    if cleaned.endswith("活动方案"):
        return False
    if cleaned.startswith(("核心活动：", "配套活动：")):
        return False
    has_number_prefix = bool(SECTION_PREFIX_RE.match(raw_line))
    if re.search(r"[，。；;,.!?！？、]", cleaned) and not has_number_prefix:
        return False
    if any(term in cleaned for term in [*NOISE_SENTENCE_TERMS, "设置", "配置"]) and cleaned not in ACTIVITY_SECTION_CANDIDATES and not has_number_prefix:
        return False
    if cleaned in SHORT_SECTION_TITLES or cleaned in ACTIVITY_SECTION_CANDIDATES:
        return True

    if any(candidate in cleaned for candidate in ACTIVITY_SECTION_CANDIDATES):
        return True

    has_keyword = any(keyword in cleaned for keyword in ACTIVITY_SECTION_KEYWORDS)
    is_short_standalone = cleaned == raw_line and len(cleaned) <= 10 and not re.search(r"[，。；;,.]", cleaned)
    return has_keyword and (has_number_prefix or is_short_standalone)


def _default_parent_for_sub_activity(name: str, current_parent: str = "") -> str:
    cleaned = str(name or "").strip()
    if current_parent:
        return current_parent
    if cleaned in {"签到仪式", "开场舞", "启动仪式", "主持推荐", "领导致辞", "授牌仪式", "合影留念"}:
        return "开幕式"
    if cleaned in {"套蚝赢趣", "钓壳寻鲜", "贝壳绘梦", "曲口寻踪"}:
        return "蚝趣游乐记"
    if cleaned in {"赶猪赛跑", "巧夹槟榔", "扁担挑椰子", "拉乌龟", "欢乐背背跑"}:
        return "民族趣味运动会"
    if cleaned in SECTION_CHILD_MODULES:
        return "手艺·海岸集"
    if cleaned in {"讲师介绍", "课程导入", "课间休息", "茶歇交流", "自由交流", "问答交流"}:
        return "专题授课"
    return ""


def _make_activity_section(
    *,
    name: str,
    raw_name: str,
    start: int,
    end: int,
    confidence: str,
    reason: str,
    source_text: str,
    channel: str,
    order: int,
    parent: str = "",
) -> dict[str, Any] | None:
    normalized_name = normalize_section_name(name) or normalize_section_name(raw_name) or str(name or "").strip()
    if not normalized_name:
        return None
    candidate_type = classify_section_candidate(normalized_name, f"{raw_name} {source_text} {reason} {channel}")
    if candidate_type not in {"activity", "sub_activity"}:
        return None
    section_level = classify_section_level(normalized_name, reason, confidence)
    if section_level == "noise":
        return None
    if section_level == "sub" and candidate_type == "activity":
        candidate_type = "sub_activity"
    if section_level == "main":
        parent = ""
    return {
        "name": normalized_name,
        "raw_name": raw_name,
        "normalized_name": normalized_name,
        "level": section_level,
        "type": candidate_type,
        "parent": parent,
        "start": start,
        "end": end,
        "order": order,
        "confidence": confidence,
        "section_confidence": confidence,
        "section_level": section_level,
        "selected": section_level == "main",
        "source_context": source_text[:80],
        "source_text": source_text,
        "reason": reason,
        "channel": channel,
        "candidate_type": candidate_type,
        "filter_reason": "",
    }


def _merge_activity_sections(sections: list[dict[str, Any]], text_length: int) -> list[dict[str, Any]]:
    merged: dict[str, dict[str, Any]] = {}
    for section in sorted(sections, key=lambda item: (int(item.get("start", 0)), str(item.get("name", "")))):
        name = str(section.get("normalized_name") or section.get("name") or "").strip()
        if not name:
            continue
        if name not in merged:
            merged[name] = dict(section)
            continue

        target = merged[name]
        target["start"] = min(int(target.get("start", 0)), int(section.get("start", 0)))
        target["end"] = max(int(target.get("end", 0)), int(section.get("end", 0)))
        target["selected"] = bool(target.get("selected")) or bool(section.get("selected"))
        if target.get("section_level") != "main" and section.get("section_level") == "main":
            target["section_level"] = "main"
            target["selected"] = True
        if target.get("section_confidence") != "strong" and section.get("section_confidence") == "strong":
            target["section_confidence"] = "strong"
        if not target.get("parent") and section.get("parent"):
            target["parent"] = section.get("parent")
        if section.get("candidate_type") == "sub_activity" and target.get("section_level") != "main":
            target["candidate_type"] = "sub_activity"
            target["type"] = "sub_activity"
        target["channel"] = _merge_text_values(target.get("channel", ""), section.get("channel", ""))
        target["reason"] = _merge_text_values(target.get("reason", ""), section.get("reason", ""))

    ordered = sorted(merged.values(), key=lambda item: (int(item.get("start", 0)), str(item.get("name", ""))))
    has_specific_main = any(
        section.get("section_level") == "main"
        and "区域标题" not in str(section.get("reason", ""))
        for section in ordered
    )
    if has_specific_main:
        ordered = [
            section
            for section in ordered
            if section.get("section_level") != "main" or "区域标题" not in str(section.get("reason", ""))
        ]
    for index, section in enumerate(ordered):
        next_start = ordered[index + 1]["start"] if index + 1 < len(ordered) else text_length
        section["end"] = next_start
        section["order"] = index + 1
    return ordered


def _extract_sections_from_ranges(text: str, content_ranges: list[tuple[int, int]], channel: str) -> list[dict[str, Any]]:
    sections: list[dict[str, Any]] = []
    for range_start, range_end in content_ranges:
        cursor = range_start
        container_until = range_start - 1
        current_parent = ""

        for raw_line in text[range_start:range_end].splitlines(keepends=True):
            normalized_line = _normalize_section_line(raw_line)
            cleaned = _clean_section_heading(normalized_line)
            if not cleaned:
                cursor += len(raw_line)
                continue

            container_active = cursor <= container_until
            structured_names, structured_reason = _extract_structured_section_names(normalized_line, container_active)
            if structured_names:
                for split_index, normalized_name in enumerate(structured_names):
                    colon_tail_under_parent = (
                        structured_reason == "数字编号活动标题"
                        and bool(current_parent)
                        and _has_explanatory_colon_tail(normalized_line, normalized_name)
                    )
                    confidence = "candidate" if normalized_name in SECTION_CHILD_MODULES or colon_tail_under_parent else "strong"
                    section = _make_activity_section(
                        name=normalized_name,
                        raw_name=normalized_name,
                        start=cursor + split_index,
                        end=range_end,
                        confidence=confidence,
                        reason=structured_reason,
                        source_text=normalized_line,
                        channel=channel,
                        order=len(sections) + 1,
                        parent=_default_parent_for_sub_activity(normalized_name, current_parent),
                    )
                    if section:
                        sections.append(section)
                        if section.get("section_level") == "main":
                            current_parent = str(section.get("name", ""))
                cursor += len(raw_line)
                continue

            confidence, reason = _classify_section_line(normalized_line, container_active)

            if confidence == "container":
                container_until = min(cursor + 3000, range_end)
                cursor += len(raw_line)
                continue

            if confidence == "noise":
                cursor += len(raw_line)
                continue

            split_names = split_compound_section_name(cleaned)
            if not split_names:
                split_names = [cleaned]

            for split_index, raw_name in enumerate(split_names):
                normalized_name = normalize_section_name(raw_name)
                if not normalized_name:
                    continue
                split_alias = _map_section_alias(raw_name)
                final_confidence = "strong" if split_alias or normalized_name in ACTIVITY_SECTION_CANDIDATES else confidence
                if normalized_name in SECTION_CHILD_MODULES:
                    final_confidence = "candidate"
                section = _make_activity_section(
                    name=normalized_name,
                    raw_name=raw_name,
                    start=cursor + split_index,
                    end=range_end,
                    confidence=final_confidence,
                    reason=reason if final_confidence == confidence else f"命中标准板块映射：{split_alias or normalized_name}",
                    source_text=normalized_line,
                    channel=channel,
                    order=len(sections) + 1,
                    parent=_default_parent_for_sub_activity(normalized_name, current_parent),
                )
                if not section:
                    continue
                if not validate_section_source(section, text):
                    continue
                sections.append(section)
                if section.get("section_level") == "main":
                    current_parent = str(section.get("name", ""))
            cursor += len(raw_line)

    return sections


def _find_title_like_whitelist_position(text: str, name: str, search_terms: list[str]) -> tuple[int, str] | None:
    cursor = 0
    for raw_line in text.splitlines(keepends=True):
        normalized_line = _normalize_section_line(raw_line)
        if not normalized_line or not any(term and term in normalized_line for term in search_terms):
            cursor += len(raw_line)
            continue
        if not _looks_like_section_heading(normalized_line):
            cursor += len(raw_line)
            continue
        hit_names = _scan_specific_section_names(normalized_line)
        if name in hit_names:
            positions = [normalized_line.find(term) for term in search_terms if term and term in normalized_line]
            return cursor + min(position for position in positions if position >= 0), normalized_line
        cursor += len(raw_line)
    return None


def _extract_whitelist_sections(text: str) -> list[dict[str, Any]]:
    sections: list[dict[str, Any]] = []
    for name in _scan_specific_section_names(text):
        search_terms = []
        alias_rule = _find_section_alias_rule(name)
        if alias_rule:
            search_terms.extend(alias_rule.get("source_terms", alias_rule.get("include", [])))
        search_terms.append(name)
        occurrence = _find_title_like_whitelist_position(text, name, search_terms)
        if not occurrence:
            continue
        start, source_text = occurrence
        section = _make_activity_section(
            name=name,
            raw_name=name,
            start=start,
            end=len(text),
            confidence="strong",
            reason="白名单强活动名识别",
            source_text=source_text,
            channel="白名单强活动名识别",
            order=len(sections) + 1,
        )
        if section:
            sections.append(section)
    return sections


def _extract_numbered_sections(text: str) -> list[dict[str, Any]]:
    sections: list[dict[str, Any]] = []
    cursor = 0
    content_ranges = extract_activity_content_ranges(text)
    for raw_line in text.splitlines(keepends=True):
        normalized_line = _normalize_section_line(raw_line)
        container_active = any(start <= cursor < end for start, end in content_ranges)
        structured_names, reason = _extract_structured_section_names(normalized_line, container_active)
        for split_index, name in enumerate(structured_names):
            if _has_explanatory_colon_tail(normalized_line, name) and name not in MAIN_SECTION_TITLES and name not in ACTIVITY_SECTION_CANDIDATES:
                continue
            section = _make_activity_section(
                name=name,
                raw_name=name,
                start=cursor + split_index,
                end=len(text),
                confidence="strong" if name not in SECTION_CHILD_MODULES else "candidate",
                reason=f"编号活动标题识别：{reason}",
                source_text=normalized_line,
                channel="编号活动标题识别",
                order=len(sections) + 1,
                parent=_default_parent_for_sub_activity(name),
            )
            if section:
                sections.append(section)
        cursor += len(raw_line)
    return sections


def _extract_continuous_short_title_sections(text: str) -> list[dict[str, Any]]:
    sections: list[dict[str, Any]] = []
    cursor = 0
    container_until = -1
    pending: list[tuple[str, int, str]] = []

    def flush_pending() -> None:
        if len(pending) < 3:
            pending.clear()
            return
        for name, start, source_text in pending:
            section = _make_activity_section(
                name=name,
                raw_name=name,
                start=start,
                end=len(text),
                confidence="strong",
                reason="连续短标题列表识别",
                source_text=source_text,
                channel="连续短标题列表识别",
                order=len(sections) + 1,
            )
            if section:
                sections.append(section)
        pending.clear()

    for raw_line in text.splitlines(keepends=True):
        normalized_line = _normalize_section_line(raw_line)
        cleaned = _clean_section_heading(normalized_line)
        if _is_container_heading(cleaned):
            flush_pending()
            container_until = min(cursor + 3000, len(text))
            cursor += len(raw_line)
            continue

        structured_names, _ = _extract_structured_section_names(normalized_line, cursor <= container_until)
        if structured_names:
            flush_pending()
            cursor += len(raw_line)
            continue

        if cursor <= container_until and _is_continuous_short_activity_title(cleaned):
            pending.append((cleaned, cursor, normalized_line))
        else:
            flush_pending()
        cursor += len(raw_line)

    flush_pending()
    return sections


def infer_activity_sections_from_quote_rows(rows: list[dict[str, Any]], text: str = "") -> list[dict[str, Any]]:
    sections: list[dict[str, Any]] = []
    seen: set[str] = set()
    for row in rows:
        section_name = infer_section_from_item(str(row.get("标准项目", "")))
        if not section_name or section_name in PUBLIC_QUOTE_SECTIONS or section_name in seen:
            continue
        start = _safe_first_match_position(row.get("匹配位置"))
        if start is None:
            start = 0
        section = _make_activity_section(
            name=section_name,
            raw_name=section_name,
            start=start,
            end=len(text),
            confidence="strong",
            reason="报价项反推活动板块",
            source_text=str(row.get("evidence_text") or row.get("原始命中词") or section_name),
            channel="报价项反推活动板块",
            order=len(sections) + 1,
        )
        if section:
            sections.append(section)
            seen.add(section_name)
    return sections


def extract_activity_sections(text: str, quote_rows: list[dict[str, Any]] | None = None) -> list[dict[str, Any]]:
    content_ranges = extract_activity_content_ranges(text)
    sections: list[dict[str, Any]] = []
    if content_ranges:
        sections.extend(_extract_sections_from_ranges(text, content_ranges, "活动内容区域识别"))

    sections.extend(_extract_sections_from_ranges(text, [(0, len(text))], "全文强标题识别"))
    sections.extend(_extract_whitelist_sections(text))
    sections.extend(_extract_numbered_sections(text))
    sections.extend(_extract_continuous_short_title_sections(text))

    merged = _merge_activity_sections(sections, len(text))
    if not any(section.get("section_level") == "main" for section in merged) and quote_rows:
        merged = _merge_activity_sections([*merged, *infer_activity_sections_from_quote_rows(quote_rows, text)], len(text))
    return merged


def is_public_promotion_item(item_name: str) -> bool:
    return _matches_keywords(_normalize_item_name(item_name), PROMOTION_KEYWORDS)


def is_public_beauty_build_item(item_name: str) -> bool:
    return _matches_keywords(_normalize_item_name(item_name), PUBLIC_BEAUTY_BUILD_KEYWORDS)


def is_public_build_item(item_name: str) -> bool:
    return _matches_keywords(_normalize_item_name(item_name), PUBLIC_BUILD_KEYWORDS)


def is_people_or_other_item(item_name: str) -> bool:
    return _matches_keywords(_normalize_item_name(item_name), PEOPLE_OR_OTHER_KEYWORDS)


def is_public_merge_item(item_name: str) -> bool:
    return _normalize_item_name(item_name) in PUBLIC_MERGE_ITEMS


def find_section_for_match(match_start: int | None, activity_sections: list[dict[str, Any]]) -> dict[str, Any] | None:
    if match_start is None or match_start < 0:
        return None
    for section in activity_sections:
        if section["start"] <= int(match_start) < section["end"]:
            return section
    return None


def prepare_confirmed_activity_sections(activity_sections: list[dict[str, Any]]) -> list[dict[str, Any]]:
    confirmed = [
        dict(section)
        for section in activity_sections
        if bool(section.get("selected")) and str(section.get("section_level", "main")) in {"main", "sub"}
    ]
    confirmed.sort(key=lambda item: int(item.get("start", 0)))
    for index, section in enumerate(confirmed):
        next_start = confirmed[index + 1]["start"] if index + 1 < len(confirmed) else section.get("end", section.get("start", 0))
        section["end"] = next_start
        section["order"] = index + 1
    return confirmed


def infer_section_from_item(item_name: str) -> str | None:
    normalized_name = _normalize_item_name(item_name)
    if normalized_name in ITEM_SECTION_INFERENCE:
        return ITEM_SECTION_INFERENCE[normalized_name]

    if _matches_keywords(normalized_name, ["阅读", "图书", "书签", "二维码"]):
        return "阅读活动"
    if _matches_keywords(normalized_name, ["市集", "摊位"]):
        return "市集活动"
    if _matches_keywords(normalized_name, ["拓印", "手作", "非遗"]):
        return "非遗体验"
    if _matches_keywords(normalized_name, ["游戏", "互动", "印章", "通行证", "主题卡"]):
        return "互动体验"
    if _matches_keywords(normalized_name, ["演出", "走秀", "巡游", "音乐会"]):
        return "演艺活动"
    return None


def assign_quote_section(
    item: dict[str, Any],
    match_start: int | None,
    context: str,
    activity_sections: list[dict[str, Any]],
) -> str:
    item_name = _normalize_item_name(item.get("标准项目", item.get("项目", "")))
    matched_section = find_section_for_match(match_start, activity_sections)

    if (
        str(item.get("evidence_type", "")) == "module_completion"
        and matched_section
        and item_name in MODULE_SECTION_PRIORITY_ITEMS
    ):
        matched_name = str(matched_section["name"])
        if matched_name not in GENERIC_SECTION_TITLES:
            return matched_name

    if is_public_promotion_item(item_name):
        return "活动宣传"
    if is_public_beauty_build_item(item_name):
        return "美陈搭建类"
    if is_public_build_item(item_name):
        return "其他搭建类"

    if is_people_or_other_item(item_name):
        if not matched_section and not activity_sections:
            return "未归属板块"
        return "人员类及其他"

    if matched_section:
        matched_name = str(matched_section["name"])
        if matched_name not in GENERIC_SECTION_TITLES:
            return matched_name

    inferred_section = infer_section_from_item(item_name)
    if inferred_section:
        confirmed_section_names = {
            str(section.get("name", "")).strip()
            for section in activity_sections
            if str(section.get("section_level", "main")) in {"main", "sub"}
        }
        if inferred_section in confirmed_section_names:
            return inferred_section
        if not confirmed_section_names:
            return inferred_section
        return "未归属板块"

    if matched_section:
        return str(matched_section["name"])

    return "未归属板块"
