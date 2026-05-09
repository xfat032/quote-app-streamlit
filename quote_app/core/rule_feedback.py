"""Persist manual feedback from unrecognized candidates into local rule files."""

from __future__ import annotations

import json
import shutil
from datetime import datetime
from pathlib import Path
from typing import Any


def _now_stamp() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def ensure_ignored_terms(path: str | Path) -> Path:
    path = Path(path)
    if not path.exists():
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text("[]\n", encoding="utf-8")
    return path


def load_ignored_terms(path: str | Path) -> list[str]:
    path = ensure_ignored_terms(path)
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return []
    if not isinstance(data, list):
        return []
    return [str(item) for item in data if str(item).strip()]


def save_ignored_term(path: str | Path, term: str) -> bool:
    term = term.strip()
    if not term:
        return False

    path = ensure_ignored_terms(path)
    terms = load_ignored_terms(path)
    if term in terms:
        return False
    terms.append(term)
    path.write_text(json.dumps(terms, ensure_ascii=False, indent=2), encoding="utf-8")
    return True


def backup_rules_config(rules_path: str | Path) -> Path:
    rules_path = Path(rules_path)
    backup_path = rules_path.with_name(f"rules_config.backup_{_now_stamp()}.json")
    shutil.copy2(rules_path, backup_path)
    return backup_path


def _load_rules_file(rules_path: str | Path) -> dict[str, dict[str, Any]]:
    rules_path = Path(rules_path)
    return json.loads(rules_path.read_text(encoding="utf-8"))


def _save_rules_file(rules_path: str | Path, rules: dict[str, dict[str, Any]]) -> None:
    Path(rules_path).write_text(json.dumps(rules, ensure_ascii=False, indent=2), encoding="utf-8")


def _risk_level(quote_type: str) -> str:
    return {"固定单价": "low", "档位报价": "medium", "模糊报价": "high"}.get(quote_type, "medium")


def _enrich_rule(standard_item: str, rule: dict[str, Any]) -> dict[str, Any]:
    enriched = dict(rule)
    quote_type = enriched.get("quote_type", "档位报价")
    enriched.setdefault("standard_item", standard_item)
    enriched.setdefault("aliases", [])
    enriched.setdefault("auto_complete_with", [])
    enriched.setdefault("risk_level", _risk_level(quote_type))
    enriched.setdefault(
        "need_confirm_fields",
        ["数量"] if enriched["risk_level"] == "low" else ["数量", "规格", "配置", "天数"],
    )
    return enriched


def add_alias_to_standard_item(rules_path: str | Path, standard_item: str, alias: str) -> tuple[bool, str]:
    alias = alias.strip()
    if not alias:
        return False, "候选词为空，未写入。"

    rules = _load_rules_file(rules_path)
    if standard_item not in rules:
        return False, f"标准项目不存在：{standard_item}"

    for existing_item, rule in rules.items():
        existing_aliases = [str(item) for item in rule.get("aliases", [])]
        if alias == existing_item or alias in existing_aliases:
            if existing_item == standard_item:
                return False, f"别名已存在：{alias}"
            return False, f"别名已存在于“{existing_item}”：{alias}"

    aliases = rules[standard_item].setdefault("aliases", [])
    backup_rules_config(rules_path)
    aliases.append(alias)
    rules[standard_item] = _enrich_rule(standard_item, rules[standard_item])
    _save_rules_file(rules_path, rules)
    return True, f"已将“{alias}”加入“{standard_item}”的别名。"


def create_standard_item(
    rules_path: str | Path,
    standard_item: str,
    alias: str,
    category: str,
    quote_type: str,
    default_unit: str,
    default_desc: str,
) -> tuple[bool, str]:
    standard_item = standard_item.strip()
    alias = alias.strip()
    if not standard_item:
        return False, "标准项目为空，未写入。"

    rules = _load_rules_file(rules_path)
    if standard_item in rules:
        return False, f"标准项目已存在：{standard_item}"
    if alias:
        for existing_item, rule in rules.items():
            existing_aliases = [str(item) for item in rule.get("aliases", [])]
            if alias == existing_item or alias in existing_aliases:
                return False, f"候选词已存在于“{existing_item}”：{alias}"

    backup_rules_config(rules_path)
    rules[standard_item] = _enrich_rule(
        standard_item,
        {
            "category": category.strip() or "待分类",
            "aliases": [alias] if alias else [],
            "quote_type": quote_type,
            "default_unit": default_unit.strip() or "项",
            "default_desc": default_desc.strip() or "待补充说明",
        },
    )
    _save_rules_file(rules_path, rules)
    return True, f"已新建标准项目：{standard_item}"


def apply_feedback_rows(
    rules_path: str | Path,
    ignored_terms_path: str | Path,
    rows: list[dict[str, Any]],
    new_item_defaults: dict[str, str],
) -> dict[str, Any]:
    messages: list[str] = []
    rules_updated = 0
    ignored_count = 0

    for row in rows:
        should_process = bool(row.get("是否处理", row.get("是否加入规则库", False)))
        if not should_process:
            continue

        candidate = str(row.get("候选词", "")).strip()
        action = str(row.get("处理方式", row.get("选择处理方式", "暂不处理"))).strip()
        standard_item = str(row.get("选择标准项目", "")).strip()

        if action == "暂不处理":
            continue

        if action == "加入已有标准项目别名":
            if not standard_item:
                messages.append(f"请选择标准项目后再写入别名：{candidate}")
                continue
            ok, message = add_alias_to_standard_item(rules_path, standard_item, candidate)
            if ok:
                rules_updated += 1
            messages.append(message)
            continue

        if action == "新建标准项目":
            item_name = new_item_defaults.get("标准项目", "").strip() or candidate
            ok, message = create_standard_item(
                rules_path,
                item_name,
                candidate,
                new_item_defaults.get("项目分类", ""),
                new_item_defaults.get("报价类型", "模糊报价"),
                new_item_defaults.get("默认单位", "项"),
                new_item_defaults.get("默认说明", ""),
            )
            if ok:
                rules_updated += 1
            messages.append(message)
            continue

        if action == "标记为无需报价":
            changed = save_ignored_term(ignored_terms_path, candidate)
            if changed:
                ignored_count += 1
            messages.append(f"已标记为无需报价：{candidate}" if changed else f"无需报价词已存在：{candidate}")

    return {"messages": messages, "rules_updated": rules_updated, "ignored_count": ignored_count}
