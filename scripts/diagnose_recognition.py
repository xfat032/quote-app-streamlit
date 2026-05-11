from __future__ import annotations

import argparse
import sys
from pathlib import Path


ROOT_DIR = Path(__file__).resolve().parents[1]
QUOTE_APP_DIR = ROOT_DIR / "quote_app"
if str(QUOTE_APP_DIR) not in sys.path:
    sys.path.insert(0, str(QUOTE_APP_DIR))

from core.activity_classifier import classify_activity_types
from core.extractor import extract_quote_items
from core.normalizer import load_rules
from core.quote_builder import (
    build_quote_rows,
    diagnose_activity_content_ranges,
    diagnose_section_candidates,
    extract_activity_sections,
    prepare_confirmed_activity_sections,
)
from core.text_reader import read_text_from_path


RULES_PATH = QUOTE_APP_DIR / "data" / "rules_config.json"
PRICE_DB_PATH = QUOTE_APP_DIR / "data" / "price_db.xlsx"
SAMPLE_TEXTS = {
    "生蚝荟链路样例": """
活动内容
01 开幕式
1 签到仪式-蚝语心愿墙
2 开场舞-南海渔歌
3 启动仪式
4 主持推荐
湾畔茶席
活动大巡游
蚝王争霸赛
食速挑战赛
渡口星光音乐会
生产者大会（市集）
蚝趣游乐记
互动游戏一 套蚝赢趣
钓壳寻鲜
贝壳绘梦
曲口寻踪
你好，海洋！艺术展
蚝仔周边
发起线上互动话题 掀起全民分享热潮
06 12.07
""",
    "玉蕊花节链路样例": """
活动内容
四感共生，构建七里花月夜体验场
花间有声
清影花月音乐会
东坡夜话
调声之夜
玉蕊夜绽赏花荟
月下有色
十二花色打卡游
月下玉蕊·花灯展
集间有味
花下食集
花间茶寮
玉蕊花宴
游中有戏
花使巡游
玉蕊寻芳游园会
花笺留诗
夜绘玉蕊
""",
}


def _iter_inputs(paths: list[str]) -> list[tuple[str, str]]:
    if not paths:
        return list(SAMPLE_TEXTS.items())

    inputs: list[tuple[str, str]] = []
    for raw_path in paths:
        path = Path(raw_path)
        if path.is_dir():
            for suffix in ("*.txt", "*.docx", "*.pdf"):
                for child in sorted(path.glob(suffix)):
                    inputs.append((str(child), read_text_from_path(child)))
            continue
        inputs.append((str(path), read_text_from_path(path)))
    return inputs


def diagnose_text(name: str, text: str) -> str:
    rules = load_rules(RULES_PATH)
    extracted_rows = extract_quote_items(text, rules)
    sections = extract_activity_sections(text, extracted_rows)
    confirmed_sections = prepare_confirmed_activity_sections(sections)
    quote_rows = build_quote_rows(extracted_rows, PRICE_DB_PATH, text, activity_sections=sections)
    activity_types = classify_activity_types(text)
    range_diagnostics = diagnose_activity_content_ranges(text)
    filtered_candidates = diagnose_section_candidates(text)

    issues: list[str] = []
    if not confirmed_sections:
        issues.append("未识别到默认主活动板块")
    if range_diagnostics.get("has_short_range"):
        issues.append("活动内容区域可能切片错误")
    if range_diagnostics.get("directory_like"):
        issues.append("活动内容区域疑似误用了目录页")
    unassigned_count = sum(1 for row in quote_rows if str(row.get("quote_section", "")) == "未归属板块")
    if unassigned_count:
        issues.append(f"存在未归属项目 {unassigned_count} 项")
    if not quote_rows:
        issues.append("未识别到报价项")

    main_lines = [
        f"{section.get('name')}[{section.get('section_level')}|{section.get('channel', '')}]"
        for section in sections
        if section.get("section_level") == "main"
    ]
    sub_lines = [
        f"{section.get('name')} -> {section.get('parent') or '待确认'}[{section.get('channel', '')}]"
        for section in sections
        if section.get("section_level") == "sub"
    ]
    filtered_lines = [
        f"{row.get('候选标题')}({row.get('分类')}：{row.get('过滤原因')})"
        for row in filtered_candidates[:30]
    ]
    type_lines = [
        f"{row.get('活动类型')}({row.get('命中关键词', '')})"
        for row in activity_types
    ]

    output = [
        f"===== {name} =====",
        f"识别到的活动类型：{', '.join(type_lines) if type_lines else '无'}",
        f"活动板块 main：{', '.join(main_lines) if main_lines else '无'}",
        f"子活动 sub：{', '.join(sub_lines) if sub_lines else '无'}",
        f"被过滤候选及原因：{', '.join(filtered_lines) if filtered_lines else '无'}",
        f"默认主活动板块：{', '.join(section.get('name', '') for section in confirmed_sections) if confirmed_sections else '无'}",
        f"未归属项目数量：{unassigned_count}",
        f"报价项数量：{len(quote_rows)}",
        f"活动内容区域数量：{range_diagnostics.get('range_count', 0)}",
        f"活动内容区域总长度：{range_diagnostics.get('total_length', 0)}",
        f"可疑问题：{'; '.join(issues) if issues else '无'}",
    ]
    return "\n".join(output)


def main() -> None:
    parser = argparse.ArgumentParser(description="诊断活动板块识别链路")
    parser.add_argument("paths", nargs="*", help="待诊断的 .txt 文件或目录；不传则运行内置回归样例")
    args = parser.parse_args()

    for name, text in _iter_inputs(args.paths):
        print(diagnose_text(name, text))
        print()


if __name__ == "__main__":
    main()
