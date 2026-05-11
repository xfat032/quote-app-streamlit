from __future__ import annotations

import sys
from pathlib import Path


ROOT_DIR = Path(__file__).resolve().parents[1]
QUOTE_APP_DIR = ROOT_DIR / "quote_app"
if str(QUOTE_APP_DIR) not in sys.path:
    sys.path.insert(0, str(QUOTE_APP_DIR))

from core.extractor import extract_quote_items
from core.normalizer import load_rules
from core.quote_builder import build_quote_rows, extract_activity_sections


RULES_PATH = QUOTE_APP_DIR / "data" / "rules_config.json"
PRICE_DB_PATH = QUOTE_APP_DIR / "data" / "price_db.xlsx"
SAMPLE_DIR = ROOT_DIR / "scripts" / "samples" / "generated_regression"


CASES = [
    {
        "file": "01_evening_gala.txt",
        "expected_main": ["潮涌新区启幕晚会", "海风民乐迎宾秀", "青年合唱音乐会", "星光市集"],
        "forbidden_main": ["签到仪式", "开场舞", "领导致辞", "启动仪式", "倒计时海报", "公众号软文"],
    },
    {
        "file": "02_food_market_festival.txt",
        "expected_main": ["开街仪式", "海鲜美食市集", "食速挑战赛", "厨艺演出", "夜间音乐会"],
        "forbidden_main": ["立体字打卡点示意图", "活动地图", "话题互动招募", "活动回顾推文"],
    },
    {
        "file": "03_museum_exhibition.txt",
        "expected_main": ["序厅图文展", "桌面标本展", "生态互动展", "环境装置", "图书阅读区"],
        "forbidden_main": ["有声阅读二维码", "传播策略", "传播话题", "公众号软文", "预期成效"],
    },
    {
        "file": "04_reading_week.txt",
        "expected_main": ["入梦签到", "亲子共读剧场", "书海寻宝", "有声阅读二维码", "阅读心得墙"],
        "forbidden_main": ["趣味指引示意图", "倒计时海报", "活动回顾推文"],
    },
    {
        "file": "05_parent_child_nature.txt",
        "expected_main": ["自然课堂体验", "亲子手作体验", "植物拓印体验", "露营音乐会"],
        "forbidden_main": ["活动地图", "通行证", "亲子招募海报", "公众号软文"],
    },
    {
        "file": "06_fun_sports_carnival.txt",
        "expected_main": ["开幕式", "民族趣味运动会", "亲子挑战赛", "游园互动体验", "闭幕式"],
        "forbidden_main": ["套圈挑战", "扁担接力", "欢乐背背跑", "活动排期大纲", "现场图片直播"],
    },
    {
        "file": "07_light_art_night.txt",
        "expected_main": ["城市灯光艺术展", "桥下音乐会", "夜游打卡路线", "光影互动体验"],
        "forbidden_main": ["主画面背景板", "灯光装置示意图", "短视频", "媒体内容集中发布"],
    },
    {
        "file": "08_intangible_heritage_market.txt",
        "expected_main": ["非遗生活市集", "民族服饰走秀", "手作体验工坊", "非遗巡游", "夜间茶席雅集"],
        "forbidden_main": ["前期宣传", "倒计时海报", "活动回顾推文", "预期成效"],
    },
    {
        "file": "09_rural_harvest_festival.txt",
        "expected_main": ["丰收开幕式", "农产品市集", "田间音乐会", "稻田长桌宴", "亲子采摘体验"],
        "forbidden_main": ["活动地图", "工作证", "主画面背景板", "海报及图文推送"],
    },
    {
        "file": "10_ocean_culture_day.txt",
        "expected_main": ["海洋科普展", "海岛服饰走秀", "海边音乐会", "沙滩互动体验"],
        "forbidden_main": ["活动地点：海口湾沙滩", "趣味指引示意图", "祝福视频", "公众号软文"],
    },
    {
        "file": "11_public_welfare_green.txt",
        "expected_main": ["公益启动仪式", "志愿者市集", "旧物换新体验", "环保手作体验"],
        "forbidden_main": ["签到仪式", "领导致辞", "合影留念", "话题互动招募", "活动回顾推文"],
    },
    {
        "file": "12_investment_promotion.txt",
        "expected_main": ["城市招商推介会", "项目路演", "签约仪式", "产业市集", "商务茶席"],
        "forbidden_main": ["宣传片", "公众号软文", "媒体内容集中发布"],
    },
]


def _names_by_level(sections: list[dict], level: str) -> list[str]:
    return [str(section.get("name", "")) for section in sections if str(section.get("section_level", "")) == level]


def _assert_contains(label: str, actual: list[str], expected: list[str]) -> None:
    missing = [name for name in expected if name not in actual]
    if missing:
        raise AssertionError(f"{label} missing: {missing}; actual={actual}")


def _assert_not_contains(label: str, actual: list[str], forbidden: list[str]) -> None:
    leaked = [name for name in forbidden if name in actual]
    if leaked:
        raise AssertionError(f"{label} leaked: {leaked}; actual={actual}")


def main() -> None:
    rules = load_rules(RULES_PATH)
    total_quote_rows = 0
    total_unassigned = 0

    for case in CASES:
        path = SAMPLE_DIR / case["file"]
        text = path.read_text(encoding="utf-8")
        extracted_rows = extract_quote_items(text, rules)
        sections = extract_activity_sections(text, extracted_rows)
        quote_rows = build_quote_rows(extracted_rows, PRICE_DB_PATH, text, activity_sections=sections)
        main_names = _names_by_level(sections, "main")
        sub_names = _names_by_level(sections, "sub")
        unassigned_count = sum(1 for row in quote_rows if str(row.get("quote_section", "")) == "未归属板块")

        _assert_contains(case["file"], main_names, case["expected_main"])
        _assert_not_contains(case["file"], main_names, case["forbidden_main"])
        if not quote_rows:
            raise AssertionError(f"{case['file']} produced no quote rows")
        if unassigned_count:
            raise AssertionError(f"{case['file']} has unassigned quote rows: {unassigned_count}")

        total_quote_rows += len(quote_rows)
        total_unassigned += unassigned_count
        print(
            f"PASS {case['file']}: main={len(main_names)} sub={len(sub_names)} "
            f"quote_rows={len(quote_rows)} unassigned={unassigned_count}"
        )

    print(
        f"ALL GENERATED SCHEME SAMPLE TESTS PASSED: cases={len(CASES)} "
        f"quote_rows={total_quote_rows} unassigned={total_unassigned}"
    )


if __name__ == "__main__":
    main()
