from __future__ import annotations

import sys
from pathlib import Path


ROOT_DIR = Path(__file__).resolve().parents[1]
QUOTE_APP_DIR = ROOT_DIR / "quote_app"
if str(QUOTE_APP_DIR) not in sys.path:
    sys.path.insert(0, str(QUOTE_APP_DIR))

from core.extractor import extract_quote_items
from core.normalizer import load_rules
from core.quote_builder import extract_activity_sections
from core.quote_builder import build_quote_rows


def _main_names(text: str) -> list[str]:
    return [
        str(section.get("name", ""))
        for section in extract_activity_sections(text)
        if str(section.get("section_level", "")) == "main"
    ]


def _sub_names(text: str) -> list[str]:
    return [
        str(section.get("name", ""))
        for section in extract_activity_sections(text)
        if str(section.get("section_level", "")) == "sub"
    ]


def _assert_contains(label: str, actual: list[str], expected: list[str]) -> None:
    missing = [name for name in expected if name not in actual]
    if missing:
        raise AssertionError(f"{label} missing: {missing}; actual={actual}")


def _assert_not_contains(label: str, actual: list[str], forbidden: list[str]) -> None:
    leaked = [name for name in forbidden if name in actual]
    if leaked:
        raise AssertionError(f"{label} leaked: {leaked}; actual={actual}")


def test_region_titles() -> None:
    text = """
空间规划
区域A：市集区
区域B：舞台区
A区：互动体验区
B区：非遗展陈区
"""
    main = _main_names(text)
    _assert_contains("region main", main, ["市集区", "舞台区", "互动体验区", "非遗展陈区"])


def test_experience_activity_title() -> None:
    text = """
活动内容
1、亲子手作体验
2、民族服饰体验
3、非遗拓印体验
"""
    main = _main_names(text)
    _assert_contains("experience main", main, ["亲子手作体验", "民族服饰体验", "非遗拓印体验"])


def test_gala_colon_title() -> None:
    text = """
活动内容
（一）主晚会：潮涌海湾文旅焕新夜
1. 签到仪式
2. 开场舞
安排主题开场舞节目演出。
"""
    main = _main_names(text)
    sub = _sub_names(text)
    _assert_contains("gala main", main, ["潮涌海湾文旅焕新夜"])
    _assert_not_contains("gala main", main, ["主晚会 潮涌海湾文旅焕新夜"])
    _assert_contains("gala sub", sub, ["签到仪式", "开场舞"])
    _assert_not_contains("gala sub", sub, ["安排主题开场舞节目演出"])


def test_prefixed_ceremony_stays_main() -> None:
    text = """
现场活动
1、公益启动仪式
设置舞台搭建、主持人和灯光音响套装。
2、志愿者市集
设置帐篷摊位。
"""
    main = _main_names(text)
    sub = _sub_names(text)
    _assert_contains("prefixed ceremony main", main, ["公益启动仪式", "志愿者市集"])
    _assert_not_contains("prefixed ceremony sub", sub, ["公益启动仪式"])


def test_exhibit_area_dedupe() -> None:
    text = """
展区规划
展区一：图文展区
展区二：桌面展区
展区三：标本展区
展区四：互动展区
"""
    main = _main_names(text)
    _assert_contains("exhibit main", main, ["图文展区", "桌面展区", "标本展区", "互动展区"])
    _assert_not_contains("exhibit main", main, ["展区一 图文展区", "展区二 桌面展区", "展区三 标本展区", "展区四 互动展区"])
    if len(main) != len(set(main)):
        raise AssertionError(f"exhibit main duplicated: {main}")


def test_whitelist_quote_item_sentence_not_main() -> None:
    text = """
展览内容
（一）图书阅读区
配置图书配置、阅读区布置、桌椅布置和有声阅读二维码。
"""
    main = _main_names(text)
    _assert_contains("reading zone main", main, ["图书阅读区"])
    _assert_not_contains("reading zone main", main, ["有声阅读二维码"])


def test_only_sub_sections_survive() -> None:
    text = """
活动内容
1. 签到仪式
2. 开场舞
3. 领导致辞
"""
    main = _main_names(text)
    sub = _sub_names(text)
    if main:
        raise AssertionError(f"only-sub sample should not create main sections: {main}")
    _assert_contains("only-sub", sub, ["签到仪式", "开场舞", "领导致辞"])


def test_promo_noise_filtered() -> None:
    text = """
宣传排期
预热期
宣传片
话题互动招募
爆发期
现场图片直播
"""
    main = _main_names(text)
    _assert_not_contains("promo noise", main, ["预热期", "宣传片", "话题互动招募", "爆发期", "现场图片直播"])
    if main:
        raise AssertionError(f"promo sample should not create main sections: {main}")


def test_nonheritage_content_boundary_stops_before_promo() -> None:
    text = """
活动时间与地点
（一）第一场：椰城非遗生活节·夏日场
活动时间：2026年6月13日至6月14日。
（二）第二场：椰城非遗生活节·秋日场
活动时间：2026年11月21日至11月22日。
活动内容规划
（一）非遗主题展示
设置图文展板和项目介绍牌。
（二）非遗展演
设置舞台、灯光和音响。
（三）非遗互动体验
设置椰雕纹样拓印体验。
（四）非遗市集
设置非遗美食摊位。
（五）非遗文创潮玩
开发贴纸、徽章和明信片。
五、现场空间规划
入口签到区设置活动规则说明、集章卡领取点和官方二维码。
六、宣传推广计划
（一）活动预热
发布活动预告图文。
（二）现场传播
开展图片直播和短视频剪辑。
（三）活动复盘
制作活动回顾视频。
"""
    main = _main_names(text)
    _assert_contains("nonheritage main", main, ["非遗主题展示", "非遗展演", "非遗互动体验", "非遗市集", "非遗文创潮玩"])
    _assert_not_contains("nonheritage main", main, ["第一场", "第二场", "现场空间规划", "活动预热", "现场传播", "活动复盘"])


def test_generic_inferred_sections_do_not_override_confirmed_sections() -> None:
    text = """
活动内容规划
（一）非遗市集
设置非遗美食摊位和摊位楣板。
五、现场空间规划
入口领取集章卡、互动体验参与、市集游逛消费，活动地点为海口湾云洞图书馆外广场。
"""
    rules = load_rules(QUOTE_APP_DIR / "data" / "rules_config.json")
    quote_rows = build_quote_rows(
        extract_quote_items(text, rules),
        QUOTE_APP_DIR / "data" / "price_db.xlsx",
        text,
        activity_sections=extract_activity_sections(text),
    )
    sections = {str(row.get("quote_section", "")) for row in quote_rows}
    forbidden = {"市集活动", "互动体验", "阅读活动"}
    leaked = sorted(sections & forbidden)
    if leaked:
        raise AssertionError(f"generic inferred sections leaked into quote rows: {leaked}; sections={sections}")
    items = {str(row.get("标准项目", "")) for row in quote_rows}
    if "图书配置" in items:
        raise AssertionError(f"venue name 图书馆 should not create 图书配置: {items}")


def main() -> None:
    tests = [
        test_region_titles,
        test_experience_activity_title,
        test_gala_colon_title,
        test_prefixed_ceremony_stays_main,
        test_exhibit_area_dedupe,
        test_whitelist_quote_item_sentence_not_main,
        test_only_sub_sections_survive,
        test_promo_noise_filtered,
        test_nonheritage_content_boundary_stops_before_promo,
        test_generic_inferred_sections_do_not_override_confirmed_sections,
    ]
    for test in tests:
        test()
        print(f"PASS {test.__name__}")
    print("ALL STRUCTURE RECOGNITION TESTS PASSED")


if __name__ == "__main__":
    main()
