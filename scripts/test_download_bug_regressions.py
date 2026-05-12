from __future__ import annotations

import sys
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile


ROOT_DIR = Path(__file__).resolve().parents[1]
QUOTE_APP_DIR = ROOT_DIR / "quote_app"
if str(QUOTE_APP_DIR) not in sys.path:
    sys.path.insert(0, str(QUOTE_APP_DIR))

from core.extractor import extract_quote_items
from core.normalizer import load_rules
from core.quote_builder import build_quote_rows, diagnose_activity_content_ranges, extract_activity_sections
from core.text_reader import read_text_from_path


RULES_PATH = QUOTE_APP_DIR / "data" / "rules_config.json"
PRICE_DB_PATH = QUOTE_APP_DIR / "data" / "price_db.xlsx"


def _main_names(text: str) -> list[str]:
    return [
        str(section.get("name", ""))
        for section in extract_activity_sections(text)
        if str(section.get("section_level", "")) == "main"
    ]


def _assert_not_contains(label: str, actual: list[str], forbidden: list[str]) -> None:
    leaked = [name for name in forbidden if name in actual]
    if leaked:
        raise AssertionError(f"{label} leaked: {leaked}; actual={actual}")


def _assert_contains(label: str, actual: list[str], expected: list[str]) -> None:
    missing = [name for name in expected if name not in actual]
    if missing:
        raise AssertionError(f"{label} missing: {missing}; actual={actual}")


def test_meta_titles_do_not_become_main_sections() -> None:
    text = """
（一）活动主题：临海阅书香 自在读海口
（二）活动口号：让阅读回到生活
（三）主题解读：城市与阅读共同生长
（四）宏观背景：数字化成主流
（五）微观痛点：市民阅读面临三难
（六）活动效益预估
活动内容
1. 开幕式
2. 非遗市集
"""
    main = _main_names(text)
    _assert_contains("meta real main", main, ["开幕式", "非遗市集"])
    _assert_not_contains(
        "meta main",
        main,
        ["活动主题", "活动口号", "主题解读", "宏观背景", "微观痛点", "活动效益预估"],
    )


def test_visual_material_and_schedule_titles_do_not_become_main_sections() -> None:
    text = """
（一）活动铺排表
9 月 15 日-17 日
（二）活动区域示意
（三）美陈 DP
1、镜月舞台：
2、海月流光：
3、月光透镜：
活动内容
1、海月市集
2、弦月音乐会
"""
    main = _main_names(text)
    _assert_contains("visual real main", main, ["海月市集", "弦月音乐会"])
    _assert_not_contains(
        "visual main",
        main,
        ["活动铺排表", "9 月 15 日-17 日", "活动区域示意", "美陈 DP", "镜月舞台", "海月流光", "月光透镜"],
    )


def test_pdf_table_fragments_do_not_become_main_sections() -> None:
    text = """
2.0X
1 2 3 4
士、沙滩音乐会等
营地、市集、农
... ... ...
东坡老码头 户外
活动内容
1. 海口乡村旅游季启动仪式
2. “村游”市集
3. 沙滩音乐会
"""
    main = _main_names(text)
    _assert_contains("pdf real main", main, ["海口乡村旅游季启动仪式", "“村游”市集", "沙滩音乐会"])
    _assert_not_contains(
        "pdf fragment main",
        main,
        ["0X", "3 4", "士、沙滩音乐会等", "营地、市集、农", "... ... ...", "东坡老码头 户外"],
    )


def test_short_inline_activity_summary_is_not_used_as_content_range() -> None:
    text = """
活动内容：启动仪式、村游玩法地图、村理人圆桌会、“村游”市集，乡土手作体验、荒野大地餐桌、音乐live、城市主理人计划、“村游”巴士、沙滩音乐会等
组织架构：主办单位
活动内容
1. 海口乡村旅游季启动仪式
设置舞台、主持人、启动仪式和活动导视，并安排来宾签到、领导致辞、合影留念等流程。现场配置基础执行人员、摄影摄像、安保和应急保障，确保主会场活动稳定执行。
2. “村游”市集
设置乡土手作体验、农产品摊位、特色小吃摊位、村游玩法地图和服务台，形成完整游逛动线。现场安排执行人员维持秩序，并通过导视指引、市集摊位包装和互动规则牌帮助游客理解参与方式。
3. 沙滩音乐会
设置小型舞台、灯光音响、乐队演出和观众休息区，形成傍晚活动聚集点。现场结合摄影摄像、视频快剪和公共导视，为后续传播提供素材。
4. 村理人圆桌会
邀请乡村主理人、民宿经营者、农产品品牌代表进行交流分享，配置基础会议桌椅、主持人、摄影摄像和现场执行人员。圆桌会围绕村游路线、产品转化和游客服务展开，形成活动内容闭环。
5. 乡土手作体验
设置剪纸、草编、拓印等体验项目，配置手作体验材料、体验老师、互动规则牌、打卡装置和活动桌椅。参与者可在现场完成作品并带走，增强活动参与感和传播素材质量。
6. 乡村运动挑战赛
设置轻量趣味互动游戏、赛事规则牌、互动道具、奖品和执行人员，形成适合亲子家庭参与的运动挑战内容。
"""
    diagnostics = diagnose_activity_content_ranges(text)
    if diagnostics.get("has_short_range"):
        raise AssertionError(f"short inline summary should not be treated as activity range: {diagnostics}")
    main = _main_names(text)
    _assert_contains("inline summary real main", main, ["海口乡村旅游季启动仪式", "“村游”市集"])


def test_common_support_items_do_not_remain_unassigned() -> None:
    text = """
活动氛围布置
结合集章地图设置打卡标识柱/小道旗，引导市民完成“集章任务”与空间探索。
活动内容
1. 开幕式
2. 非遗购物节
非遗购物节设置传统老爸茶摊位，提供海南特色茶饮和茶点。
"""
    rules = load_rules(RULES_PATH)
    quote_rows = build_quote_rows(
        extract_quote_items(text, rules),
        PRICE_DB_PATH,
        text,
        activity_sections=extract_activity_sections(text),
    )
    unassigned = [row for row in quote_rows if str(row.get("quote_section", "")) == "未归属板块"]
    if unassigned:
        raise AssertionError(f"support items should not stay unassigned: {unassigned}")


def test_orphan_explicit_items_get_public_or_other_sections() -> None:
    text = """
年味创意礼品、非遗手工摊位和特色摊位构成入口氛围。
现场设置品牌联动、贴纸投票、NPC大巡游、星光音乐会和轻量竞赛。
"""
    activity_sections = [
        {
            "name": "非遗分享会",
            "start": 10000,
            "end": 11000,
            "selected": True,
            "section_level": "main",
        }
    ]
    rules = load_rules(RULES_PATH)
    quote_rows = build_quote_rows(
        extract_quote_items(text, rules),
        PRICE_DB_PATH,
        text,
        activity_sections=activity_sections,
    )
    unassigned = [row for row in quote_rows if str(row.get("quote_section", "")) == "未归属板块"]
    if unassigned:
        raise AssertionError(f"orphan explicit items should not stay unassigned: {unassigned}")

    section_by_item = {str(row.get("标准项目", "")): str(row.get("quote_section", "")) for row in quote_rows}
    if section_by_item.get("合作品牌联合宣传") != "活动宣传":
        raise AssertionError(f"brand collaboration should be promotion: {section_by_item}")
    if section_by_item.get("帐篷摊位") != "其他搭建类":
        raise AssertionError(f"orphan tent stalls should be public build: {section_by_item}")


def test_broken_docx_null_relationship_fallback(tmp_dir: Path | None = None) -> None:
    temp_dir = Path("/tmp/quote_app_regression_docx")
    temp_dir.mkdir(parents=True, exist_ok=True)
    path = temp_dir / "broken_null_relationship.docx"
    with ZipFile(path, "w", ZIP_DEFLATED) as archive:
        archive.writestr(
            "[Content_Types].xml",
            """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>""",
        )
        archive.writestr(
            "_rels/.rels",
            """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>""",
        )
        archive.writestr(
            "word/_rels/document.xml.rels",
            """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../NULL"/>
</Relationships>""",
        )
        archive.writestr(
            "word/document.xml",
            """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>海南三月三活动方案</w:t></w:r></w:p>
    <w:p><w:r><w:t>活动内容</w:t></w:r></w:p>
    <w:p><w:r><w:t>开幕式</w:t></w:r></w:p>
  </w:body>
</w:document>""",
        )

    text = read_text_from_path(path)
    if "海南三月三活动方案" not in text or "开幕式" not in text:
        raise AssertionError(f"broken docx fallback failed: {text}")


def main() -> None:
    tests = [
        test_meta_titles_do_not_become_main_sections,
        test_visual_material_and_schedule_titles_do_not_become_main_sections,
        test_pdf_table_fragments_do_not_become_main_sections,
        test_short_inline_activity_summary_is_not_used_as_content_range,
        test_common_support_items_do_not_remain_unassigned,
        test_orphan_explicit_items_get_public_or_other_sections,
        test_broken_docx_null_relationship_fallback,
    ]
    for test in tests:
        test()
        print(f"PASS {test.__name__}")
    print("ALL DOWNLOAD BUG REGRESSION TESTS PASSED")


if __name__ == "__main__":
    main()
