"""Streamlit app paths and UI table constants."""

from __future__ import annotations

from pathlib import Path

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
