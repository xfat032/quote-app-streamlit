"""Build editable quote rows and enrich them with the empty price DB template."""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any

import pandas as pd
from openpyxl import Workbook


PRICE_DB_COLUMNS = ["项目分类", "标准项目", "默认规格", "单位", "默认单价", "报价类型", "备注"]
QUOTE_COLUMNS = [
    "是否保留",
    "项目分类",
    "标准项目",
    "原始命中词",
    "内容/尺寸/工艺",
    "数量",
    "单位",
    "单价",
    "合计",
    "报价类型",
    "来源状态",
    "备注",
    "internal_category",
    "quote_section",
    "quote_section_order",
    "匹配位置",
    "命中模块",
    "source_context_start",
    "source_context_text",
    "matched_section_name",
    "evidence_type",
    "evidence_text",
    "trigger_module",
]

CATEGORY_ORDER = [
    "宣传推广类",
    "场景搭建类",
    "活动内容类",
    "活动物料类",
    "设备租赁类",
    "人员服务类",
    "后勤保障类",
]

CATEGORY_MAP = {
    "宣传推广类": "宣传推广类",
    "宣传拍摄类": "宣传推广类",
    "视觉设计类": "宣传推广类",
    "文创物料类": "活动物料类",
    "舞台搭建类": "场景搭建类",
    "美陈装置类": "场景搭建类",
    "展览展陈类": "场景搭建类",
    "氛围装饰类": "场景搭建类",
    "活动布置类": "场景搭建类",
    "市集摊位类": "场景搭建类",
    "演艺节目类": "活动内容类",
    "互动体验类": "活动内容类",
    "餐饮体验类": "活动内容类",
    "执行服务类": "活动内容类",
    "活动道具类": "活动物料类",
    "灯光音响类": "设备租赁类",
    "人员执行类": "人员服务类",
    "后勤物料类": "后勤保障类",
    "其他类": "后勤保障类",
    "待分类": "后勤保障类",
    "视觉制作类": "场景搭建类",
    "视觉物料类": "场景搭建类",
}

SCENE_KEYWORDS = [
    "舞台",
    "背景板",
    "舞台背景",
    "舞台视觉包装",
    "签到墙",
    "签到背景墙",
    "展示墙",
    "故事画面墙",
    "阅读心得墙",
    "展板",
    "图文展板",
    "知识展板",
    "主题展板",
    "艺术展陈",
    "展览展陈",
    "展架",
    "指引牌",
    "导视牌",
    "导视指引",
    "区域指引",
    "趣味指引",
    "入口指引",
    "活动地图墙",
    "打卡装置",
    "打卡点位",
    "立体字打卡装置",
    "美陈装置",
    "主题美陈",
    "互动装置",
    "灯光艺术装置",
    "花灯装置",
    "花艺花境",
    "氛围布置",
    "现场氛围",
    "阅读区布置",
    "互动区布置",
    "市集布置",
    "帐篷摊位",
    "市集摊位",
    "地摊市集布置",
    "摊位门头",
    "摊位楣板",
    "摊位装饰",
    "摊位包装",
    "摊位视觉物料",
]

MATERIAL_KEYWORDS = [
    "工作证",
    "工作牌",
    "工作卡",
    "胸卡",
    "人员证件",
    "通行证",
    "集章卡",
    "打卡卡",
    "任务卡",
    "体验券",
    "餐券",
    "活动书签",
    "定制书签",
    "明信片",
    "贴纸",
    "姓名贴",
    "签到贴纸",
    "投票贴纸",
    "手卡",
    "主持手卡",
    "流程手卡",
    "印章",
    "活动印章",
    "主题卡",
    "规则手册",
    "活动地图折页",
    "攻略手册",
    "游玩攻略手册",
    "文创礼品",
    "小礼品",
    "奖品",
    "奖品礼包",
    "纪念品",
    "徽章",
    "帆布袋",
    "环保袋",
    "IP周边",
    "互动道具",
    "游戏道具",
    "手作材料",
    "拓印材料",
    "贝壳绘梦材料",
    "香囊材料",
    "标本材料",
    "活动贴纸",
    "IP周边制作",
    "工作证件",
    "互动规则牌",
]

CONTENT_KEYWORDS = [
    "节目演出",
    "开场表演",
    "节目表演",
    "乐队演绎",
    "舞蹈演绎",
    "脱口秀演绎",
    "非遗演绎",
    "琼剧表演",
    "民族器乐演绎",
    "音乐会",
    "巡游表演",
    "服饰走秀",
    "活动大巡游",
    "NPC互动服务",
    "趣味互动游戏",
    "赛事活动",
    "游园会",
    "问答互动",
    "手作体验",
    "非遗拓印体验",
    "阅读活动",
    "共读活动",
    "长桌宴",
    "餐饮体验",
    "茶席茶寮",
    "市集招募服务",
    "商家招募服务",
    "招募服务",
]

EQUIPMENT_KEYWORDS = [
    "灯光音响套装",
    "音响设备",
    "线阵音响",
    "扩声音响",
    "舞台灯光",
    "灯光设备",
    "帕灯",
    "面光灯",
    "射灯",
    "光束灯",
    "灯架",
    "控台",
    "调音台",
    "麦克风",
    "手持麦",
    "对讲机",
    "摄影设备",
    "直播设备",
    "发电机",
    "移动发电车",
    "电力设备",
]

PERSONNEL_KEYWORDS = [
    "主持人",
    "执行人员",
    "工作人员",
    "现场执行",
    "点位人员",
    "志愿者",
    "安保人员",
    "保洁人员",
    "后勤保障人员",
    "搭建撤场工人",
    "技术人员",
    "音响师",
    "灯光师",
    "摄影师",
    "摄像师",
    "医疗人员",
    "救生员",
    "交通引导人员",
    "停车引导人员",
    "礼仪人员",
    "体验老师",
    "非遗传承人",
    "NPC演员",
    "演员",
    "模特",
    "交通停车引导",
    "医疗保障",
]

LOGISTICS_KEYWORDS = [
    "运输费",
    "物料运输",
    "物流运输",
    "餐费",
    "工作餐",
    "饮用水",
    "矿泉水",
    "信息咨询台",
    "服务台",
    "失物招领",
    "铁马",
    "围挡",
    "警戒带",
    "安全线",
    "应急物资",
    "急救包",
    "灭火器",
    "雨衣",
    "防滑垫",
    "电力保障",
    "临时用电",
    "配电",
    "备用电源",
    "保险",
    "不可预见费",
    "管理费",
    "税费",
]

PROMOTION_KEYWORDS = [
    "海报",
    "设计服务",
    "宣传服务",
    "传播服务",
    "公众号软文",
    "活动推文",
    "新闻稿",
    "媒体报道",
    "主流媒体宣传",
    "媒体发布",
    "达人KOL推广",
    "网红博主推广",
    "直播宣传",
    "话题传播",
    "线上话题",
    "广告投放",
    "短视频推广",
    "预热视频",
    "宣传片",
    "视频快剪",
    "图片直播",
    "照片直播",
    "摄影摄像服务",
    "传播图文",
    "短视频内容",
    "活动记录",
    "官方媒体宣传",
    "社交平台传播",
    "小红书推广",
    "抖音推广",
    "微博话题",
    "主视觉设计",
    "KV设计",
    "活动海报设计",
    "倒计时海报设计",
    "延展设计",
    "物料画面设计",
    "导视系统设计",
    "舞台视觉设计",
    "摊位视觉设计",
    "IP形象设计",
    "品牌视觉设计",
    "宣传物料设计",
    "活动视觉设计",
    "美陈画面设计",
    "展板画面设计",
    "背景板画面设计",
    "视觉设计",
    "画面设计",
]

QUOTE_SECTION_PUBLIC_ORDER = ["活动宣传", "美陈搭建类", "其他搭建类", "未归属板块", "人员类及其他"]
PUBLIC_QUOTE_SECTIONS = set(QUOTE_SECTION_PUBLIC_ORDER)
PUBLIC_MERGE_ITEMS = {
    "摄影摄像服务",
    "图片直播",
    "视频快剪",
    "预热视频",
    "公众号软文",
    "媒体宣传",
    "主流媒体宣传",
    "话题传播",
    "达人KOL推广",
    "设计服务",
    "执行人员",
    "工作人员",
    "安保人员",
    "志愿者",
    "技术人员",
    "搭建撤场工人",
    "运输费",
    "餐费",
    "饮用水",
    "对讲机",
    "应急物资",
    "电力保障",
    "全场氛围布置",
    "主入口装置",
    "大型美陈装置",
    "公共导视",
    "导视指引",
    "指引牌",
    "服务台",
    "铁马围挡",
}
DESIGN_TRIGGER_KEYWORDS = [
    "设计服务",
    "主视觉设计",
    "KV设计",
    "视觉设计",
    "延展设计",
    "画面设计",
    "物料设计",
    "品牌视觉",
    "IP形象设计",
    "导视系统设计",
]
ACTIVITY_SECTION_CANDIDATES = [
    "开幕式",
    "闭幕式",
    "入梦签到",
    "奇幻故事共读",
    "梦醒留念",
    "活动宣传",
    "预热宣传",
    "现场活动",
    "主舞台活动",
    "星空茶谷音乐会",
    "茶园日落剧场",
    "得闲煮茶",
    "茶肆新春市集",
    "互动活动体验",
    "游园会",
    "市集活动",
    "美食市集",
    "非遗体验",
    "阅读活动",
    "书海寻宝",
    "白日梦阅读记",
    "一分钟阅读挑战",
    "阅读盲选快答",
    "候船微阅读角",
    "诗意阶梯",
    "双语阅读角",
    "有声阅读二维码",
    "阅读心得墙",
    "图文展",
    "桌面展",
    "标本展",
    "环境装置",
    "互动展",
    "图书阅读区",
    "传播与互动",
    "海边民族联欢会",
    "民族服饰走秀",
    "海岛服饰走秀",
    "特色长桌宴",
    "中华民族共享的文化符号装置",
    "民族知识知多少",
    "三三有礼市集",
    "黎韵绣面",
    "民族服饰体验",
    "民族风写真体验",
    "民族趣味运动会",
    "清影花月音乐会",
    "调声之夜",
    "东坡夜话",
    "玉蕊夜绽赏花荟",
    "月下玉蕊·灯光艺术展",
    "十二花色打卡游",
    "花下食集",
    "玉蕊花宴",
    "花间茶寮",
    "花使巡游",
    "玉蕊寻芳游",
    "花笺留诗",
    "夜绘玉蕊",
    "灯光艺术展",
    "艺术展",
    "巡游活动",
    "音乐会",
    "赛事活动",
    "打卡活动",
    "市集招募",
    "活动大巡游",
    "蚝王争霸赛",
    "食速挑战赛",
    "渡口星光音乐会",
    "生产者大会",
    "生产者大会（市集）",
    "蚝趣游乐记",
    "你好，海洋！艺术展",
    "湾畔茶席",
    "活动保障",
    "应急保障",
    "人员保障",
]
SECTION_IGNORE_TITLES = {
    "活动概述",
    "活动名称",
    "活动时间",
    "活动地点",
    "活动定位",
    "活动调性",
    "活动创意",
    "活动特色",
    "活动内容规划",
    "基本信息",
    "组织架构",
    "安全保障",
    "设备保障",
    "物料保障",
    "后勤保障",
    "项目背景",
    "项目分析",
    "传播目标",
    "经济目标",
    "文化目标",
    "活动目标",
    "参与对象",
    "预计参与人数",
    "核心活动",
    "配套活动",
    "互动机制",
    "候船场景转化为文化体验场景",
    "强化阅读活动与海口文旅场景的关联印象",
    "实现活动整体曝光量不少于10万次",
}
GENERIC_SECTION_TITLES = {"核心活动", "配套活动", "现场活动", "活动内容", "分项活动"}
PROMOTION_CONTAINER_HINTS = {
    "传播与互动",
    "传播策略",
    "内容驱动传播",
    "多渠道协同传播",
    "用户共创传播",
    "传播阶段",
    "预热阶段",
    "活动阶段",
    "复盘阶段",
    "优质内容激励",
    "话题引导",
    "宣传内容",
}
ACTIVITY_SECTION_KEYWORDS = [
    "活动",
    "开幕式",
    "闭幕式",
    "音乐会",
    "演唱",
    "演出",
    "走秀",
    "巡游",
    "市集",
    "食集",
    "茶席",
    "茶寮",
    "长桌宴",
    "花宴",
    "游园",
    "互动",
    "体验",
    "赛事",
    "挑战",
    "争霸赛",
    "展览",
    "艺术展",
    "灯光艺术展",
    "打卡",
    "阅读",
    "共读",
    "寻宝",
    "非遗",
    "拓印",
    "手作",
    "花灯",
    "夜游",
    "赏花",
    "展",
    "展区",
    "区域",
]
SHORT_SECTION_TITLES = {
    "活动宣传",
    "现场活动",
    "前期活动",
    "分项活动",
    "活动内容",
    "活动保障",
    "宣传规划",
    "预热宣传",
    "中期引爆",
    "长尾延续",
}
STRONG_ACTIVITY_KEYWORDS = [
    "开幕式",
    "闭幕式",
    "音乐会",
    "演唱",
    "演出",
    "走秀",
    "巡游",
    "市集",
    "食集",
    "茶席",
    "茶寮",
    "长桌宴",
    "花宴",
    "游园会",
    "互动体验",
    "非遗体验",
    "手作体验",
    "阅读活动",
    "共读",
    "寻宝",
    "打卡游",
    "艺术展",
    "灯光艺术展",
    "赏花荟",
    "趣味运动会",
    "挑战赛",
    "争霸赛",
    "问答",
    "签到仪式",
    "启动仪式",
]
SECTION_PREFIX_RE = re.compile(
    r"^(?:"
    r"[一二三四五六七八九十]+[、.]|"
    r"[（(][一二三四五六七八九十]+[)）]|"
    r"\d+(?:\.\d+){0,2}[、.)．]"
    r")\s*"
)
SECTION_SPACE_NUMBER_PREFIX_RE = re.compile(r"^\d{1,2}\s+")
SECTION_BLOCK_RE = re.compile(r"^(?:（[一二三四五六七八九十]+）)?\s*板块[一二三四五六七八九十]+[：:]\s*(.+)$")
SECTION_EXTEND_RE = re.compile(r"^[一二三四五六七八九十]+、\s*延展活动[：:]\s*(.+)$")
SECTION_CN_NUMBER_RE = re.compile(r"^\s*\d+[、.]\s*(.{2,30})$")
SECTION_SPACE_NUMBER_TITLE_RE = re.compile(r"^\s*\d{1,2}\s+(.{2,30})$")
SECTION_PAREN_TITLE_RE = re.compile(r"^\s*[（(][一二三四五六七八九十]+[)）]\s*(.{2,30})$")
SECTION_AREA_RE = re.compile(r"^\s*(?:区域\s*[A-ZＡ-Ｚ]|[A-ZＡ-Ｚ]\s*区)[：:]\s*(.{2,30})$", re.IGNORECASE)
SECTION_EXHIBIT_AREA_RE = re.compile(r"^\s*展区[一二三四五六七八九十\d]+[：:]\s*(.{2,30})$")
SECTION_EXTEND_INLINE_RE = re.compile(r"^\s*(?:[一二三四五六七八九十]+、\s*)?延展活动[：:]\s*(.{2,40})$")
SECTION_NAMED_CONTAINER_PREFIX_RE = re.compile(
    r"^(?:主晚会|主题晚会|启动晚会|颁奖晚会|文艺晚会|迎宾晚会|主活动|主题活动)[：:]\s*(.{2,40})$"
)
SECTION_NAMED_CONTAINER_PREFIX_SPACE_RE = re.compile(
    r"^(?:主晚会|主题晚会|启动晚会|颁奖晚会|文艺晚会|迎宾晚会|主活动|主题活动)\s+(.{2,40})$"
)
STRUCTURE_PATTERNS = [
    ("板块标题", SECTION_BLOCK_RE),
    ("延展活动标题", SECTION_EXTEND_INLINE_RE),
    ("展区标题", SECTION_EXHIBIT_AREA_RE),
    ("区域标题", SECTION_AREA_RE),
    ("中文括号活动/展区标题", SECTION_PAREN_TITLE_RE),
    ("数字编号活动标题", SECTION_CN_NUMBER_RE),
    ("数字编号活动标题", SECTION_SPACE_NUMBER_TITLE_RE),
]
TIME_LINE_RE = re.compile(
    r"(活动时间|时间[:：]|日期[:：]|\d{4}年|\d{1,2}月\d{1,2}日|"
    r"\d{1,2}[:：]\d{2}\s*[-—~～]\s*\d{1,2}[:：]\d{2}|"
    r"\d{1,2}\.\d{1,2}\s*[-—~～]\s*\d{1,2}(?:\.\d{1,2})?)"
)
SECTION_EXPLICIT_NOISE_WORDS = ["内容：", "内容:", "介绍：", "介绍:", "通过", "设置", "邀请", "围绕", "打造", "形成"]
ACTIVITY_CONTENT_START_TITLES = {
    "活动内容",
    "活动内容规划",
    "主活动内容",
    "分项活动",
    "现场活动",
    "核心活动",
    "配套活动",
    "辅助内容",
    "互动体验",
    "展览内容",
    "展区规划",
    "空间规划",
    "区域规划",
    "活动形式",
    "活动安排",
    "延展活动",
    "执行内容",
    "体验内容",
    "四感共生",
    "花间有声",
    "月下有色",
    "集间有味",
    "游中有戏",
}
ACTIVITY_CONTENT_STOP_TITLES = {
    "活动背景",
    "项目背景",
    "活动概况",
    "活动目标",
    "活动亮点",
    "活动思想",
    "传播与运营策略",
    "传播策略",
    "运营策略",
    "宣传排期",
    "宣传规划",
    "预期成效",
    "项目作用",
    "活动保障",
    "应急预案",
    "海南非遗引入",
    "名录",
    "组织架构",
    "感谢观看",
    "目录",
    "CONTENTS",
}
SECTION_FORBIDDEN_TITLES = {
    "核心话题矩阵构建",
    "“小红书式”内容种草",
    "小红书式内容种草",
    "官方“整活”与互动",
    "官方整活与互动",
    "游戏化动线设计",
    "社群化用户运营",
    "数据化效果评估",
    "预热期",
    "爆发期",
    "长尾期",
    "宣传片",
    "话题互动招募",
    "海报及图文推送",
    "现场打卡传播",
    "媒体内容集中发布",
    "现场图片直播",
    "活动回顾推文",
    "现场宣传片",
    "项目作用",
    "预期成效",
    "文化传播",
    "品牌影响",
    "海南非遗引入",
    "联合国级",
    "国家级非遗",
    "省级非遗",
    "市县级非遗",
}
SECTION_FORBIDDEN_WORDS = [
    "构建",
    "设计",
    "运营",
    "评估",
    "推送",
    "发布",
    "招募",
    "宣传",
    "策略",
    "目标",
    "作用",
    "成效",
    "维度",
    "模型",
    "品牌",
    "内容",
    "传播",
]
SECTION_STATEMENT_PREFIXES = ("以", "通过", "围绕", "形成", "提升", "增强", "拓展", "丰富")
SECTION_CHILD_MODULES = {
    "味觉实验室",
    "新生造物局",
    "巧手研习社",
    "夕阳·慢调",
    "入夜·新潮",
    "手艺护照通关挑战",
    "手艺主理人分享会",
}
MAIN_SECTION_TITLES = {
    "开幕式",
    "闭幕式",
    "活动大巡游",
    "蚝王争霸赛",
    "食速挑战赛",
    "渡口星光音乐会",
    "生产者大会",
    "生产者大会（市集）",
    "蚝趣游乐记",
    "你好，海洋！艺术展",
    "海边民族联欢会",
    "民族服饰走秀",
    "海岛服饰走秀",
    "特色长桌宴",
    "民族长桌宴",
    "老爸茶长桌宴",
    "三三有礼市集",
    "黎韵绣面",
    "民族服饰体验",
    "民族风写真体验",
    "民族趣味运动会",
    "中华民族共享的文化符号装置",
    "非遗生活市集",
    "寻味东坡夜游路线",
    "装置互动区",
    "手艺·海岸集",
    "手艺·声浪场",
    "手艺·共生境",
    "手艺·传承礼",
    "手艺启元礼.启动仪式",
    "清影花月音乐会",
    "调声之夜",
    "东坡夜话",
    "玉蕊夜绽赏花荟",
    "月下玉蕊·灯光艺术展",
    "十二花色打卡游",
    "花下食集",
    "玉蕊花宴",
    "花间茶寮",
    "花使巡游",
    "玉蕊寻芳游",
    "花笺留诗",
    "夜绘玉蕊",
    "入梦签到",
    "奇幻故事共读",
    "梦醒留念",
    "书海寻宝",
    "白日梦阅读记",
    "诗意阶梯",
    "候船微阅读角",
    "有声阅读二维码",
    "双语阅读角",
    "阅读心得墙",
    "图文展",
    "桌面展",
    "标本展",
    "环境装置",
    "互动展",
    "图书阅读区",
}
SUB_SECTION_TITLES = {
    "签到仪式",
    "开场舞",
    "启动仪式",
    "主持推荐",
    "领导致辞",
    "授牌仪式",
    "合影留念",
    *SECTION_CHILD_MODULES,
    "套蚝赢趣",
    "钓壳寻鲜",
    "贝壳绘梦",
    "曲口寻踪",
    "赶猪赛跑",
    "巧夹槟榔",
    "扁担挑椰子",
    "拉乌龟",
    "欢乐背背跑",
}
PROMO_SECTION_TERMS = {
    "宣传",
    "视频",
    "祝福视频",
    "宣传片",
    "短视频",
    "快剪",
    "公众号软文",
    "软文",
    "推文",
    "海报",
    "倒计时海报",
    "媒体",
    "直播",
    "图片直播",
    "话题",
    "KOL",
    "达人",
    "发布",
    "传播",
    "预热",
    "爆发期",
    "长尾期",
    "前期宣传",
    "中期引爆",
    "长尾延续",
    "线上宣传",
}
BUILD_SECTION_TERMS = {
    "示意图",
    "装置",
    "互动装置",
    "打卡点",
    "打卡装置",
    "立体字",
    "展板",
    "知识展板",
    "文化符号装置",
    "民族团结装置",
    "指引",
    "趣味指引",
    "导视",
    "地图",
    "主画面",
    "背景板",
    "舞台示意",
    "市集示意",
    "氛围布置",
    "灯光艺术装置",
}
BUILD_ACTIVITY_EXCEPTIONS = {"中华民族共享的文化符号装置", "装置互动区", "环境装置"}
LOCATION_SECTION_TERMS = {
    "海口·FUNBAY",
    "FUNBAY",
    "自在湾",
    "活动地点",
    "地点",
    "主会场",
    "分会场",
    "停车位",
    "服务中心",
    "场地",
    "跑步步道",
    "公路",
    "台阶",
    "铺面",
    "图书馆8楼",
    "海口图书馆8楼",
}
PROCESS_SECTION_TERMS = {
    "前期筹备",
    "筹备",
    "活动排期",
    "排期大纲",
    "流程",
    "流程安排",
    "领取规则",
    "参与规则",
    "活动规则",
    "规则",
    "攻略",
    "玩法介绍",
    "活动现场参与规则",
    "项目作用",
    "预期成效",
    "活动目标",
}
MATERIAL_SECTION_TERMS = {
    "蚝仔周边",
    "通行证",
    "工作证",
    "集章卡",
    "印章",
    "礼品",
    "奖品",
    "手册",
    "桌牌",
    "贴纸",
    "明信片",
    "环保袋",
    "徽章",
    "帆布袋",
    "周边",
}
ACTIVITY_FORM_TERMS = {
    "音乐会",
    "演唱",
    "演出",
    "走秀",
    "巡游",
    "市集",
    "食集",
    "长桌宴",
    "茶席",
    "茶寮",
    "游园会",
    "运动会",
    "挑战赛",
    "争霸赛",
    "体验",
    "写真体验",
    "服饰体验",
    "绣面",
    "拓印",
    "手作",
    "夜游路线",
    "艺术展",
    "展",
    "展区",
    "阅读区",
    "秀",
    "推介会",
    "路演",
    "签约仪式",
}
PROMOTION_ACTION_TITLES = {
    "宣传片",
    "话题互动招募",
    "海报及图文推送",
    "现场打卡传播",
    "媒体内容集中发布",
    "现场图片直播",
    "活动回顾推文",
    "现场宣传片",
    "发起线上互动话题",
}
PROMOTION_SENTENCE_PREFIXES = ("发起线上互动话题", "推出倒计时海报", "实时图片直播", "现场视频记录")
CONTINUOUS_SHORT_TITLE_KEYWORDS = [
    "演唱",
    "演出",
    "音乐会",
    "夜话",
    "调声",
    "赏花荟",
    "打卡游",
    "灯光艺术展",
    "花灯展",
    "食集",
    "茶寮",
    "花宴",
    "巡游",
    "留诗",
    "游园",
    "手作",
    "夜绘",
    "市集",
    "体验",
    "互动",
    "艺术展",
    "挑战赛",
    "争霸赛",
    "开幕式",
    "闭幕式",
    "展",
    "展区",
]
SHORT_TITLE_BLOCK_WORDS = ["通过", "设置", "打造", "介绍", "邀请", "引导", "呈现", "增强", "形成", "提供", "进行", "参与", "围绕", "结合"]
NOISE_INFO_TERMS = {
    "活动地点",
    "活动名称",
    "主办单位",
    "承办单位",
    "协办单位",
    "指导单位",
    "参与对象",
    "预计参与人数",
    "项目背景",
    "活动背景",
    "活动目标",
    "活动宗旨",
    "品牌定位",
    "传播目标",
    "经济目标",
    "文化目标",
    "感谢观看",
    "目录",
    "CONTENTS",
    "PART",
}
NOISE_SENTENCE_TERMS = [
    "通过",
    "为了",
    "在",
    "让",
    "将",
    "以",
    "围绕",
    "进行",
    "呈现",
    "打造",
    "提升",
    "引导",
    "增强",
    "安排",
]
GENERIC_ACTIVITY_TERMS = {
    "活动",
    "内容",
    "体验",
    "互动",
    "宣传",
    "保障",
    "流程",
    "安排",
    "大纲",
    "现场",
    "核心",
    "辅助内容",
    "配套活动",
    "核心活动",
    "活动内容",
    "活动内容规划",
    "展览内容",
    "展区规划",
    "空间规划",
    "区域规划",
    "执行内容",
    "体验内容",
    "活动排期大纲",
    "市集",
    "美食",
    "巡游",
    "展览",
    "音乐会",
    "长桌宴",
    "走秀",
    "打卡",
    "阅读",
    "茶席",
    "食集",
    "装置",
    "游园",
    "项互动体验",
    "市集布置",
    "市集氛围布置",
    "市集摊位",
    "帐篷市集",
    "美食摊位",
    "市集活动",
    "美食市集",
    "巡游活动",
    "音乐会",
    "艺术展",
    "灯光艺术展",
    "活动宣传",
    "预热宣传",
    "活动保障",
    "应急保障",
    "人员保障",
    "市集招募",
    "打卡活动",
    "赛事活动",
    "现场活动",
    "主舞台活动",
    "互动活动体验",
}
CONTAINER_SECTION_TITLES = {
    "活动内容",
    "活动内容规划",
    "分项活动",
    "现场活动",
    "核心活动",
    "配套活动",
    "辅助内容",
    "活动安排",
    "活动流程",
    "活动排期",
    "活动框架",
    "活动亮点",
    "互动体验",
    "活动形式",
    "展览内容",
    "展区规划",
    "空间规划",
    "区域规划",
    "执行内容",
    "体验内容",
    "整体规划",
    "内容规划",
    "宣传规划",
    "活动保障",
    "四感共生",
    "花间有声",
    "月下有色",
    "集间有味",
    "游中有戏",
}
SECTION_TRAILING_NOISE_WORDS = [
    "活动时间",
    "活动地点",
    "拟定",
    "招募",
    "内容",
    "场地",
    "舞台",
    "现场",
    "时间",
    "地点",
    "通过",
    "在",
    "中",
    "每",
    "嘉",
    "停",
    "利面",
    "体验动",
    "内游",
]
SECTION_NOISE_SENTENCE_WORDS = [
    "拟定",
    "招募",
    "预计",
    "设置",
    "铺满",
    "通过",
    "引导",
    "展示",
    "介绍",
    "邀请",
    "活动现场",
    "大众",
    "游客",
    "进行",
    "参与",
    "了解",
    "体验",
    "同时",
    "不仅",
    "还能",
    "形成",
]
SECTION_ALIAS_MAP = [
    {
        "name": "海边民族联欢会",
        "include": ["海边民族联欢会", "民族联欢会", "全民嗨曲演唱", "民族金曲演唱", "民族器乐演绎"],
        "exclude": [],
        "source_terms": ["海边民族联欢会", "民族联欢会", "全民嗨曲演唱", "民族金曲演唱", "民族器乐演绎"],
    },
    {
        "name": "民族服饰走秀",
        "include": ["民族服饰走秀", "民族服饰"],
        "exclude": ["体验"],
        "source_terms": ["民族服饰走秀", "民族服饰"],
    },
    {
        "name": "海岛服饰走秀",
        "include": ["海岛服饰走秀", "海岛服饰"],
        "exclude": ["体验"],
        "source_terms": ["海岛服饰走秀", "海岛服饰"],
    },
    {
        "name": "民族长桌宴",
        "include": ["民族长桌宴"],
        "exclude": [],
        "source_terms": ["民族长桌宴"],
    },
    {
        "name": "老爸茶长桌宴",
        "include": ["老爸茶长桌宴"],
        "exclude": [],
        "source_terms": ["老爸茶长桌宴"],
    },
    {
        "name": "中华民族共享的文化符号装置",
        "include": ["中华民族共享的文化符号装置", "文化符号装置", "文化符号知识展板", "中华民族共享的文化符号知识展板"],
        "exclude": [],
        "source_terms": ["中华民族共享的文化符号装置", "文化符号装置", "文化符号知识展板", "中华民族共享的文化符号知识展板"],
    },
    {
        "name": "民族知识知多少",
        "include": ["民族知识知多少", "趣味问答", "答题转盘"],
        "exclude": [],
        "source_terms": ["民族知识知多少", "趣味问答", "答题转盘"],
    },
    {
        "name": "三三有礼市集",
        "include": ["三三有礼市集"],
        "exclude": [],
        "source_terms": ["三三有礼市集"],
    },
    {
        "name": "地摊市集",
        "include": ["地摊市集"],
        "exclude": [],
        "source_terms": ["地摊市集"],
    },
    {
        "name": "黎韵绣面",
        "include": ["黎韵绣面", "绣面"],
        "exclude": [],
        "source_terms": ["黎韵绣面", "绣面"],
    },
    {
        "name": "民族服饰体验",
        "include": ["民族服饰体验", "换装体验", "换装区"],
        "exclude": [],
        "source_terms": ["民族服饰体验", "换装体验", "换装区"],
    },
    {
        "name": "民族风写真体验",
        "include": ["民族风写真体验", "写真体验", "专业写真拍摄"],
        "exclude": [],
        "source_terms": ["民族风写真体验", "写真体验", "专业写真拍摄"],
    },
    {
        "name": "民族趣味运动会",
        "include": ["民族趣味运动会", "趣味民族运动会", "赶猪赛跑", "巧夹槟榔", "扁担挑椰子", "拉乌龟", "欢乐背背跑"],
        "exclude": [],
        "source_terms": ["民族趣味运动会", "趣味民族运动会", "赶猪赛跑", "巧夹槟榔", "扁担挑椰子", "拉乌龟", "欢乐背背跑"],
    },
    {
        "name": "生产者大会（市集）",
        "include": ["生产者大会（市集）", "生产者大会"],
        "exclude": [],
        "source_terms": ["生产者大会（市集）", "生产者大会"],
    },
    {
        "name": "花下食集",
        "include": ["花下食集"],
        "exclude": [],
        "source_terms": ["花下食集"],
    },
    {
        "name": "茶肆新春市集",
        "include": ["茶肆新春市集"],
        "exclude": [],
        "source_terms": ["茶肆新春市集"],
    },
    {
        "name": "湾畔茶席",
        "include": ["湾畔茶席"],
        "exclude": [],
        "source_terms": ["湾畔茶席"],
    },
    {
        "name": "你好，海洋！艺术展",
        "include": ["你好，海洋！艺术展", "你好 海洋 艺术展", "你好海洋艺术展"],
        "exclude": [],
        "source_terms": ["你好，海洋！艺术展", "你好 海洋 艺术展", "你好海洋艺术展"],
    },
    {
        "name": "清影花月音乐会",
        "include": ["清影花月音乐会"],
        "exclude": [],
        "source_terms": ["清影花月音乐会"],
    },
    {
        "name": "调声之夜",
        "include": ["调声之夜"],
        "exclude": [],
        "source_terms": ["调声之夜"],
    },
    {
        "name": "东坡夜话",
        "include": ["东坡夜话"],
        "exclude": [],
        "source_terms": ["东坡夜话"],
    },
    {
        "name": "玉蕊夜绽赏花荟",
        "include": ["玉蕊夜绽赏花荟"],
        "exclude": [],
        "source_terms": ["玉蕊夜绽赏花荟"],
    },
    {
        "name": "月下玉蕊·灯光艺术展",
        "include": ["月下玉蕊·灯光艺术展", "月下玉蕊·花灯展"],
        "exclude": [],
        "source_terms": ["月下玉蕊·灯光艺术展", "月下玉蕊·花灯展"],
    },
    {
        "name": "十二花色打卡游",
        "include": ["十二花色打卡游"],
        "exclude": [],
        "source_terms": ["十二花色打卡游"],
    },
    {
        "name": "玉蕊花宴",
        "include": ["玉蕊花宴"],
        "exclude": [],
        "source_terms": ["玉蕊花宴"],
    },
    {
        "name": "花间茶寮",
        "include": ["花间茶寮"],
        "exclude": [],
        "source_terms": ["花间茶寮"],
    },
    {
        "name": "花使巡游",
        "include": ["花使巡游"],
        "exclude": [],
        "source_terms": ["花使巡游"],
    },
    {
        "name": "玉蕊寻芳游",
        "include": ["玉蕊寻芳游", "玉蕊寻芳游园会"],
        "exclude": [],
        "source_terms": ["玉蕊寻芳游", "玉蕊寻芳游园会"],
    },
    {
        "name": "花笺留诗",
        "include": ["花笺留诗"],
        "exclude": [],
        "source_terms": ["花笺留诗"],
    },
    {
        "name": "夜绘玉蕊",
        "include": ["夜绘玉蕊"],
        "exclude": [],
        "source_terms": ["夜绘玉蕊"],
    },
]
PUBLIC_BEAUTY_BUILD_KEYWORDS = [
    "主入口装置",
    "主视觉装置",
    "大型打卡装置",
    "立体字打卡装置",
    "主题美陈",
    "大型美陈装置",
    "全场氛围布置",
    "现场氛围布置",
    "灯光艺术装置",
    "花灯装置",
    "花艺花境",
    "主题装置",
    "互动装置",
    "民族团结互动装置",
    "环保回收装置",
    "蓝色守护舱",
    "花境",
    "花瀑",
    "河灯装置",
    "夜游灯光装置",
    "装置艺术",
    "美陈装置",
    "打卡装置",
    "氛围布置",
]
PUBLIC_BUILD_KEYWORDS = [
    "舞台",
    "主舞台",
    "舞台搭建",
    "舞台背景",
    "背景板",
    "主舞台背景",
    "签到墙",
    "签到背景墙",
    "展板",
    "图文展板",
    "知识展板",
    "主题展板",
    "艺术展陈",
    "展架",
    "指引牌",
    "导视牌",
    "区域指引",
    "趣味指引",
    "入口指引",
    "活动地图墙",
    "服务台",
    "信息咨询台",
    "签到台",
    "领取点",
    "铁马",
    "围挡",
    "警戒带",
    "安全线",
    "临时搭建",
    "基础搭建",
    "公共桌椅布置",
    "公共休息区布置",
    "舞台视觉包装",
    "导视指引",
    "签到墙",
    "铁马围挡",
]
PEOPLE_OR_OTHER_KEYWORDS = [
    "主持人",
    "执行人员",
    "工作人员",
    "现场执行",
    "点位人员",
    "志愿者",
    "安保人员",
    "保洁人员",
    "后勤保障人员",
    "搭建撤场工人",
    "技术人员",
    "音响师",
    "灯光师",
    "摄影师",
    "摄像师",
    "医疗人员",
    "救生员",
    "交通引导人员",
    "停车引导人员",
    "交通停车引导",
    "医疗保障",
    "礼仪人员",
    "体验老师",
    "非遗传承人",
    "NPC演员",
    "演员",
    "模特",
    "运输费",
    "物料运输",
    "物流运输",
    "餐费",
    "工作餐",
    "饮用水",
    "矿泉水",
    "对讲机",
    "应急物资",
    "急救包",
    "灭火器",
    "雨衣",
    "防滑垫",
    "电力保障",
    "临时用电",
    "配电",
    "备用电源",
    "保险",
    "不可预见费",
    "管理费",
    "税费",
    "工作证",
]
ITEM_SECTION_INFERENCE = {
    "节目演出": "演艺活动",
    "趣味互动游戏": "互动体验",
    "赛事活动": "赛事活动",
    "长桌宴": "长桌宴",
    "茶席茶寮": "茶席茶寮",
    "帐篷摊位": "市集活动",
    "阅读区布置": "阅读活动",
    "图书配置": "阅读活动",
    "手作体验材料": "手作体验",
    "非遗拓印材料": "非遗体验",
    "非遗体验材料": "非遗体验",
    "灯光艺术装置": "美陈搭建类",
    "美陈装置": "美陈搭建类",
    "打卡装置": "美陈搭建类",
    "打卡点位": "美陈搭建类",
    "诗句展示装置": "美陈搭建类",
    "展板": "其他搭建类",
    "指引牌": "其他搭建类",
    "导视牌": "其他搭建类",
    "导视指引": "其他搭建类",
    "签到墙": "其他搭建类",
    "音频二维码点位": "阅读活动",
    "视频快剪": "活动宣传",
    "公众号软文": "活动宣传",
    "倒计时海报": "活动宣传",
    "话题传播": "活动宣传",
    "摄影摄像服务": "活动宣传",
    "传播图文": "活动宣传",
    "短视频内容": "活动宣传",
}

MODULE_SECTION_PRIORITY_ITEMS = {
    "节目演出",
    "灯光音响套装",
    "舞台搭建",
    "舞台视觉包装",
    "启动仪式道具",
    "签到墙",
    "签到物料",
    "主持人",
    "摄影摄像服务",
    "技术人员",
    "演艺服化道",
    "茶席茶寮",
    "花艺花境",
    "灯光艺术装置",
    "导视指引",
    "打卡装置",
    "美陈装置",
    "展板",
    "艺术展陈",
    "电力保障",
    "互动规则牌",
    "文创礼品",
    "帐篷摊位",
    "餐饮体验",
    "市集招募服务",
    "摊位视觉物料",
    "长桌宴",
    "桌椅布置",
    "氛围布置",
    "NPC互动服务",
    "趣味互动游戏",
    "通行证",
    "印章",
    "活动贴纸",
    "故事画面墙",
    "手作体验材料",
    "体验老师",
    "互动说明牌",
    "活动桌椅",
    "作品展示材料",
    "阅读区布置",
    "图书配置",
    "音频二维码点位",
}


def _resolve_item_column(df: pd.DataFrame) -> str | None:
    for column in ("标准项目", "项目", "建议补充项"):
        if column in df.columns:
            return column
    return None


def _matches_keywords(item_name: str, keywords: list[str]) -> bool:
    return any(keyword in item_name for keyword in keywords if keyword)


def _safe_first_match_position(value: Any) -> int | None:
    if isinstance(value, list):
        for item in value:
            try:
                return int(item)
            except (TypeError, ValueError):
                continue
        return None
    try:
        return int(value)
    except (TypeError, ValueError):
        return None


def _normalize_item_name(item_name: str) -> str:
    item_name = str(item_name or "").strip()
    if item_name in {"市集摊位", "帐篷摊位"}:
        return "帐篷摊位"
    if item_name in {
        "开场表演",
        "节目表演",
        "乐队演绎",
        "舞蹈演绎",
        "脱口秀演绎",
        "巡游表演",
        "非遗演绎",
        "服饰走秀",
        "音乐会",
        "琼剧表演",
    }:
        return "节目演出"
    if item_name in {"工作牌", "工作卡", "胸卡", "人员证件"}:
        return "工作证"
    if item_name in {
        "主视觉设计",
        "KV设计",
        "活动海报设计",
        "倒计时海报设计",
        "延展设计",
        "物料画面设计",
        "导视系统设计",
        "舞台视觉设计",
        "摊位视觉设计",
        "IP形象设计",
        "品牌视觉设计",
        "宣传物料设计",
        "活动视觉设计",
        "美陈画面设计",
        "展板画面设计",
        "背景板画面设计",
    }:
        return "设计服务"
    return item_name


def normalize_category(category: str, item_name: str = "") -> str:
    item_name = _normalize_item_name(item_name)
    category = str(category or "").strip()

    if _matches_keywords(item_name, SCENE_KEYWORDS):
        return "场景搭建类"
    if _matches_keywords(item_name, MATERIAL_KEYWORDS):
        return "活动物料类"
    if _matches_keywords(item_name, CONTENT_KEYWORDS):
        return "活动内容类"
    if _matches_keywords(item_name, EQUIPMENT_KEYWORDS):
        return "设备租赁类"
    if _matches_keywords(item_name, PERSONNEL_KEYWORDS):
        return "人员服务类"
    if _matches_keywords(item_name, LOGISTICS_KEYWORDS):
        return "后勤保障类"
    if _matches_keywords(item_name, PROMOTION_KEYWORDS):
        return "宣传推广类"

    mapped_category = CATEGORY_MAP.get(category)
    if mapped_category:
        return mapped_category
    return "后勤保障类"


def normalize_quote_categories(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty or "项目分类" not in df.columns:
        return df.copy()

    normalized_df = df.copy()
    item_column = _resolve_item_column(normalized_df)
    if item_column is None:
        normalized_df["项目分类"] = normalized_df["项目分类"].map(lambda value: normalize_category(str(value), ""))
        return normalized_df

    normalized_df["项目分类"] = normalized_df.apply(
        lambda row: normalize_category(row.get("项目分类", ""), row.get(item_column, "")),
        axis=1,
    )
    return normalized_df


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

    if any(term in filter_source for term in LOCATION_SECTION_TERMS):
        return "location"
    if any(term in filter_source for term in PROMO_SECTION_TERMS):
        return "promo"
    if any(term in filter_source for term in PROCESS_SECTION_TERMS):
        return "process"
    if any(term in filter_source for term in MATERIAL_SECTION_TERMS):
        return "material"
    if cleaned in CONTAINER_SECTION_TITLES or cleaned in GENERIC_ACTIVITY_TERMS or cleaned in SECTION_IGNORE_TITLES:
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
    if cleaned in ACTIVITY_SECTION_CANDIDATES and cleaned not in GENERIC_ACTIVITY_TERMS:
        return "activity"
    if any(term in cleaned for term in ACTIVITY_FORM_TERMS):
        return "activity"
    structured_context = (
        "活动内容区域识别" in source
        or "编号活动标题" in source
        or "区域标题" in source
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
    if cleaned in CONTAINER_SECTION_TITLES or cleaned in GENERIC_ACTIVITY_TERMS or cleaned in SECTION_IGNORE_TITLES:
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
    if cleaned in ACTIVITY_SECTION_CANDIDATES and cleaned not in GENERIC_ACTIVITY_TERMS:
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
    if any(word in cleaned for word in SECTION_FORBIDDEN_WORDS):
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
            if reason in {"板块标题", "延展活动标题", "展区标题", "区域标题"}:
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
        if candidate not in GENERIC_ACTIVITY_TERMS and candidate not in CONTAINER_SECTION_TITLES and len(candidate) >= 4
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

    if cleaned in GENERIC_ACTIVITY_TERMS or cleaned in SECTION_IGNORE_TITLES:
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
    if text in GENERIC_ACTIVITY_TERMS:
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

    if cleaned in GENERIC_ACTIVITY_TERMS:
        return "noise", "通用框架词"

    if _starts_with_any(cleaned, SUB_SECTION_TITLES):
        return "candidate", "子活动/流程环节"

    alias_hit = _map_section_alias(cleaned)
    if alias_hit:
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
        return inferred_section

    if matched_section:
        return str(matched_section["name"])

    return "未归属板块"


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


def _append_note(existing: str, note: str) -> str:
    parts = [part for part in str(existing or "").split("；") if part]
    for note_part in [part for part in str(note or "").split("；") if part]:
        if note_part not in parts:
            parts.append(note_part)
    return "；".join(parts)


def _merge_text_values(left: Any, right: Any, separator: str = "、") -> str:
    values: list[str] = []
    for raw_value in (left, right):
        raw_parts = raw_value if isinstance(raw_value, list) else str(raw_value or "").split(separator)
        for part in raw_parts:
            value = str(part).strip()
            if value and value not in values:
                values.append(value)
    return separator.join(values)


def _merge_list_values(left: Any, right: Any) -> list[Any]:
    values: list[Any] = []
    for raw_value in (left, right):
        raw_items = raw_value if isinstance(raw_value, list) else [raw_value]
        for item in raw_items:
            if item in ("", None):
                continue
            if item not in values:
                values.append(item)
    return values


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


def sort_quote_items(df: pd.DataFrame) -> pd.DataFrame:
    """Normalize final categories and keep same-category items contiguous."""
    if df.empty or "项目分类" not in df.columns:
        return df.copy()

    sorted_df = normalize_quote_categories(df)
    category_order = {category: index for index, category in enumerate(CATEGORY_ORDER)}
    item_column = _resolve_item_column(sorted_df)

    sorted_df["_category_rank"] = sorted_df["项目分类"].map(lambda value: category_order.get(str(value), len(CATEGORY_ORDER)))
    sorted_df["_item_sort"] = sorted_df[item_column].astype(str) if item_column else ""
    sorted_df = sorted_df.sort_values(
        by=["_category_rank", "项目分类", "_item_sort"],
        kind="stable",
    ).drop(columns=["_category_rank", "_item_sort"])
    return sorted_df


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


def _is_empty(value: Any) -> bool:
    return pd.isna(value) or value == ""


def _to_number(value: Any) -> float | None:
    if _is_empty(value):
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


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

        built_rows.append(row)

    df = recalculate_totals(pd.DataFrame(built_rows, columns=QUOTE_COLUMNS))
    return finalize_quote_rows(df.to_dict("records"), source_text, activity_sections=activity_sections)
