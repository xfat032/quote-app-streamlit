"""Extract standard quote items from plan text."""

from __future__ import annotations

import re
from collections import OrderedDict
from datetime import date
from typing import Any

from .normalizer import find_alias_matches


NUMBER_UNIT_RE = re.compile(r"(?<![\d.])(\d+(?:\.\d+)?)\s*(个|顶|盏|名|人|台|套|场|天|项|份|箱|瓶|辆|张|批|篇|条|家|米|册)")
ATTACHED_GAP_RE = re.compile(r"^[\s,，、：:；;约共计为]*$")
BOOK_COUNT_RE = re.compile(r"(?:提供)?约?\s*(\d+(?:\.\d+)?)\s*册图书")
INTERACTIVE_PROJECT_RE = re.compile(r"(\d+(?:\.\d+)?)\s*个互动项目")
STAMP_COUNT_RE = re.compile(r"集满\s*(\d+(?:\.\d+)?)\s*枚印章")
DATE_RANGE_RE = re.compile(r"(\d{1,2})月(\d{1,2})日\s*[—\-－~～至到]+\s*(?:(\d{1,2})月)?(\d{1,2})日")
DOT_DATE_RANGE_RE = re.compile(r"(\d{1,2})\.(\d{1,2})\s*[—\-－~～至到]+\s*(?:(\d{1,2})\.)?(\d{1,2})")
MARKET_STALL_RE = re.compile(r"(?:招募|设置|计划设置|配置|提供)?\s*(\d+(?:\.\d+)?)\s*个\s*(?:市集摊位|帐篷摊位|摊位|帐篷)")
MEAL_QUANTITY_RE = re.compile(r"(?:预计|每天设有|限额|提供)?\s*(\d+(?:\.\d+)?)\s*(?:份|人)")
GAME_SESSION_RE = re.compile(r"(?:每天)?\s*(\d+(?:\.\d+)?)\s*场比赛|共计\s*(\d+(?:\.\d+)?)\s*场")
GAME_PLAYER_RE = re.compile(r"每场\s*(\d+(?:\.\d+)?)\s*位参赛者")
VIDEO_DURATION_RE = re.compile(r"(\d+(?:\.\d+)?)\s*秒")
KOL_COUNT_RE = re.compile(r"(?:网红博主|KOL|达人)\s*(\d+(?:\.\d+)?)\s*人", re.IGNORECASE)
BRAND_COUNT_RE = re.compile(r"联动品牌\s*(\d+(?:\.\d+)?)\s*余?家")
PERSONNEL_COUNT_PATTERNS = {
    "安保人员": [re.compile(r"安保(?:人员)?[（(]\s*(\d+(?:\.\d+)?)\s*人[）)]")],
    "志愿者": [re.compile(r"志愿者[（(]\s*(\d+(?:\.\d+)?)\s*人[）)]")],
    "后勤保障人员": [re.compile(r"后勤保障[（(]\s*(\d+(?:\.\d+)?)\s*人[）)]")],
    "技术人员": [re.compile(r"舞台与灯光音响团队[（(]\s*(\d+(?:\.\d+)?)\s*人[）)]"), re.compile(r"技术人员[（(]\s*(\d+(?:\.\d+)?)\s*人[）)]")],
    "医疗保障": [re.compile(r"医疗保障[（(]\s*(\d+(?:\.\d+)?)\s*人[）)]")],
    "交通停车引导": [re.compile(r"交通与停车引导[（(]\s*(\d+(?:\.\d+)?)\s*人[）)]"), re.compile(r"停车引导[（(]\s*(\d+(?:\.\d+)?)\s*人[）)]")],
}
LONG_TABLE_ALLOWED_KEYWORDS = {
    "长桌宴",
    "花宴",
    "玉蕊花宴",
    "民族长桌宴",
    "老爸茶长桌宴",
    "长桌雅宴",
    "共宴",
    "餐券",
    "围炉共宴",
    "长桌形式",
}
TEA_SEAT_ALLOWED_KEYWORDS = {"茶席", "茶寮", "围炉煮茶", "茶饮空间"}
TEA_ONLY_FOOD_ALIASES = {"围炉煮茶"}
DINING_DIRECT_TERMS = {"餐饮体验", "餐饮服务", "食品留样"}
DINING_EXECUTION_TERMS = {
    "提供",
    "供应",
    "售卖",
    "销售",
    "餐饮品牌",
    "餐饮区",
    "就餐区",
    "餐标",
    "餐券",
    "餐食",
    "餐品",
    "用餐",
    "就餐",
}
DINING_STRONG_EXECUTION_TERMS = {
    "餐饮体验",
    "餐饮服务",
    "提供餐饮",
    "供应餐食",
    "售卖餐饮",
    "餐标",
    "餐券",
    "用餐",
    "就餐",
    "食品留样",
}
DINING_PRODUCT_CONTEXT_TERMS = {
    "制作体验",
    "制作环节",
    "点心制作",
    "器具",
    "视觉来源",
    "文创产品",
    "文创元素",
    "纹样",
    "贴纸",
    "徽章",
    "冰箱贴",
    "手机挂件",
    "杯垫",
    "明信片",
    "帆布袋",
    "礼盒",
    "伴手礼",
    "会客厅",
    "习俗",
}
TEA_BREAK_EXECUTION_TERMS = {
    "茶歇",
    "茶歇服务",
    "课间",
    "休息区",
    "会议",
    "培训",
    "提供",
    "配置",
    "安排",
    "咖啡",
    "水果",
}
TEA_BREAK_PRODUCT_CONTEXT_TERMS = {
    "礼盒",
    "伴手礼",
    "摊位",
    "美食摊位",
    "传统糕点",
    "老爸茶点心",
    "点心制作",
    "制作体验",
    "产品",
}
DIMENSION_PATTERNS = [
    re.compile(r"(\d+(?:\.\d+)?)\s*(m|米)\s*[×xX*]\s*(\d+(?:\.\d+)?)\s*(m|米)"),
    re.compile(r"(?:长|长度)\s*(\d+(?:\.\d+)?)\s*(m|米)[，,、\s]*(?:宽|宽度|纵深|深度|深)\s*(\d+(?:\.\d+)?)\s*(m|米)"),
]
HEIGHT_RE = re.compile(r"(?:高|高度)\s*(\d+(?:\.\d+)?)\s*(m|米)")


MODULE_COMPLETION_RULES: list[tuple[list[str], list[str]]] = [
    (["清影花月音乐会"], ["节目演出", "灯光音响套装", "舞台搭建", "摄影摄像服务", "技术人员"]),
    (["调声之夜"], ["节目演出", "演艺服化道", "灯光音响套装", "摄影摄像服务"]),
    (["东坡夜话"], ["节目演出", "阅读区布置", "茶席茶寮", "灯光音响套装"]),
    (["玉蕊夜绽赏花荟"], ["花艺花境", "灯光艺术装置", "导视指引", "打卡装置"]),
    (["月下玉蕊·灯光艺术展", "月下玉蕊·花灯展"], ["灯光艺术装置", "美陈装置", "导视指引", "电力保障"]),
    (["十二花色打卡游"], ["打卡装置", "互动规则牌", "导视指引", "文创礼品"]),
    (["花下食集"], ["帐篷摊位", "餐饮体验", "市集招募服务", "摊位视觉物料"]),
    (["玉蕊花宴"], ["长桌宴", "餐饮体验", "桌椅布置", "花艺花境"]),
    (["花间茶寮"], ["茶席茶寮", "桌椅布置", "餐饮体验", "氛围布置"]),
    (["花使巡游"], ["节目演出", "演艺服化道", "NPC互动服务", "摄影摄像服务"]),
    (["玉蕊寻芳游", "玉蕊寻芳游园会"], ["NPC互动服务", "趣味互动游戏", "通行证", "印章", "互动规则牌"]),
    (["花笺留诗"], ["互动规则牌", "活动贴纸", "手作体验材料", "故事画面墙"]),
    (["夜绘玉蕊"], ["手作体验材料", "体验老师", "活动桌椅", "互动规则牌"]),
    (["开营破冰", "团队破冰", "破冰分组"], ["拓展教练", "团建道具", "互动规则牌", "摄影摄像服务"]),
    (["团队拓展挑战", "户外拓展挑战", "协作挑战赛", "团队协作挑战"], ["拓展教练", "拓展器材", "团建道具", "互动规则牌", "文创礼品"]),
    (["复盘共创", "团队复盘", "行动计划共创"], ["培训讲师", "课程物料", "培训设备", "执行人员"]),
    (["结营仪式", "颁奖仪式"], ["主持人", "证书奖杯", "摄影摄像服务", "灯光音响套装"]),
    (["开班仪式"], ["签到墙", "签到物料", "主持人", "摄影摄像服务"]),
    (["专题授课", "主题培训", "高效沟通工作坊", "管理者训练营"], ["培训讲师", "课程物料", "培训设备", "会议场地"]),
    (["分组研讨", "案例研讨"], ["培训讲师", "课程物料", "培训设备", "茶歇服务"]),
    (["情景演练", "沙盘推演", "实操演练"], ["培训讲师", "课程物料", "团建道具", "互动规则牌"]),
    (["结业测评", "结业仪式", "成果汇报"], ["培训讲师", "课程物料", "证书奖杯", "摄影摄像服务"]),
    (["入梦签到"], ["签到墙", "签到物料", "活动书签", "导视指引"]),
    (["奇幻故事共读"], ["主持人", "阅读区布置", "摄影摄像服务"]),
    (["梦醒留念"], ["摄影摄像服务"]),
    (["书海寻宝"], ["阅读点位", "文创礼品", "互动规则牌", "执行人员"]),
    (["白日梦阅读记"], ["通行证", "印章", "文创礼品", "互动规则牌", "执行人员"]),
    (["非遗拓印体验"], ["手作体验材料", "体验老师", "打卡装置", "执行人员"]),
    (["一分钟阅读挑战"], ["互动规则牌", "主题卡", "执行人员"]),
    (["阅读盲选快答"], ["主题卡", "互动规则牌", "执行人员"]),
    (["诗意阶梯"], ["美陈装置", "打卡装置"]),
    (["候船微阅读角"], ["阅读区布置", "图书配置", "金句撕页", "执行人员"]),
    (["有声阅读二维码", "音频二维码", "扫码收听"], ["音频二维码点位"]),
    (["双语阅读角"], ["阅读区布置", "图书配置", "志愿者"]),
    (["阅读心得墙"], ["打卡装置", "活动贴纸"]),
    (["传播与互动"], ["公众号软文", "视频快剪", "摄影摄像服务"]),
    (["开幕式", "开场仪式", "启动仪式"], ["舞台搭建", "舞台视觉包装", "灯光音响套装", "主持人", "节目演出", "启动仪式道具", "签到墙", "签到物料", "文创礼品", "摄影摄像服务", "执行人员"]),
    (["市集", "生产者大会", "食集", "花下食集", "三三有礼市集", "美食摊位", "餐饮品牌"], ["帐篷摊位", "餐饮体验", "市集招募服务"]),
    (["地摊市集"], ["地摊市集布置", "市集招募服务", "导视指引", "执行人员", "氛围布置"]),
    (["长桌宴", "花宴", "玉蕊花宴", "民族长桌宴", "老爸茶长桌宴", "长桌雅宴", "共宴", "餐券", "围炉共宴", "长桌形式"], ["长桌宴", "餐饮体验", "签到物料", "桌椅布置", "花艺花境", "氛围布置", "执行人员"]),
    (["美食", "餐饮", "小吃", "饮品", "轻食", "茶饮", "餐饮体验"], ["餐饮体验"]),
    (["茶席", "茶寮", "围炉煮茶", "茶饮空间"], ["茶席茶寮"]),
    (["音乐会", "调声之夜", "东坡夜话", "民族器乐", "演唱"], ["节目演出", "舞台搭建", "灯光音响套装", "主持人", "摄影摄像服务", "技术人员"]),
    (["巡游", "NPC", "游走式表演", "花使巡游"], ["节目演出", "NPC互动服务", "演艺服化道", "导视指引", "执行人员", "摄影摄像服务"]),
    (["游园会", "集章", "通行证", "打卡游", "寻芳游"], ["通行证", "印章", "互动规则牌", "趣味互动游戏", "文创礼品", "执行人员", "导视指引"]),
    (["趣味运动会", "互动游戏", "挑战赛", "争霸赛"], ["趣味互动游戏", "赛事活动", "互动道具", "互动规则牌", "文创礼品", "执行人员"]),
    (["艺术展", "知识展", "展览", "科普展示"], ["艺术展陈", "展板", "导视指引", "灯光艺术装置", "执行人员"]),
    (["灯展", "夜游", "赏花", "花境", "河灯"], ["灯光艺术装置", "花艺花境", "氛围布置", "导视指引", "打卡装置", "安保人员", "电力保障"]),
    (["打卡装置", "互动墙", "心愿墙", "留诗", "投票墙"], ["打卡装置", "活动贴纸", "互动规则牌", "执行人员", "摄影摄像服务"]),
    (["手作体验", "非遗拓印体验", "拓印体验", "香囊体验", "标本制作"], ["手作体验材料", "体验老师", "互动规则牌", "执行人员"]),
    (["阅读区", "阅读角", "候船微阅读角", "双语阅读角", "阅读空间", "图书阅读区"], ["阅读区布置"]),
    (["图书", "书籍", "阅读书目", "双语图书", "200册图书", "提供图书"], ["图书配置"]),
    (["传播规划", "预热宣传", "中期引爆", "长尾延续"], ["预热视频", "公众号软文", "倒计时海报", "话题传播", "图片直播", "视频快剪", "主流媒体宣传", "达人KOL推广"]),
    (["晒照有礼", "用户共创", "线上话题挑战"], ["话题传播", "文创礼品", "达人KOL推广", "摄影摄像服务"]),
    (["活动保障", "应急预案", "安全保障"], ["安保人员", "志愿者", "后勤保障人员", "医疗保障", "交通停车引导", "对讲机", "饮用水", "餐费", "铁马围挡", "应急物资", "信息咨询台", "电力保障"]),
    (["水域安全", "海边", "沿岸", "涨潮"], ["救生员", "铁马围挡", "应急物资", "导视指引", "安保人员"]),
]

STRICT_MODULE_COMPLETION_ITEMS = {
    "开幕式": ["舞台搭建", "舞台视觉包装", "灯光音响套装", "主持人", "节目演出", "启动仪式道具", "签到墙", "签到物料", "摄影摄像服务"],
    "海边民族联欢会": ["节目演出", "灯光音响套装", "舞台搭建", "主持人"],
    "民族服饰走秀": ["节目演出", "演艺服化道", "舞台搭建", "摄影摄像服务"],
    "海岛服饰走秀": ["节目演出", "演艺服化道", "舞台搭建", "摄影摄像服务"],
    "民族长桌宴": ["长桌宴", "餐饮体验", "桌椅布置", "花艺花境"],
    "老爸茶长桌宴": ["长桌宴", "餐饮体验", "桌椅布置", "花艺花境"],
    "特色长桌宴": ["长桌宴", "餐饮体验", "桌椅布置", "花艺花境"],
    "海边民族长桌宴": ["长桌宴", "餐饮体验", "桌椅布置", "花艺花境"],
    "三三有礼市集": ["帐篷摊位", "摊位视觉物料", "市集招募服务", "导视指引"],
    "地摊市集": ["帐篷摊位", "摊位视觉物料", "市集招募服务", "导视指引"],
    "民族趣味运动会": ["趣味互动游戏", "互动道具", "互动规则牌", "文创礼品", "印章"],
    "你好，海洋！艺术展": ["展板", "艺术展陈", "导视指引", "灯光艺术装置"],
    "中华民族共享的文化符号装置": ["展板", "美陈装置", "音频二维码点位", "导视指引"],
    "民族知识知多少": ["趣味互动游戏", "互动道具", "互动规则牌", "文创礼品"],
    "清影花月音乐会": ["节目演出", "灯光音响套装", "舞台搭建", "摄影摄像服务", "技术人员"],
    "调声之夜": ["节目演出", "演艺服化道", "灯光音响套装", "摄影摄像服务"],
    "东坡夜话": ["节目演出", "阅读区布置", "茶席茶寮", "灯光音响套装"],
    "玉蕊夜绽赏花荟": ["花艺花境", "灯光艺术装置", "导视指引", "打卡装置"],
    "月下玉蕊·灯光艺术展": ["灯光艺术装置", "美陈装置", "导视指引", "电力保障"],
    "月下玉蕊·花灯展": ["灯光艺术装置", "美陈装置", "导视指引", "电力保障"],
    "十二花色打卡游": ["打卡装置", "互动规则牌", "导视指引", "文创礼品"],
    "花下食集": ["帐篷摊位", "餐饮体验", "市集招募服务", "摊位视觉物料"],
    "玉蕊花宴": ["长桌宴", "餐饮体验", "桌椅布置", "花艺花境"],
    "花间茶寮": ["茶席茶寮", "桌椅布置", "餐饮体验", "氛围布置"],
    "花使巡游": ["节目演出", "演艺服化道", "NPC互动服务", "摄影摄像服务"],
    "玉蕊寻芳游": ["NPC互动服务", "趣味互动游戏", "通行证", "印章", "互动规则牌"],
    "玉蕊寻芳游园会": ["NPC互动服务", "趣味互动游戏", "通行证", "印章", "互动规则牌"],
    "花笺留诗": ["互动规则牌", "活动贴纸", "手作体验材料", "故事画面墙"],
    "夜绘玉蕊": ["手作体验材料", "体验老师", "活动桌椅", "互动规则牌"],
    "开营破冰": ["拓展教练", "团建道具", "互动规则牌", "摄影摄像服务"],
    "团队破冰": ["拓展教练", "团建道具", "互动规则牌", "摄影摄像服务"],
    "破冰分组": ["拓展教练", "团建道具", "互动规则牌", "摄影摄像服务"],
    "团队拓展挑战": ["拓展教练", "拓展器材", "团建道具", "互动规则牌", "文创礼品"],
    "户外拓展挑战": ["拓展教练", "拓展器材", "团建道具", "互动规则牌", "文创礼品"],
    "协作挑战赛": ["拓展教练", "拓展器材", "团建道具", "互动规则牌", "文创礼品"],
    "团队协作挑战": ["拓展教练", "拓展器材", "团建道具", "互动规则牌", "文创礼品"],
    "复盘共创": ["培训讲师", "课程物料", "培训设备", "执行人员"],
    "团队复盘": ["培训讲师", "课程物料", "培训设备", "执行人员"],
    "行动计划共创": ["培训讲师", "课程物料", "培训设备", "执行人员"],
    "结营仪式": ["主持人", "证书奖杯", "摄影摄像服务", "灯光音响套装"],
    "颁奖仪式": ["主持人", "证书奖杯", "摄影摄像服务", "灯光音响套装"],
    "开班仪式": ["签到墙", "签到物料", "主持人", "摄影摄像服务"],
    "专题授课": ["培训讲师", "课程物料", "培训设备", "会议场地"],
    "主题培训": ["培训讲师", "课程物料", "培训设备", "会议场地"],
    "高效沟通工作坊": ["培训讲师", "课程物料", "培训设备", "会议场地"],
    "管理者训练营": ["培训讲师", "课程物料", "培训设备", "会议场地"],
    "分组研讨": ["培训讲师", "课程物料", "培训设备", "茶歇服务"],
    "案例研讨": ["培训讲师", "课程物料", "培训设备", "茶歇服务"],
    "情景演练": ["培训讲师", "课程物料", "团建道具", "互动规则牌"],
    "沙盘推演": ["培训讲师", "课程物料", "团建道具", "互动规则牌"],
    "实操演练": ["培训讲师", "课程物料", "团建道具", "互动规则牌"],
    "结业测评": ["培训讲师", "课程物料", "证书奖杯", "摄影摄像服务"],
    "结业仪式": ["培训讲师", "课程物料", "证书奖杯", "摄影摄像服务"],
    "成果汇报": ["培训讲师", "课程物料", "证书奖杯", "摄影摄像服务"],
}

GENERIC_MODULE_TRIGGERS = {
    "市集",
    "食集",
    "美食",
    "餐饮",
    "小吃",
    "饮品",
    "轻食",
    "茶饮",
    "音乐会",
    "民族器乐",
    "演唱",
    "巡游",
    "NPC",
    "游走式表演",
    "互动游戏",
    "趣味运动会",
    "挑战赛",
    "争霸赛",
    "艺术展",
    "知识展",
    "展览",
    "科普展示",
    "灯展",
    "夜游",
    "赏花",
    "花境",
    "河灯",
    "阅读角",
    "共读",
    "夜读",
    "有声阅读",
    "手作体验",
    "绘梦",
    "夜绘",
    "拓印",
    "香囊",
    "标本",
}

READING_LAYOUT_ALLOWED_TRIGGERS = {"阅读区", "阅读角", "候船微阅读角", "双语阅读角", "阅读空间", "图书阅读区"}
BOOK_CONFIG_ALLOWED_TRIGGERS = {"图书", "书籍", "阅读书目", "双语图书", "200册图书", "提供图书"}
AUDIO_QR_ALLOWED_TRIGGERS = {"有声阅读二维码", "音频二维码", "扫码收听", "二维码点位", "桌牌二维码"}

READING_LAYOUT_BLOCK_TERMS = {"夜读空间", "诗词吟诵", "夜绘玉蕊", "文化活动", "艺术周", "有声有色"}
BOOK_CONFIG_BLOCK_TERMS = {"夜话", "诗词吟诵", "文化活动", "艺术周", "有声有色"}

EXPLICIT_EVIDENCE_TYPES = {"explicit_text", "module_completion", "user_selected_suggestion"}

CANDIDATE_KEYWORDS = [
    "活动", "体验", "装置", "展", "墙", "区", "点位", "市集", "巡游", "音乐会", "宴", "挑战",
    "比赛", "话题", "海报", "视频", "直播", "推文", "保障", "人员", "物料", "道具", "礼品",
]


def _format_number(value: str) -> int | float:
    number = float(value)
    return int(number) if number.is_integer() else number


def _context(text: str, start: int, end: int, before: int = 20, after: int = 40) -> str:
    return text[max(0, start - before) : min(len(text), end + after)]


def _is_near_alias(context: str, alias: str, match: re.Match[str], max_gap: int = 8) -> bool:
    alias_start = context.find(alias)
    if alias_start == -1:
        return False

    alias_end = alias_start + len(alias)
    if match.end() <= alias_start:
        gap = context[match.end() : alias_start]
        return len(gap) <= max_gap and gap.strip() == ""
    if match.start() >= alias_end:
        return match.start() - alias_end <= max_gap
    return True


def _extract_dimensions(context: str, alias: str) -> list[str]:
    dimensions: list[str] = []

    for pattern in DIMENSION_PATTERNS:
        for match in pattern.finditer(context):
            if not _is_near_alias(context, alias, match):
                continue

            width = match.group(1)
            width_unit = match.group(2)
            depth = match.group(3)
            depth_unit = match.group(4)
            dimensions.append(f"{width}{width_unit}×{depth}{depth_unit}")

    for match in HEIGHT_RE.finditer(context):
        if not _is_near_alias(context, alias, match):
            continue
        dimensions.append(f"高度{match.group(1)}{match.group(2)}")

    return list(dict.fromkeys(dimensions))


def _looks_attached_to_alias(context: str, alias: str, match: re.Match[str]) -> bool:
    alias_start = context.find(alias)
    if alias_start == -1:
        return False

    alias_end = alias_start + len(alias)
    max_gap = 4

    if match.end() <= alias_start:
        gap = context[match.end() : alias_start]
        return len(gap) <= max_gap and bool(ATTACHED_GAP_RE.match(gap))

    if match.start() >= alias_end:
        gap = context[alias_end : match.start()]
        return len(gap) <= max_gap and bool(ATTACHED_GAP_RE.match(gap))

    return False


def _extract_quantity(context: str, alias: str) -> dict[str, Any] | None:
    candidates: list[dict[str, Any]] = []

    for match in NUMBER_UNIT_RE.finditer(context):
        if not _looks_attached_to_alias(context, alias, match):
            continue

        candidates.append(
            {
                "quantity": _format_number(match.group(1)),
                "unit": match.group(2),
                "text": match.group(0),
                "distance": abs(match.start() - context.find(alias)),
            }
        )

    if not candidates:
        return None

    return min(candidates, key=lambda item: item["distance"])


def _append_unique(items: list[str], value: str) -> None:
    if value and value not in items:
        items.append(value)


def _make_row(standard_item: str, rule: dict[str, Any], source_status: str) -> dict[str, Any]:
    return {
        "是否保留": True,
        "项目分类": rule.get("category", ""),
        "标准项目": standard_item,
        "原始命中词": [],
        "内容/尺寸/工艺": rule.get("default_desc", ""),
        "数量": 1,
        "单位": rule.get("default_unit", "项"),
        "单价": 0,
        "合计": 0,
        "报价类型": rule.get("quote_type", "固定单价"),
        "来源状态": source_status,
        "备注": "",
        "evidence_type": "",
        "evidence_text": "",
        "trigger_module": "",
        "_quantity_candidates": [],
        "_dimensions": [],
        "_module_hits": [],
        "_module_names": [],
        "_match_positions": [],
        "_notes": [],
        "_has_direct_hit": False,
        "_has_explicit_quantity": False,
    }


def _get_or_create_row(
    grouped: OrderedDict[str, dict[str, Any]],
    rules: dict[str, dict[str, Any]],
    standard_item: str,
    source_status: str,
    key: str | None = None,
) -> dict[str, Any] | None:
    rule = rules.get(standard_item)
    if not rule:
        return None

    row_key = key or standard_item
    if row_key not in grouped:
        grouped[row_key] = _make_row(standard_item, rule, source_status)
    return grouped[row_key]


def _find_module_hits(text: str) -> list[tuple[str, list[str]]]:
    hits: list[tuple[str, list[str]]] = []
    for trigger, standard_items in STRICT_MODULE_COMPLETION_ITEMS.items():
        if trigger in text:
            hits.append((trigger, standard_items))
    return hits


def _is_shadowed_by_strict_module(text: str, trigger: str, standard_items: list[str]) -> bool:
    strict_items = STRICT_MODULE_COMPLETION_ITEMS.get(trigger)
    if strict_items is not None:
        return list(standard_items) != strict_items

    for module_name, module_items in STRICT_MODULE_COMPLETION_ITEMS.items():
        if trigger not in module_name:
            continue
        start = text.find(module_name)
        if start == -1:
            continue
        trigger_start = text.find(trigger, start, start + len(module_name))
        if trigger_start != -1 and list(standard_items) != module_items:
            return True
    return False


def _is_allowed_dining_hit(matched_text: str, context: str) -> bool:
    if matched_text in DINING_DIRECT_TERMS:
        return True

    has_execution_context = _has_any_keyword(context, DINING_EXECUTION_TERMS)
    has_product_context = _has_any_keyword(context, DINING_PRODUCT_CONTEXT_TERMS)
    if has_product_context and not _has_any_keyword(context, DINING_STRONG_EXECUTION_TERMS):
        return False

    return has_execution_context


def _is_allowed_tea_break_hit(matched_text: str, context: str) -> bool:
    if matched_text == "茶歇" or matched_text == "茶歇服务":
        return True
    if _has_any_keyword(context, TEA_BREAK_PRODUCT_CONTEXT_TERMS):
        return False
    return _has_any_keyword(context, TEA_BREAK_EXECUTION_TERMS)


def _is_allowed_explicit_hit(standard_item: str, matched_text: str, context: str, full_text: str) -> bool:
    if any(
        term in context
        for term in [
            "活动调性",
            "活动目标",
            "项目目标",
            "项目背景",
            "整体思路",
            "核心创意",
            "活动亮点",
            "场景美学化",
            "体验游戏化",
            "内容复合化",
            "空间设计",
            "执行设计",
            "重点测试",
            "测试意图",
            "方案写法",
            "等形式",
        ]
    ):
        return False

    if standard_item == "阅读区布置":
        if matched_text not in READING_LAYOUT_ALLOWED_TRIGGERS:
            return False
        return not any(term in context or term in full_text for term in READING_LAYOUT_BLOCK_TERMS)

    if standard_item == "图书配置":
        if matched_text not in BOOK_CONFIG_ALLOWED_TRIGGERS:
            return False
        if matched_text == "图书" and "图书馆" in context:
            return False
        return not any(term in context or term in full_text for term in BOOK_CONFIG_BLOCK_TERMS)

    if standard_item == "音频二维码点位":
        return matched_text in AUDIO_QR_ALLOWED_TRIGGERS

    if standard_item == "餐饮体验":
        return _is_allowed_dining_hit(matched_text, context)

    if standard_item == "茶歇服务":
        return _is_allowed_tea_break_hit(matched_text, context)

    if standard_item == "帐篷摊位" and matched_text in {"市集", "市集活动"}:
        if any(term in context for term in ["生活消费", "消费场景", "传播力"]):
            return False
        if not any(term in context for term in ["设置", "配置", "摊位", "帐篷", "市集区", "招募", "商户", "摊主", "布置"]):
            return False

    if standard_item == "互动规则牌" and matched_text == "互动规则":
        if any(term in context for term in ["统一互动规则", "传播话题", "完整体验动线"]):
            return False

    if standard_item == "趣味互动游戏" and matched_text == "手作体验":
        return False

    return True


def _detect_activity_days(text: str) -> int | None:
    days: list[int] = []
    for match in DATE_RANGE_RE.finditer(text):
        start_month = int(match.group(1))
        start_day = int(match.group(2))
        end_month = int(match.group(3) or start_month)
        end_day = int(match.group(4))

        try:
            start = date(2000, start_month, start_day)
            end = date(2000, end_month, end_day)
        except ValueError:
            continue

        if end >= start:
            days.append((end - start).days + 1)

    for match in DOT_DATE_RANGE_RE.finditer(text):
        start_month = int(match.group(1))
        start_day = int(match.group(2))
        end_month = int(match.group(3) or start_month)
        end_day = int(match.group(4))

        try:
            start = date(2000, start_month, start_day)
            end = date(2000, end_month, end_day)
        except ValueError:
            continue

        if end >= start:
            days.append((end - start).days + 1)

    return max(days) if days else None


def _near_any_keyword(text: str, start: int, end: int, keywords: set[str], window: int = 30) -> bool:
    context = text[max(0, start - window) : min(len(text), end + window)]
    return any(keyword in context for keyword in keywords)


def _looks_like_activity_scale(text: str, start: int, end: int, window: int = 30) -> bool:
    context = text[max(0, start - window) : min(len(text), end + window)]
    return any(term in context for term in ["人次", "客流", "人流", "活动规模", "参与人数", "日均", "累计到访", "预计到访", "游客"])


def _has_unbound_quantity_signal(text: str) -> bool:
    return bool(NUMBER_UNIT_RE.search(text) and any(term in text for term in ["人次", "客流", "人流", "活动规模", "参与人数", "日均", "累计到访", "预计到访", "游客"]))


def _has_any_keyword(text: str, keywords: set[str]) -> bool:
    return any(keyword in text for keyword in keywords)


def _apply_pattern_quantity(
    text: str,
    row: dict[str, Any],
    pattern: re.Pattern[str],
    unit: str,
    notes: list[str],
    note: str,
    context_keywords: set[str] | None = None,
) -> bool:
    for match in pattern.finditer(text):
        if _looks_like_activity_scale(text, match.start(), match.end()):
            continue
        if context_keywords and not _near_any_keyword(text, match.start(), match.end(), context_keywords):
            continue

        value = next((group for group in match.groups() if group), None)
        if not value:
            continue

        row["数量"] = _format_number(value)
        row["单位"] = unit
        row["_has_explicit_quantity"] = True
        _append_unique(notes, note.format(value=value, text=match.group(0)))
        return True

    return False


def _apply_special_quantity_rules(text: str, row: dict[str, Any], notes: list[str]) -> None:
    standard_item = row["标准项目"]

    if standard_item == "帐篷摊位":
        _apply_pattern_quantity(
            text,
            row,
            MARKET_STALL_RE,
            "个",
            notes,
            "根据“{text}”识别摊位数量",
        )

    if standard_item in {"长桌宴", "餐饮体验"}:
        _apply_pattern_quantity(
            text,
            row,
            MEAL_QUANTITY_RE,
            "份" if standard_item == "长桌宴" else row.get("单位", "项"),
            notes,
            "根据“{text}”识别餐饮/体验数量",
            {"长桌宴", "花宴", "餐券", "餐饮", "美食", "围炉烤蚝", "茶席", "茶寮"},
        )

    if standard_item == "绣面体验":
        _apply_pattern_quantity(
            text,
            row,
            MEAL_QUANTITY_RE,
            "人",
            notes,
            "根据“{text}”识别绣面体验名额",
            {"绣面", "体验券"},
        )

    if standard_item in {"赛事活动", "趣味互动游戏"}:
        has_session = _apply_pattern_quantity(
            text,
            row,
            GAME_SESSION_RE,
            "场",
            notes,
            "根据“{text}”识别比赛/互动场次",
            {"比赛", "竞赛", "挑战赛", "争霸赛", "趣味互动", "游戏"},
        )
        player_match = GAME_PLAYER_RE.search(text)
        if player_match:
            _append_unique(notes, f"每场人数线索：{player_match.group(0)}")
            if not has_session:
                row["_has_explicit_quantity"] = True

    if standard_item == "倒计时海报":
        if "倒计时海报3天" in text or "前三天上线系列倒计时海报" in text or "活动前三天上线系列倒计时海报" in text:
            row["数量"] = 3
            row["单位"] = "张"
            row["_has_explicit_quantity"] = True
            _append_unique(notes, "根据倒计时海报3天/2天/1天推算3张")

    if standard_item in {"预热视频", "视频快剪"}:
        duration_match = VIDEO_DURATION_RE.search(text)
        if duration_match and _near_any_keyword(text, duration_match.start(), duration_match.end(), {"视频", "宣传片", "快剪", "短片"}):
            row["数量"] = 1
            row["单位"] = "条"
            row["_has_explicit_quantity"] = True
            _append_unique(notes, f"视频时长线索：{duration_match.group(0)}")

    for personnel_item, patterns in PERSONNEL_COUNT_PATTERNS.items():
        if standard_item != personnel_item:
            continue
        for pattern in patterns:
            match = pattern.search(text)
            if match:
                row["数量"] = _format_number(match.group(1))
                row["单位"] = "人"
                row["_has_explicit_quantity"] = True
                _append_unique(notes, f"人员数量线索：{match.group(0)}")
                break

    if standard_item == "达人KOL推广":
        _apply_pattern_quantity(
            text,
            row,
            KOL_COUNT_RE,
            "人",
            notes,
            "根据“{text}”识别达人/KOL数量",
        )

    if standard_item == "合作品牌联合宣传":
        _apply_pattern_quantity(
            text,
            row,
            BRAND_COUNT_RE,
            "家",
            notes,
            "根据“{text}”识别品牌联动数量",
        )

    if standard_item == "图书配置":
        match = BOOK_COUNT_RE.search(text)
        if match:
            row["数量"] = _format_number(match.group(1))
            row["单位"] = "册"
            row["_has_explicit_quantity"] = True
            _append_unique(notes, f"方案提及约{match.group(1)}册图书")

    if standard_item == "互动规则牌":
        match = INTERACTIVE_PROJECT_RE.search(text)
        if match:
            row["数量"] = _format_number(match.group(1))
            row["单位"] = "个"
            row["_has_explicit_quantity"] = True
            _append_unique(notes, f"根据{match.group(1)}个互动项目推算")

    if standard_item == "印章":
        match = STAMP_COUNT_RE.search(text)
        if match:
            row["数量"] = _format_number(match.group(1))
            row["单位"] = "个"
            row["_has_explicit_quantity"] = True
            _append_unique(notes, f"根据集章互动{match.group(1)}枚印章推算")


def _is_complex_rule(rule: dict[str, Any]) -> bool:
    return rule.get("risk_level") == "high" or rule.get("quote_type") == "模糊报价"


def _line_looks_like_title(line: str) -> bool:
    stripped = line.strip()
    if not (2 <= len(stripped) <= 36):
        return False
    if any(keyword in stripped for keyword in CANDIDATE_KEYWORDS):
        return True
    return bool(re.match(r"^[（(]?[一二三四五六七八九十]+[）)、.．]|^\d+[.、]", stripped))


def _candidate_action(candidate: str) -> str:
    if any(keyword in candidate for keyword in ["开幕", "市集", "夜游", "保障", "传播", "展", "音乐会", "宴"]):
        return "加入模块补全规则"
    if any(keyword in candidate for keyword in ["装置", "物料", "道具", "礼品", "海报", "视频", "推文", "直播"]):
        return "加入关键词库"
    if any(keyword in candidate for keyword in ["人员", "保障", "安全"]):
        return "人工确认"
    return "人工确认"


def _candidate_context(text: str, term: str) -> str:
    index = text.find(term)
    if index == -1:
        return ""
    context = text[max(0, index - 40) : min(len(text), index + len(term) + 60)]
    return re.sub(r"\s+", " ", context).strip()


def extract_unrecognized_candidates(
    text: str,
    recognized_rows: list[dict[str, Any]],
    rules: dict[str, dict[str, Any]],
    ignored_terms: list[str] | None = None,
    limit: int = 80,
) -> list[dict[str, Any]]:
    """Return suspicious activity/quote phrases that are not covered by current rules."""
    ignored_terms = ignored_terms or []
    known_terms: set[str] = set()
    for row in recognized_rows:
        known_terms.add(str(row.get("标准项目", "")))
        for term in str(row.get("原始命中词", "")).split("、"):
            known_terms.add(term)

    for standard_item, rule in rules.items():
        known_terms.add(standard_item)
        known_terms.update(rule.get("aliases", []))

    candidates: OrderedDict[str, dict[str, Any]] = OrderedDict()

    for raw_line in text.splitlines():
        line = raw_line.strip()
        if not line:
            continue
        if _line_looks_like_title(line):
            cleaned = re.sub(r"^[（(]?[一二三四五六七八九十0-9]+[）)、.．\s]*", "", line).strip()
            if 2 <= len(cleaned) <= 36:
                candidates[cleaned] = {
                    "候选词": cleaned,
                    "所在上下文": _candidate_context(text, cleaned),
                    "建议动作": _candidate_action(cleaned),
                    "是否加入规则库": False,
                }

    phrase_pattern = re.compile(r"[\u4e00-\u9fa5A-Za-z0-9·《》“”()（）]{2,36}")
    for match in phrase_pattern.finditer(text):
        phrase = match.group(0).strip("，。、；：:（）()“”")
        if not (2 <= len(phrase) <= 36):
            continue
        if not any(keyword in phrase for keyword in CANDIDATE_KEYWORDS):
            continue
        candidates.setdefault(
            phrase,
            {
                "候选词": phrase,
                "所在上下文": _candidate_context(text, phrase),
                "建议动作": _candidate_action(phrase),
                "是否加入规则库": False,
            },
        )

    filtered: list[dict[str, Any]] = []
    for candidate, row in candidates.items():
        if any(term and (candidate == term or candidate in term or term in candidate) for term in ignored_terms):
            continue
        if any(term and (candidate == term or candidate in term or term in candidate) for term in known_terms):
            continue
        filtered.append(row)
        if len(filtered) >= limit:
            break

    return filtered


def calculate_coverage(recognized_count: int, candidate_count: int) -> float:
    total = recognized_count + candidate_count
    if total <= 0:
        return 1.0
    return min(1.0, recognized_count / total)


def extract_quote_items(text: str, rules: dict[str, dict[str, Any]]) -> list[dict[str, Any]]:
    """Extract quote item hits without collapsing same-name items across sections."""
    matches = find_alias_matches(text, rules)
    grouped: OrderedDict[str, dict[str, Any]] = OrderedDict()
    activity_days = _detect_activity_days(text)

    for match in matches:
        standard_item = match["standard_item"]
        if standard_item == "长桌宴" and not _has_any_keyword(text, LONG_TABLE_ALLOWED_KEYWORDS):
            continue
        if standard_item == "茶席茶寮" and not _has_any_keyword(text, TEA_SEAT_ALLOWED_KEYWORDS):
            continue
        if standard_item == "餐饮体验" and match["matched_text"] in TEA_ONLY_FOOD_ALIASES:
            continue

        context = _context(text, match["start"], match["end"])
        if not _is_allowed_explicit_hit(standard_item, match["matched_text"], context, text):
            continue
        quantity = _extract_quantity(context, match["matched_text"])
        dimensions = _extract_dimensions(context, match["matched_text"])

        rule = rules.get(standard_item, {})
        source_status = "需确认" if _is_complex_rule(rule) else match["default_source_status"]
        row_key = f"explicit::{standard_item}::{match['start']}::{match['matched_text']}"
        row = _get_or_create_row(grouped, rules, standard_item, source_status, row_key)
        if row is None:
            continue

        row["_has_direct_hit"] = True
        row["evidence_type"] = "explicit_text"
        if match["matched_text"] not in str(row.get("evidence_text", "")).split("、"):
            row["evidence_text"] = "、".join(filter(None, [str(row.get("evidence_text", "")).strip("、"), match["matched_text"]])).strip("、")
        if row["来源状态"] == "系统推算":
            row["来源状态"] = source_status
        _append_unique(row["原始命中词"], match["matched_text"])
        row["_match_positions"].append(match["start"])

        if quantity:
            row["_quantity_candidates"].append(
                {
                    "matched_text": match["matched_text"],
                    "quantity": quantity["quantity"],
                    "unit": quantity["unit"],
                    "text": quantity["text"],
                }
            )

        if dimensions:
            row["_dimensions"].extend(dimensions)

    for module, standard_items in _find_module_hits(text):
        for standard_item in standard_items:
            if standard_item == "长桌宴" and not _has_any_keyword(text, LONG_TABLE_ALLOWED_KEYWORDS):
                continue
            if standard_item == "茶席茶寮" and not _has_any_keyword(text, TEA_SEAT_ALLOWED_KEYWORDS):
                continue

            module_start = text.find(module.split("、")[0])
            row_key = f"module::{standard_item}::{module_start}::{module}"
            row = _get_or_create_row(grouped, rules, standard_item, "系统推算", row_key)
            if row is None:
                continue
            _append_unique(row["_module_hits"], module)
            _append_unique(row["_module_names"], module)
            row["_match_positions"].append(module_start)
            if not row.get("evidence_type"):
                row["evidence_type"] = "module_completion"
                row["evidence_text"] = module
                row["trigger_module"] = module

    rows: list[dict[str, Any]] = []

    for row in grouped.values():
        hit_terms = list(dict.fromkeys(row["原始命中词"]))
        quantity_candidates = row.pop("_quantity_candidates")
        dimensions = list(dict.fromkeys(row.pop("_dimensions")))
        module_hits = list(dict.fromkeys(row.pop("_module_hits")))
        module_names = list(dict.fromkeys(row.pop("_module_names")))
        match_positions = [position for position in row.pop("_match_positions") if isinstance(position, int) and position >= 0]
        notes = row.pop("_notes")
        has_direct_hit = row.pop("_has_direct_hit")
        has_explicit_quantity = row.pop("_has_explicit_quantity")

        row["原始命中词"] = "、".join(hit_terms) if hit_terms else "模块补全"

        if len(quantity_candidates) == 1:
            row["数量"] = quantity_candidates[0]["quantity"]
            row["单位"] = quantity_candidates[0]["unit"]
            has_explicit_quantity = True
        elif len(quantity_candidates) > 1:
            detail = "、".join(
                f"{item['matched_text']}({item['text']})" for item in quantity_candidates
            )
            _append_unique(notes, f"数量线索：{detail}")

        _apply_special_quantity_rules(text, row, notes)
        has_explicit_quantity = has_explicit_quantity or row.pop("_has_explicit_quantity", False)

        if dimensions:
            dimension_note = "；".join(f"尺寸线索：{dimension}" for dimension in dimensions)
            _append_unique(notes, dimension_note)

        if len(module_hits) > 1:
            _append_unique(notes, f"由多个活动模块共同命中：{'、'.join(module_hits)}")
            if row.get("evidence_type") == "module_completion":
                row["evidence_text"] = "、".join(module_hits)
                row["trigger_module"] = "、".join(module_hits)
        elif module_hits and not has_direct_hit:
            _append_unique(notes, f"由活动模块补全：{module_hits[0]}")
            if row.get("evidence_type") == "module_completion":
                row["evidence_text"] = module_hits[0]
                row["trigger_module"] = module_hits[0]

        if row["标准项目"] == "节目演出" and hit_terms:
            _append_unique(notes, f"演出形式线索：{'、'.join(hit_terms)}")

        if activity_days and (
            row.get("项目分类") in {"人员执行类", "灯光音响类", "舞台搭建类"}
            or row["标准项目"] in {"对讲机", "铁马围挡", "电力保障", "桌椅布置"}
        ):
            _append_unique(notes, f"活动周期涉及多日（约{activity_days}天），人员/设备租赁天数需人工确认")

        if not has_explicit_quantity:
            row["来源状态"] = "需确认"
            if _has_unbound_quantity_signal(text):
                _append_unique(notes, "方案提及数量线索，但未能确认是否对应本项目")
            _append_unique(notes, "方案提及该项目，但未明确数量，需人工确认数量")

        if row["来源状态"] == "需确认":
            _append_unique(notes, "信息不足，需人工确认报价口径")

        row["备注"] = "；".join(notes)
        row["匹配位置"] = match_positions
        row["命中模块"] = module_names

        if row.get("evidence_type") not in EXPLICIT_EVIDENCE_TYPES:
            continue

        rows.append(row)

    return rows
