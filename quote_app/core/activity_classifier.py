"""Classify activity types and suggest missing template quote items."""

from __future__ import annotations

from typing import Any


LONG_TABLE_ALLOWED_KEYWORDS = [
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
]

SUGGESTION_TRIGGER_KEYWORDS: dict[str, list[str]] = {
    "长桌宴": LONG_TABLE_ALLOWED_KEYWORDS,
    "餐饮体验": ["美食", "餐饮", "市集", "食集", "小吃", "饮品", "轻食", "茶饮", "餐饮体验", "烤蚝"],
    "茶席茶寮": ["茶席", "茶寮", "围炉煮茶", "茶饮空间"],
    "帐篷摊位": ["市集", "食集", "美食摊位", "餐饮品牌", "摊位", "帐篷"],
    "市集招募服务": ["市集", "食集", "摊位", "商家招募", "品牌入驻", "餐饮品牌"],
    "NPC互动服务": ["NPC", "巡游", "角色", "沉浸式互动", "沉浸式", "游走式表演"],
    "电力保障": ["电力", "供电", "灯光", "夜游", "发电", "配电"],
    "演艺服化道": ["演出", "巡游", "走秀", "音乐会", "NPC", "服饰", "妆造"],
    "互动道具": ["互动", "游戏", "道具", "转盘", "投票", "挑战", "比赛"],
    "培训讲师": ["培训", "内训", "授课", "讲师", "导师", "工作坊", "课程"],
    "课程物料": ["课程物料", "培训物料", "学员手册", "培训手册", "讲义", "教材", "学习资料"],
    "培训设备": ["培训设备", "投影", "投屏", "白板", "翻页笔", "麦克风", "音响"],
    "会议场地": ["会议室", "培训室", "会场", "会议场地", "培训场地", "会议厅"],
    "茶歇服务": ["茶歇", "茶点", "咖啡", "点心", "水果"],
    "拓展教练": ["团建", "拓展", "教练", "带队教练", "拓展训练"],
    "拓展器材": ["拓展器材", "拓展道具", "绳索", "鼓球", "团队挑战"],
    "团建道具": ["团建道具", "破冰道具", "游戏道具", "分组道具", "任务卡"],
    "证书奖杯": ["证书", "奖杯", "奖牌", "结业证书", "荣誉证书", "颁奖"],
    "美陈装置": ["美陈", "装置", "花灯", "花境", "打卡", "互动墙"],
    "安保人员": ["安保", "安检", "限流", "秩序", "安全保障"],
    "医疗保障": ["医疗", "急救", "AED", "120"],
    "交通停车引导": ["交通", "停车", "车辆引导", "人车分流"],
}

ACTIVITY_TYPE_TEMPLATES: dict[str, dict[str, list[str]]] = {
    "美食节类": {
        "keywords": ["美食节", "生蚝荟", "长桌宴", "食集", "美食摊位", "餐饮品牌", "食速挑战", "餐券", "茶席", "围炉", "烤蚝", "饮品", "农产品", "海鲜市集"],
        "focus_items": ["帐篷摊位", "餐饮体验", "长桌宴", "茶席茶寮", "市集招募服务", "赛事活动", "文创礼品", "导视指引", "执行人员", "安保人员", "饮用水", "餐费"],
    },
    "夜游文旅类": {
        "keywords": ["夜游", "夜间", "花灯", "灯展", "灯光艺术", "河灯", "赏花", "花境", "古村", "夜绽", "夜色", "流光", "巡游", "NPC", "沉浸式"],
        "focus_items": ["灯光艺术装置", "花艺花境", "氛围布置", "节目演出", "NPC互动服务", "导视指引", "打卡装置", "电力保障", "安保人员"],
    },
    "阅读文化类": {
        "keywords": ["阅读", "共读", "书签", "图书", "二维码", "有声阅读", "阅读角", "心得墙", "书海寻宝", "集章卡", "非遗拓印", "故事墙"],
        "focus_items": ["阅读区布置", "图书配置", "活动书签", "音频二维码点位", "故事画面墙", "互动规则牌", "手作体验材料", "主持人", "执行人员"],
    },
    "民俗节庆类": {
        "keywords": ["三月三", "民族", "民俗", "非遗", "调声", "服饰走秀", "民族团结", "文化符号", "传统节日", "长桌宴", "民族运动会"],
        "focus_items": ["节目演出", "演艺服化道", "展板", "美陈装置", "趣味互动游戏", "长桌宴", "帐篷摊位", "文创礼品", "导视指引", "安保人员"],
    },
    "市集活动类": {
        "keywords": ["市集", "摊位", "帐篷", "展销", "手作", "文创", "农产品", "品牌入驻", "商家招募", "地摊市集"],
        "focus_items": ["帐篷摊位", "地摊市集布置", "市集招募服务", "摊位视觉物料", "导视指引", "执行人员", "运输费"],
    },
    "演艺活动类": {
        "keywords": ["音乐会", "演唱", "乐队", "舞蹈", "开场表演", "节目演出", "巡游", "走秀", "主持人", "舞台", "灯光音响"],
        "focus_items": ["节目演出", "舞台搭建", "舞台视觉包装", "灯光音响套装", "主持人", "技术人员", "摄影摄像服务"],
    },
    "展览展陈类": {
        "keywords": ["展览", "艺术展", "展板", "科普展示", "图文展板", "装置", "打卡点", "互动墙", "展示墙"],
        "focus_items": ["艺术展陈", "展板", "美陈装置", "打卡装置", "灯光艺术装置", "导视指引"],
    },
    "互动体验类": {
        "keywords": ["互动", "游园", "游园会", "集章", "通行证", "打卡", "挑战", "比赛", "运动会", "手作", "拓印", "体验"],
        "focus_items": ["趣味互动游戏", "互动道具", "互动规则牌", "通行证", "印章", "文创礼品", "执行人员"],
    },
    "宣传推广类": {
        "keywords": ["传播", "预热", "推文", "海报", "话题", "短视频", "直播", "媒体", "KOL", "达人", "宣传"],
        "focus_items": ["预热视频", "公众号软文", "倒计时海报", "话题传播", "图片直播", "视频快剪", "主流媒体宣传", "达人KOL推广"],
    },
    "企业团建类": {
        "keywords": ["团建", "团队建设", "团队拓展", "拓展训练", "破冰", "协作挑战", "凝聚力", "户外拓展", "结营", "复盘共创"],
        "focus_items": ["拓展教练", "拓展器材", "团建道具", "互动规则牌", "文创礼品", "摄影摄像服务", "餐饮体验", "茶歇服务"],
    },
    "培训内训类": {
        "keywords": ["培训", "内训", "专题培训", "课程", "授课", "讲师", "导师", "分组研讨", "工作坊", "情景演练", "结业测评"],
        "focus_items": ["培训讲师", "课程物料", "培训设备", "会议场地", "茶歇服务", "证书奖杯", "摄影摄像服务"],
    },
    "企业会务类": {
        "keywords": ["公司会议", "企业会议", "年度会议", "战略会", "复盘会", "动员会", "表彰会", "研讨会", "会务", "会议议程"],
        "focus_items": ["会议场地", "签到物料", "主持人", "灯光音响套装", "培训设备", "茶歇服务", "摄影摄像服务"],
    },
}


READING_TYPE_STRONG_KEYWORDS = {
    "阅读",
    "共读",
    "书签",
    "有声阅读",
    "阅读角",
    "心得墙",
    "书海寻宝",
    "故事墙",
    "双语图书",
    "图书阅读区",
}

TRAINING_TYPE_STRONG_KEYWORDS = {
    "专题培训",
    "培训课程",
    "培训内容",
    "培训安排",
    "培训计划",
    "培训对象",
    "培训目标",
    "参训",
    "课程",
    "授课",
    "讲师",
    "导师",
    "分组研讨",
    "工作坊",
    "情景演练",
    "结业测评",
    "开班",
}


def _should_keep_activity_type(activity_type: str, text: str, matched_keywords: list[str]) -> bool:
    if activity_type == "阅读文化类":
        return any(keyword in text for keyword in READING_TYPE_STRONG_KEYWORDS)

    if activity_type == "培训内训类":
        if any(keyword in text for keyword in TRAINING_TYPE_STRONG_KEYWORDS):
            return True
        weak_training_mentions = ["工作人员培训", "人员培训", "员工培训", "志愿者培训"]
        if matched_keywords == ["培训"] and any(term in text for term in weak_training_mentions):
            return False
        return len(matched_keywords) >= 2

    return True


def _has_text_signal(text: str, standard_item: str, rule: dict[str, Any]) -> bool:
    trigger_keywords = SUGGESTION_TRIGGER_KEYWORDS.get(standard_item)
    if trigger_keywords is not None:
        return any(keyword in text for keyword in trigger_keywords)

    aliases = [str(alias) for alias in rule.get("aliases", [])]
    return standard_item in text or any(alias and alias in text for alias in aliases)


def classify_activity_types(text: str) -> list[dict[str, Any]]:
    """Return all matched activity type templates for the plan text."""
    results: list[dict[str, Any]] = []

    for activity_type, config in ACTIVITY_TYPE_TEMPLATES.items():
        matched_keywords = [keyword for keyword in config["keywords"] if keyword in text]
        if not matched_keywords:
            continue
        if not _should_keep_activity_type(activity_type, text, matched_keywords):
            continue
        results.append(
            {
                "活动类型": activity_type,
                "命中关键词": "、".join(matched_keywords),
                "重点规则": "、".join(config["focus_items"]),
                "focus_items": config["focus_items"],
            }
        )

    return results


def build_suggested_items(
    activity_types: list[dict[str, Any]],
    recognized_rows: list[dict[str, Any]],
    rules: dict[str, dict[str, Any]],
    text: str = "",
) -> list[dict[str, Any]]:
    """Suggest missing focus items from matched activity templates."""
    recognized_items = {str(row.get("标准项目", "")) for row in recognized_rows}
    suggestions: dict[str, dict[str, Any]] = {}

    for activity_type in activity_types:
        for standard_item in activity_type.get("focus_items", []):
            if standard_item in recognized_items or standard_item not in rules:
                continue

            rule = rules[standard_item]
            if not _has_text_signal(text, standard_item, rule):
                continue

            source_status = "需确认" if rule.get("risk_level") == "high" else "系统推算"
            if standard_item not in suggestions:
                suggestions[standard_item] = {
                    "是否加入报价单": False,
                    "建议补充项": standard_item,
                    "项目分类": rule.get("category", ""),
                    "报价类型": rule.get("quote_type", ""),
                    "来源状态": source_status,
                    "建议原因": [],
                    "备注": rule.get("default_desc", ""),
                }
            suggestions[standard_item]["建议原因"].append(activity_type["活动类型"])

    rows: list[dict[str, Any]] = []
    for row in suggestions.values():
        row["建议原因"] = f"活动类型模板建议：{'、'.join(row['建议原因'])}"
        rows.append(row)

    return rows
