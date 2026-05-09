"""Load rules and normalize raw hit words into standard quote items."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from .rule_presets import PRIORITY_STANDARD_ITEMS, STANDARD_RULE_UPDATES


QUOTE_TYPES = ("固定单价", "档位报价", "模糊报价")
SOURCE_STATUSES = ("方案明确", "系统推算", "需确认")


DEFAULT_RULES: dict[str, dict[str, Any]] = {
    "帐篷摊位": {
        "category": "市集摊位类",
        "aliases": ["市集摊位", "帐篷摊位", "美食摊位", "文创摊位", "手工摊位", "农产品摊位", "茶肆帐篷", "商家市集", "特色摊位"],
        "quote_type": "档位报价",
        "default_unit": "个",
        "default_desc": "帐篷摊位，具体配置按方案确认",
    },
    "节目演出": {
        "category": "演艺节目类",
        "aliases": ["开场表演", "节目表演", "节目演出", "演艺演出", "乐队演绎", "舞蹈演绎", "脱口秀演绎", "非遗演绎", "琼剧表演", "巡游表演", "鼓舞表演", "歌舞表演"],
        "quote_type": "模糊报价",
        "default_unit": "场",
        "default_desc": "演出形式按方案内容确认，价格按人数、场次、时长、演员级别核定。",
        "default_source_status": "需确认",
    },
    "现场工作证": {
        "category": "后勤物料类",
        "aliases": ["工作证", "工作牌", "工作卡", "人员证件"],
        "quote_type": "固定单价",
        "default_unit": "个",
        "default_desc": "现场工作人员证件",
    },
    "指引牌": {
        "category": "视觉物料类",
        "aliases": ["指引牌", "导视牌", "区域指引牌", "活动现场牌子", "入口指示牌", "T型指引牌", "木质指引牌"],
        "quote_type": "固定单价",
        "default_unit": "个",
        "default_desc": "现场导视及动线指引",
    },
    "执行人员": {
        "category": "人员执行类",
        "aliases": ["工作人员", "执行人员", "现场执行人员", "现场控场人员", "点位执行", "各点位管理人员"],
        "quote_type": "固定单价",
        "default_unit": "人",
        "default_desc": "现场执行及点位管理人员",
    },
    "安保人员": {
        "category": "人员执行类",
        "aliases": ["安保人员", "安保团队", "专业安保", "现场安保管理"],
        "quote_type": "固定单价",
        "default_unit": "人",
        "default_desc": "现场安保人员",
    },
    "搭建撤场工人": {
        "category": "人员执行类",
        "aliases": ["搭建工人", "撤场工人", "工人", "进场撤场工人"],
        "quote_type": "固定单价",
        "default_unit": "人",
        "default_desc": "进场搭建及撤场人工",
    },
    "餐费": {
        "category": "后勤物料类",
        "aliases": ["餐费", "工作餐", "人员餐费"],
        "quote_type": "固定单价",
        "default_unit": "份",
        "default_desc": "工作人员餐费",
    },
    "饮用水": {
        "category": "后勤物料类",
        "aliases": ["饮用水", "矿泉水", "瓶装水"],
        "quote_type": "固定单价",
        "default_unit": "箱",
        "default_desc": "现场饮用水",
    },
    "对讲机": {
        "category": "后勤物料类",
        "aliases": ["对讲机", "通讯设备"],
        "quote_type": "固定单价",
        "default_unit": "台",
        "default_desc": "现场通讯设备",
    },
    "运输费": {
        "category": "后勤物料类",
        "aliases": ["运输", "运输费", "物料运输", "物流运输"],
        "quote_type": "固定单价",
        "default_unit": "项",
        "default_desc": "物料运输及物流服务",
    },
    "氛围布置": {
        "category": "氛围装饰类",
        "aliases": ["氛围布置", "氛围营造", "现场氛围", "现场氛围装饰", "场景打造", "美陈布置", "美陈装置", "主题美陈"],
        "quote_type": "模糊报价",
        "default_unit": "项",
        "default_desc": "现场氛围装饰，需结合面积、材质、数量及工艺确认",
        "default_source_status": "需确认",
    },
    "舞台搭建": {
        "category": "舞台搭建类",
        "aliases": ["舞台", "主舞台", "舞台区", "钢架舞台", "T台"],
        "quote_type": "档位报价",
        "default_unit": "项",
        "default_desc": "舞台搭建，尺寸、结构及饰面按方案确认",
        "default_source_status": "需确认",
    },
    "舞台背景": {
        "category": "舞台搭建类",
        "aliases": ["舞台背景", "背景板", "主舞台背景", "舞台背景板", "异形背景"],
        "quote_type": "档位报价",
        "default_unit": "项",
        "default_desc": "舞台背景制作安装，材质、尺寸及造型按方案确认",
        "default_source_status": "需确认",
    },
    "音响设备": {
        "category": "灯光音响类",
        "aliases": ["音响", "音响设备", "线阵音响", "扩声音响"],
        "quote_type": "档位报价",
        "default_unit": "套",
        "default_desc": "音响设备，配置按场地规模及演出需求确认",
        "default_source_status": "需确认",
    },
    "灯光设备": {
        "category": "灯光音响类",
        "aliases": ["灯光", "舞台灯光", "灯光设备", "帕灯", "面光灯", "射灯"],
        "quote_type": "档位报价",
        "default_unit": "套",
        "default_desc": "灯光设备，灯具数量及控台配置按方案确认",
        "default_source_status": "需确认",
    },
    "灯光音响套装": {
        "category": "灯光音响类",
        "aliases": ["灯光音响", "灯光音响设备"],
        "quote_type": "档位报价",
        "default_unit": "套",
        "default_desc": "灯光音响套装，按活动规模、演出需求及场地条件确认",
        "default_source_status": "需确认",
    },
}

READING_ACTIVITY_RULES: dict[str, dict[str, Any]] = {
    "签到墙": {
        "category": "视觉物料类",
        "aliases": ["入梦签到墙", "签到墙", "签到背景", "签到背景墙", "签到装置"],
        "quote_type": "档位报价",
        "default_unit": "项",
        "default_desc": "签到墙设计制作及现场安装，具体尺寸和材质按方案确认",
    },
    "签到物料": {
        "category": "活动道具类",
        "aliases": ["签到贴纸", "姓名贴", "贴纸", "签到卡", "签到笔", "签到物料"],
        "quote_type": "固定单价",
        "default_unit": "批",
        "default_desc": "签到贴纸、签到笔、姓名贴等基础签到物料",
    },
    "活动书签": {
        "category": "视觉物料类",
        "aliases": ["活动书签", "书签", "定制书签"],
        "quote_type": "固定单价",
        "default_unit": "张",
        "default_desc": "活动定制书签设计制作",
    },
    "指引物料": {
        "category": "视觉物料类",
        "aliases": ["指引物料", "指引牌", "导视牌", "活动指引", "现场指引", "区域指引牌"],
        "quote_type": "固定单价",
        "default_unit": "个",
        "default_desc": "活动现场导视及指引物料",
    },
    "主持人": {
        "category": "人员执行类",
        "aliases": ["主持人", "主持", "主持引导", "由主持人引导"],
        "quote_type": "档位报价",
        "default_unit": "人",
        "default_desc": "活动主持人服务，按场次和时长核价",
    },
    "摄影摄像服务": {
        "category": "宣传拍摄类",
        "aliases": ["合影留念", "统一合影", "活动记录", "记录活动过程", "传播素材", "现场拍摄", "摄影", "摄像"],
        "quote_type": "档位报价",
        "default_unit": "项",
        "default_desc": "活动摄影摄像及现场素材记录服务",
    },
    "阅读点位": {
        "category": "活动布置类",
        "aliases": ["阅读点位", "隐藏阅读点位", "隐藏书籍", "书海寻宝", "阅读任务", "寻宝点位"],
        "quote_type": "模糊报价",
        "default_unit": "项",
        "default_desc": "阅读寻宝点位布置，需确认点位数量、物料和奖品配置",
        "default_source_status": "需确认",
    },
    "小礼品": {
        "category": "活动道具类",
        "aliases": ["小礼品", "礼品", "周边礼品", "新媒体周边礼品", "奖品", "兑换礼品"],
        "quote_type": "固定单价",
        "default_unit": "份",
        "default_desc": "活动参与礼品或互动奖品",
    },
    "集章卡": {
        "category": "视觉物料类",
        "aliases": ["集章卡", "打卡卡", "通关卡", "任务卡"],
        "quote_type": "固定单价",
        "default_unit": "张",
        "default_desc": "活动集章卡设计制作",
    },
    "印章": {
        "category": "活动道具类",
        "aliases": ["印章", "活动印章", "集章印章", "一枚印章"],
        "quote_type": "固定单价",
        "default_unit": "个",
        "default_desc": "活动互动印章",
    },
    "非遗拓印材料": {
        "category": "活动道具类",
        "aliases": ["非遗拓印", "拓印体验", "拓印材料", "黎族图腾拓印", "拓印作品"],
        "quote_type": "模糊报价",
        "default_unit": "批",
        "default_desc": "非遗拓印体验材料包，需确认参与人数和材料配置",
        "default_source_status": "需确认",
    },
    "非遗传承人": {
        "category": "人员执行类",
        "aliases": ["非遗传承人", "传承人指导", "在非遗传承人指导下"],
        "quote_type": "档位报价",
        "default_unit": "人",
        "default_desc": "非遗传承人或体验老师服务",
    },
    "故事画面墙": {
        "category": "视觉物料类",
        "aliases": ["故事画面墙", "画面墙", "作品展示墙", "展示墙", "阅读心得展示墙", "心得墙"],
        "quote_type": "档位报价",
        "default_unit": "项",
        "default_desc": "互动展示墙设计制作及安装",
    },
    "互动规则牌": {
        "category": "视觉物料类",
        "aliases": ["游戏规则", "互动规则", "规则牌", "活动规则"],
        "quote_type": "固定单价",
        "default_unit": "个",
        "default_desc": "互动项目规则说明牌",
    },
    "主题卡": {
        "category": "活动道具类",
        "aliases": ["主题卡", "抽取主题卡", "随机抽取一张主题卡"],
        "quote_type": "固定单价",
        "default_unit": "套",
        "default_desc": "阅读互动主题卡制作",
    },
    "诗句展示装置": {
        "category": "美陈装置类",
        "aliases": ["诗意阶梯", "诗句展示装置", "诗句展示", "空间装置"],
        "quote_type": "模糊报价",
        "default_unit": "项",
        "default_desc": "诗句展示装置及空间氛围布置，需确认尺寸、材质和安装方式",
        "default_source_status": "需确认",
    },
    "打卡点位": {
        "category": "美陈装置类",
        "aliases": ["打卡点位", "打卡装置", "拍照打卡", "引导公众参与拍照"],
        "quote_type": "模糊报价",
        "default_unit": "项",
        "default_desc": "活动打卡点位搭建，需确认点位数量和制作形式",
        "default_source_status": "需确认",
    },
    "阅读区布置": {
        "category": "活动布置类",
        "aliases": ["阅读区域", "阅读区", "候船微阅读角", "双语阅读角", "中英文阅读区域", "阅读空间"],
        "quote_type": "档位报价",
        "default_unit": "项",
        "default_desc": "阅读区域桌椅、书籍陈列及基础氛围布置",
    },
    "图书配置": {
        "category": "活动道具类",
        "aliases": ["图书", "双语图书", "200册图书", "提供约200册图书"],
        "quote_type": "模糊报价",
        "default_unit": "批",
        "default_desc": "活动图书配置，需确认图书来源、数量和是否采购",
        "default_source_status": "需确认",
    },
    "金句撕页": {
        "category": "活动道具类",
        "aliases": ["金句撕页", "撕页互动"],
        "quote_type": "固定单价",
        "default_unit": "批",
        "default_desc": "金句撕页互动材料制作",
    },
    "音频二维码点位": {
        "category": "视觉物料类",
        "aliases": ["有声阅读二维码", "音频二维码", "二维码点位", "扫码收听", "桌牌二维码"],
        "quote_type": "固定单价",
        "default_unit": "个",
        "default_desc": "音频二维码点位画面及桌牌制作",
    },
    "志愿者": {
        "category": "人员执行类",
        "aliases": ["志愿者", "基础引导服务", "引导服务"],
        "quote_type": "固定单价",
        "default_unit": "人",
        "default_desc": "志愿者或现场引导人员服务",
    },
    "传播图文": {
        "category": "宣传推广类",
        "aliases": ["图文内容", "现场图文", "活动总结内容", "宣传图文", "发布活动预告"],
        "quote_type": "档位报价",
        "default_unit": "项",
        "default_desc": "活动图文内容策划、编辑及发布",
    },
    "短视频内容": {
        "category": "宣传推广类",
        "aliases": ["短视频", "现场短视频", "活动短视频", "精彩瞬间"],
        "quote_type": "档位报价",
        "default_unit": "条",
        "default_desc": "活动短视频拍摄剪辑或传播内容制作",
    },
}


def _rule_risk_level(quote_type: str) -> str:
    return {"固定单价": "low", "档位报价": "medium", "模糊报价": "high"}.get(quote_type, "medium")


def _merge_rule(base: dict[str, Any], update: dict[str, Any]) -> dict[str, Any]:
    merged = {**base, **update}
    aliases = list(dict.fromkeys([*base.get("aliases", []), *update.get("aliases", [])]))
    merged["aliases"] = aliases
    return merged


def _merge_rule_sets(target: dict[str, dict[str, Any]], updates: dict[str, dict[str, Any]]) -> None:
    for standard_item, rule in updates.items():
        if standard_item in target:
            target[standard_item] = _merge_rule(target[standard_item], rule)
        else:
            target[standard_item] = rule


def enrich_rule_metadata(standard_item: str, rule: dict[str, Any]) -> dict[str, Any]:
    enriched = dict(rule)
    enriched.setdefault("standard_item", standard_item)
    enriched.setdefault("auto_complete_with", [])
    enriched.setdefault("risk_level", _rule_risk_level(enriched.get("quote_type", "档位报价")))
    if "need_confirm_fields" not in enriched:
        enriched["need_confirm_fields"] = (
            ["数量"] if enriched["risk_level"] == "low" else ["数量", "规格", "配置", "天数"]
        )
    return enriched


DEFAULT_RULES.update(READING_ACTIVITY_RULES)
_merge_rule_sets(DEFAULT_RULES, STANDARD_RULE_UPDATES)


def ensure_rules_config(path: str | Path) -> Path:
    """Create a default rules config when the file is missing."""
    path = Path(path)
    if not path.exists():
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps(DEFAULT_RULES, ensure_ascii=False, indent=2), encoding="utf-8")
    return path


def load_rules(path: str | Path) -> dict[str, dict[str, Any]]:
    """Load and lightly validate rules from JSON."""
    path = ensure_rules_config(path)
    rules = json.loads(path.read_text(encoding="utf-8"))
    _merge_rule_sets(rules, STANDARD_RULE_UPDATES)

    for standard_item, rule in rules.items():
        rules[standard_item] = enrich_rule_metadata(standard_item, rule)
        rule = rules[standard_item]
        if rule.get("quote_type") not in QUOTE_TYPES:
            raise ValueError(f"{standard_item} 的报价类型无效：{rule.get('quote_type')}")
        source_status = rule.get("default_source_status", "方案明确")
        if source_status not in SOURCE_STATUSES:
            raise ValueError(f"{standard_item} 的来源状态无效：{source_status}")

    return rules


def find_alias_matches(text: str, rules: dict[str, dict[str, Any]]) -> list[dict[str, Any]]:
    """Find alias matches and avoid overlapping shorter aliases."""
    raw_matches: list[dict[str, Any]] = []
    priority = {item: index for index, item in enumerate(PRIORITY_STANDARD_ITEMS)}

    for standard_item, rule in rules.items():
        for alias in rule.get("aliases", []):
            if not alias:
                continue

            start = text.find(alias)
            while start != -1:
                raw_matches.append(
                    {
                        "standard_item": standard_item,
                        "matched_text": alias,
                        "start": start,
                        "end": start + len(alias),
                        "category": rule.get("category", ""),
                        "quote_type": rule.get("quote_type", "固定单价"),
                        "default_unit": rule.get("default_unit", "项"),
                        "default_desc": rule.get("default_desc", ""),
                        "default_source_status": rule.get("default_source_status", "方案明确"),
                        "priority": priority.get(standard_item, len(priority)),
                    }
                )
                start = text.find(alias, start + len(alias))

    raw_matches.sort(key=lambda item: (item["start"], -(item["end"] - item["start"]), item["priority"]))
    accepted: list[dict[str, Any]] = []
    occupied: list[tuple[int, int]] = []

    for match in raw_matches:
        span = (match["start"], match["end"])
        overlaps = any(span[0] < old_end and span[1] > old_start for old_start, old_end in occupied)
        if overlaps:
            continue
        accepted.append(match)
        occupied.append(span)

    return accepted
