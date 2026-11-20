"""开项科目表 —— 25 个一级科目 / 约 139 个子项的结构定义与读取。

数据来源：提资材料《估算开项数据采集要求-20261016.xlsx》的「项目名录」表。
这张表是投资匡算的科目来源（讲稿：「匡算不是拍一个总数，是一张科目表」），
每个子项要填一个面积/长度/数量，这个量由无人机采集成果与业主资料提供。

三类改造的一级科目：
  老旧街区 10 个 / 老旧厂区 10 个 / 老旧小区 5 个，合计 25 个。
"""

from __future__ import annotations

import re
from pathlib import Path

# 子项目类型（决定该科目属于保底/完善/提升，对应投资分档）
SUBTYPES = ("基础类改造", "完善类改造", "提升类改造")

# 三类改造 → 一级科目（序号 + 名称），与《估算开项数据采集要求》一致
SUBJECTS = {
    "老旧街区": [
        ("一", "街巷安全隐患整治"),
        ("二", "市政管网全域更新"),
        ("三", "街巷基础环境整治"),
        ("四", "慢行交通系统优化"),
        ("五", "便民公共服务配套"),
        ("六", "街巷环境品质提升"),
        ("七", "历史风貌保护与建筑活化"),
        ("八", "街区业态升级与场景再造"),
        ("九", "特色街巷景观与风貌塑造"),
        ("十", "智慧街区与绿色低碳改造"),
    ],
    "老旧厂区": [
        ("十一", "建构筑物安全治理"),
        ("十二", "场地环境风险治理"),
        ("十三", "厂区市政管网更新"),
        ("十四", "厂区道路交通系统"),
        ("十五", "厂房建筑改造再利用"),
        ("十六", "拆除新建"),
        ("十七", "产业与公共配套"),
        ("十八", "景观环境改造"),
        ("十九", "工业文化遗存保护利用"),
        ("二十", "智能化与绿色低碳改造"),
    ],
    "老旧小区": [
        ("二十一", "建筑本体改造"),
        ("二十二", "小区公共基础设施"),
        ("二十三", "配套设施补齐"),
        ("二十四", "环境景观提升"),
        ("二十五", "智能化改造"),
    ],
}

# 关键子项关键词 → 量纲（用于从采集成果里定位工程量）
# 只登记口径，不含任何单价（单价必须来自可研/定额，见财务参数表）
UNIT_HINTS = (
    ("面积", ("面积", "㎡", "m2", "平方米")),
    ("长度", ("长度", "管线", "米", "m ", "延米")),
    ("数量", ("数量", "套", "盏", "个", "处", "台", "座", "间", "处")),
)


def locate_table(workspace):
    """在工作区输入文件里找《估算开项数据采集要求》表。"""
    root = Path(workspace) / "输入文件"
    if not root.is_dir():
        return None
    for path in sorted(root.rglob("*.xlsx")):
        if "开项" in path.name or "数据采集要求" in path.name:
            return path
    return None


def read_catalog(path):
    """读「项目名录」表 → {改造类型: [{seq, name, subtype, items:[{name, unit_hint}]}]}。"""
    try:
        import openpyxl
    except ImportError:
        return None
    try:
        workbook = openpyxl.load_workbook(path, data_only=True)
    except Exception:
        return None
    if "项目名录" not in workbook.sheetnames:
        return None
    sheet = workbook["项目名录"]
    catalog = {}
    current_type = None
    current_subtype = None
    current_subject = None
    for row in sheet.iter_rows(min_row=2, values_only=True):
        values = list(row) + [None] * 5
        ptype, psub, seq, name, note = values[:5]
        if ptype:
            current_type = str(ptype).strip()
            catalog.setdefault(current_type, [])
        if psub:
            current_subtype = str(psub).strip()
        if not seq or not name:
            continue
        seq, name = str(seq).strip(), str(name).strip()
        if not current_type:
            continue
        # 一级科目：序号为中文数字
        if not re.fullmatch(r"\d+", seq):
            current_subject = {"seq": seq, "name": name, "subtype": current_subtype, "items": []}
            catalog[current_type].append(current_subject)
            continue
        if current_subject is None:
            continue
        current_subject["items"].append({
            "name": name,
            "unit_hint": str(note).strip() if note else "",
        })
    return catalog


def unit_for(item):
    """从子项备注推断量纲关键词（面积/长度/数量），推断不出返回空串。"""
    hint = (item.get("unit_hint") or "") + " " + (item.get("name") or "")
    for unit, keys in UNIT_HINTS:
        if any(key in hint for key in keys):
            return unit
    return ""
