"""写出 财务参数.xlsx（9 张表 + 出处台账），**保留用户已填的格子**。

幂等要求（实测踩过）：重复运行时表里不能出现「表头+数据」重复两遍
（曾导致工程费翻倍 → 总投资算成 10600.11）。
因此写入流程是：**读现有表 → 与新抽到的值合并（用户值优先）→ 清空重写**，
而不是直接 append。

单独成模块，便于 build_params.py 保持可读。
"""

from __future__ import annotations

import re
import sys
from decimal import Decimal
from pathlib import Path

# 运行时是 python313._pth 模式，脚本目录不进 sys.path：兄弟模块必须显式加进来
sys.path.insert(0, str(Path(__file__).resolve().parent))

from field_hints import FIELD_HINTS, build_notes  # noqa: E402

SHEETS = ["项目", "工程费", "其他费", "预备费", "建设期利息",
           "收入", "成本", "资金结构", "可研财务指标"]

PROJECT_KEYS = ["项目名称", "项目方向", "改造面积(㎡)", "建设期(年)",
                "运营期(年)", "折现率", "土地使用税计税面积(㎡)"]
INTEREST_KEYS = ["借款额(万元)", "年利率", "建设期年数", "半额计息", "建设期融资费用(万元)"]
COST_KEYS = ["班组人数", "人均年薪酬(万元)", "福利系数", "年增长率"]
FINANCE_KEYS = ["资本金(万元)", "借款(万元)", "补助(万元)", "借款年利率",
                "借款期限(年)", "宽限期(年)", "还款方式", "财务评价投资基数(万元)",
                "专项债申报上限(万元)", "资金申报口径合计(万元)"]
METRIC_KEYS = ["全部投资FIRR(税后)", "全部投资FNPV(税后)", "全部投资回收期",
               "资本金FIRR(税后)", "资本金FNPV(税后)", "资本金回收期", "FNPV折现率"]

# 三列式表（第一列是编号 E1/F1/R1，名称在第二列，金额在第三列）
_ROW_TABLES = ("工程费", "其他费", "收入")


def read_existing(workbook):
    """读现有 财务参数.xlsx 的用户填写内容：{表名: {键: 值}}。

    两列式表（项目/预备费/成本/资金结构/建设期利息/可研财务指标）取 第一列→第二列；
    三列式表（工程费/其他费/收入）取 名称→金额。
    """
    out = {}
    for name in SHEETS:
        if name not in workbook.sheetnames:
            continue
        worksheet = workbook[name]
        rows = list(worksheet.iter_rows(values_only=True))
        if not rows:
            continue
        bucket = {}
        rows_map = {}
        for row in rows[1:]:
            if not row or len(row) < 2:
                continue
            key = row[0]
            if key is None:
                continue
            key = str(key).strip()
            if not key or key in ("编号", "项目", "指标"):
                continue
            if name in _ROW_TABLES:
                # 数据行第一格是 E1/F1/R1 这类编号；表头行第一格是「编号」
                if not re.fullmatch(r"[A-Z]+\d+", key):
                    continue
                item_name = row[1] if len(row) > 1 else None
                amount = row[2] if len(row) > 2 else None
                if item_name and amount is not None:
                    rows_map[str(item_name).strip()] = amount
            else:
                value = row[1]
                if value is not None and value != "":
                    bucket[key] = value
        if rows_map:
            bucket["_rows"] = rows_map
        out[name] = bucket
    return out


def write_params(path, table, provenance, notes):
    """写出 财务参数.xlsx（9 张表 + 出处台账 + 填写说明）。"""
    try:
        import openpyxl
        from openpyxl.styles import Font, PatternFill
    except ImportError:
        raise SystemExit("[01b] 缺 openpyxl，无法写出财务参数.xlsx")

    if path.is_file():
        workbook = openpyxl.load_workbook(path)
        existing = read_existing(workbook)
    else:
        workbook = openpyxl.Workbook()
        workbook.remove(workbook.active)
        existing = {}

    head_font = Font(bold=True)
    head_fill = PatternFill("solid", fgColor="DDDDDD")

    def sheet(name):
        if name in workbook.sheetnames:
            return workbook[name]
        return workbook.create_sheet(name)

    def hint(sheet_name, label):
        for field, text in FIELD_HINTS.get(sheet_name, []):
            if field == label:
                return text
        return None

    def mark(worksheet, row_index, col_index, text):
        """给空值格子挂填写指引批注（点格子即见）。"""
        if not text:
            return
        try:
            from openpyxl.comments import Comment
            worksheet.cell(row=row_index, column=col_index).comment = Comment(
                "填写指引：{}".format(text), "01b-可研财务提取")
        except Exception:
            pass

    def reset(worksheet, header):
        """清空工作表并写新表头（保证重复运行不叠加）。"""
        if worksheet.max_row:
            worksheet.delete_rows(1, worksheet.max_row)
        worksheet.append(header)
        for cell in worksheet[1]:
            cell.font = head_font
            cell.fill = head_fill

    def user_value(table_name, key):
        return (existing.get(table_name) or {}).get(key)

    def write_kv(name, keys, source_sheet):
        worksheet = sheet(name)
        reset(worksheet, ["项目", "值"])
        for label in keys:
            hit = (table.get(source_sheet) or {}).get(label)
            value = hit["value"] if hit else None
            filled = user_value(name, label)
            if filled is not None and filled != "":
                value = filled          # 用户填的优先，不覆盖
            if isinstance(value, Decimal):
                value = float(value)
            worksheet.append([label, value])
            if value is None or value == "":
                mark(worksheet, worksheet.max_row, 2, hint(name, label))
        worksheet.column_dimensions["A"].width = 24
        worksheet.column_dimensions["B"].width = 28

    def _user_extends(name, header):
        """用户表比模板宽（加了单位/工程量/测算依据列）或补了模板没有的行 → 原样保留。

        背景：01b 的抽取结果只覆盖它认得的科目，重整会把这些科目之外的行删掉。
        实测 2026-11-14：其他费按附表1 的 19 项顶层科目录入后，被压回 10 项，
        勘察设计费/建设管理费重复计列又回来了（总投资虚高的根因）。
        """
        if name not in workbook.sheetnames:
            return False
        worksheet = workbook[name]
        head = next(worksheet.iter_rows(min_row=1, max_row=1, values_only=True), ()) or ()
        if len([cell for cell in head if cell not in (None, "")]) > len(header):
            return True
        known = {str(key).strip() for key in (table.get(name) or {})}
        for row in worksheet.iter_rows(min_row=2, values_only=True):
            if row and len(row) > 1 and row[1] and str(row[1]).strip() not in known:
                return True
        return False

    def write_table(name, header, source_sheet):
        if _user_extends(name, header):
            return
        worksheet = sheet(name)
        reset(worksheet, header)
        index = 0
        for label, hit in (table.get(source_sheet) or {}).items():
            index += 1
            value = hit["value"]
            if isinstance(value, Decimal):
                value = float(value)
            filled = (user_value(name, "_rows") or {}).get(label)
            if filled is not None and filled != "":
                value = filled
            if source_sheet == "工程费":
                worksheet.append(["E{}".format(index), label, value])
                if value is None or value == "":
                    mark(worksheet, worksheet.max_row, 3, hint(name, "金额(万元)"))
            elif source_sheet == "其他费":
                worksheet.append(["F{}".format(index), label, value, None, None])
                if value is None or value == "":
                    mark(worksheet, worksheet.max_row, 3, hint(name, "金额(万元)"))
                mark(worksheet, worksheet.max_row, 4, hint(name, "取费基数(万元)"))
                mark(worksheet, worksheet.max_row, 5, hint(name, "费率"))
            else:
                worksheet.append(["R{}".format(index), label, value, None, None])
                if value is None or value == "":
                    mark(worksheet, worksheet.max_row, 3, hint(name, "年均收入(万元)"))
        for letter in "ABCDE":
            worksheet.column_dimensions[letter].width = 26

    write_kv("项目", PROJECT_KEYS, "项目")
    write_table("工程费", ["编号", "名称", "金额(万元)"], "工程费")
    write_table("其他费", ["编号", "名称", "金额(万元)", "取费基数(万元)", "费率"], "其他费")
    write_kv("预备费", ["预备费率"], "预备费")
    write_kv("建设期利息", INTEREST_KEYS, "建设期利息")
    write_table("收入", ["编号", "名称", "年均收入(万元)", "增值税率", "年增长率"], "收入")
    write_kv("成本", COST_KEYS, "成本")
    write_kv("资金结构", FINANCE_KEYS, "资金结构")

    worksheet = sheet("可研财务指标")
    reset(worksheet, ["指标", "值", "单位", "说明"])
    for label in METRIC_KEYS:
        unit = "%" if ("FIRR" in label or "折现率" in label) else \
            ("年" if "回收期" in label else "万元")
        worksheet.append([label, user_value("可研财务指标", label), unit, "可研附表"])
        if user_value("可研财务指标", label) is None:
            mark(worksheet, worksheet.max_row, 2, hint("可研财务指标", label))
    worksheet.column_dimensions["A"].width = 24

    worksheet = sheet("出处台账")
    reset(worksheet, ["工作表", "字段", "来源文件", "表页/工作表", "行号", "原文片段"])
    for item in provenance:
        source = item.get("source") or {}
        worksheet.append([
            item.get("sheet"), item.get("key"), source.get("file"),
            source.get("sheet") or ("第{}页".format(source.get("page"))
                                    if source.get("page") else ""),
            source.get("row"), source.get("quote"),
        ])
    for letter in "ABCDEF":
        worksheet.column_dimensions[letter].width = 22

    worksheet = sheet("填写说明")
    if worksheet.max_row:
        worksheet.delete_rows(1, worksheet.max_row)
    for line in build_notes():
        worksheet.append(line)
    worksheet.column_dimensions["A"].width = 110

    # 保证 9 张表齐全（缺表补空表，避免 01 取料报「表结构不可用」）
    for name in SHEETS:
        sheet(name)

    path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(path)
