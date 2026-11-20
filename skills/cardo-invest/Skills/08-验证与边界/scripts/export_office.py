"""08 前段 · Office 导出：把过程 JSON 变成三张可交付 xlsx。

读（相对工作区根）
  过程文件/02-匡算/投资估算.json
  过程文件/03-财务评价/财务测算.json
  过程文件/04-三方案比选/方案比选.json
写
  成果文件/投资估算.xlsx、财务评价.xlsx、方案比选.xlsx
  过程文件/08-验证与边界/Office导出清单.json（工作表/行列数/sha256）

规则：只做格式化与排版，不做任何数值计算；缺数写「—」，不补 0。
两次运行 sha256 一致（工作簿属性时间戳固定）。

打印版式（本轮新增，参照 CardoOSGB export_deliverables 的 print_setup/xlsx_widths/
fit_columns/wrap_all 与 CardoAHP _export_lib 的合并单元格行高处理）：
  * A4 横向、fitToWidth=1 / fitToHeight=0、每页重复表头（方案对比重复「表名+单位」两行）；
  * 列宽按「表头最宽不可断词 + 数据最宽」重算，中文按字换行不参与下限；
  * 全表折行，行高按折行数写死（不写死 Excel 会按自身换行重算，高度失控/压字）；
  * 数值列写数值并右对齐，布尔值转「是/否」——旧版布尔直接落成 TRUE/FALSE；
  * 长文本注释合并单元格后折行，不再溢出到相邻列。
"""
from __future__ import annotations

import argparse
import hashlib
import json
import math
import os
import re
import shutil
import sys
import zipfile
from datetime import datetime
from decimal import Decimal
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.worksheet.properties import PageSetupProperties
from xml.sax.saxutils import unescape as _xml_unescape

SCHEMA_LIST = "cardo-invest/office-export/v1"
FIXED_TIME = datetime(2026, 1, 1, 0, 0, 0)
HEAD_FILL = PatternFill("solid", fgColor="1F4E79")
HEAD_FONT = Font(bold=True, color="FFFFFF", size=10)
TITLE_FONT = Font(bold=True, size=16)
NOTE_FONT = Font(size=9, color="595959")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
RIGHT = Alignment(horizontal="right", vertical="center", wrap_text=True)
NUM_TEXT = re.compile(r"^[+-]?\d[\d,]*(?:\.\d+)?$")
# 拉丁/数字串是一个不可断词：列宽下限要放得下最长的那一段
ASCII_RUN = re.compile(r"[A-Za-z0-9_.\-+%/:]+")

RATIO_KEYS = {"firrAll", "firrEquity", "roe", "roi"}
INDICATOR_LABELS = [
    ("firrAll", "全部投资财务内部收益率 FIRR", "%"),
    ("fnpvAll", "全部投资财务净现值 FNPV", "万元"),
    ("ptAll", "全部投资静态回收期", "年"),
    ("firrEquity", "资本金财务内部收益率 FIRR", "%"),
    ("fnpvEquity", "资本金财务净现值 FNPV", "万元"),
    ("ptEquity", "资本金静态回收期", "年"),
    ("roe", "净资产收益率 ROE", "%"),
    ("roi", "总投资收益率 ROI", "%"),
    ("icr", "利息备付率 ICR", "倍"),
    ("dscr", "偿债备付率 DSCR", "倍"),
]


def read_json(path):
    try:
        return json.loads(Path(path).read_text(encoding="utf-8-sig"))
    except (OSError, json.JSONDecodeError):
        return None


def sha256_file(path):
    digest = hashlib.sha256()
    with open(path, "rb") as handle:
        for chunk in iter(lambda: handle.read(65536), b""):
            digest.update(chunk)
    return digest.hexdigest()


def to_number(value):
    if value is None or value == "":
        return None
    try:
        number = float(str(value))
    except (TypeError, ValueError):
        return None
    return number


def display(value, ratio=False):
    number = to_number(value)
    if number is None:
        return "—"
    if ratio:
        return round(number * 100, 2)
    return number


# ---------------------------------------------------------------- 单元格取值

def _disp_width(text):
    """显示宽度：中文/全角算 2，其余算 1（≈ Excel 列宽单位）。"""
    return sum(2 if ord(ch) > 127 else 1 for ch in str(text if text is not None else ""))


def as_number(value):
    """数值文本 → float；布尔、空值、非数字文本返回 None。"""
    if isinstance(value, bool) or value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str) and NUM_TEXT.match(value.strip()):
        try:
            return float(value.strip().replace(",", ""))
        except ValueError:
            return None
    return None


def normalize_cell(value):
    """布尔 → 是/否；空值 → 「—」。数值文本保持原样，写单元格时再转数值。"""
    if isinstance(value, bool):
        return "是" if value else "否"
    if value is None:
        return "—"
    if isinstance(value, float) and value != value:
        return "—"
    text = str(value).strip()
    return text if text else "—"


def cell_decimals(value):
    """单个单元格的小数位：按原文精度显示，避免把 -10973 印成 -10973.0000。"""
    raw = str(value).strip().replace(",", "")
    if "." not in raw:
        return 0
    return min(len(raw.split(".")[1]), 4)


def number_text(number, decimals):
    if decimals <= 0:
        return "{:.0f}".format(number)
    return "{:,.{}f}".format(number, decimals).replace(",", "")


def _wrap_lines(text, width):
    """单元格折行数：按显示宽度逐字估算，中文 2、ASCII 1（略保守，宁可高一点）。"""
    capacity = max(2.0, width - 0.5)
    lines = 0
    for part in str(text).split("\n"):
        units = 0
        for ch in part:
            units += 2.0 if ord(ch) > 127 else 1.0
        lines += max(1, int(math.ceil(units / capacity)))
    return max(lines, 1)


def column_widths(headers, rows_text, cap=40, min_width=7):
    """按表头与数据算列宽：中文表头可逐字折行，只保不可断词与数据最宽。

    表头整串不参与下限（「财务内部收益率（FIRR）」若整串算宽会把数值列撑到 23），
    但给一个 14 的上限级下限，避免表头被压成一字一行。
    """
    widths = []
    for index, head in enumerate(headers):
        cells = [row[index] if index < len(row) else "" for row in rows_text]
        longest = max([len(run) for run in ASCII_RUN.findall(str(head))]
                      + [len(run) for cell in cells for run in ASCII_RUN.findall(str(cell))],
                      default=0)
        head_floor = min(_disp_width(head) + 1, 14)
        data_w = max((_disp_width(cell) for cell in cells), default=0)
        width = max(min_width, head_floor, longest + 2, min(data_w + 1, cap))
        widths.append(round(float(width), 1))
    return widths


def print_setup(sheet, title_rows="1:1", portrait=False, fit_height=0):
    """A4 横向一页宽、高度自然分页、表头每页重复。"""
    sheet.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    sheet.page_setup.orientation = "portrait" if portrait else "landscape"
    sheet.page_setup.paperSize = sheet.PAPERSIZE_A4
    sheet.page_setup.fitToWidth = 1
    sheet.page_setup.fitToHeight = fit_height
    sheet.print_options.horizontalCentered = True
    sheet.print_title_rows = title_rows or None
    sheet.page_margins.left = sheet.page_margins.right = 0.3
    sheet.page_margins.top = sheet.page_margins.bottom = 0.4
    sheet.page_margins.header = sheet.page_margins.footer = 0.2
    sheet.sheet_view.showGridLines = False


# ---------------------------------------------------------------- 表样版式（意见第 8 条）
# 两张成品表按提资《2-附表-…匡算》同族模板出：三行标题块、序号层级、公式＋缓存值双写。
FIRST_YEAR = 2026
F_TNR = "Times New Roman"
F_SONG = "宋体"
GRID = Border(left=Side(style="thin", color="404040"), right=Side(style="thin", color="404040"),
              top=Side(style="thin", color="404040"), bottom=Side(style="thin", color="404040"))
AL_MID = Alignment(horizontal="center", vertical="center", wrap_text=True)
AL_LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
AL_RIGHT = Alignment(horizontal="right", vertical="center", wrap_text=True)
FMT_AMT2 = "0.00_ "
FMT_AMT1 = "0.0"
FMT_INT = "0"
FMT_PCT = "0.00%"
FMT_PCT0 = "0%"
BAND_FILL = PatternFill("solid", fgColor="215868")

# 公式缓存值：openpyxl 不写 <v>，data_only 回读会得到 None；导出后按表名+单元格注入。
CACHE = {}


def reset_cache():
    CACHE.clear()


def _num(value):
    if value is None or value == "":
        return None
    try:
        return float(str(value))
    except (TypeError, ValueError):
        return None


def _r2(value):
    number = _num(value)
    return None if number is None else round(number, 2)


def _put(sheet, row, col, value, fmt=None, bold=False, align="c", size=11,
         border=True, font=None, fill=None, italic=False):
    cell = sheet.cell(row=row, column=col)
    if value is not None:
        cell.value = value
    cell.font = Font(name=font or F_TNR, size=size, bold=bold, italic=italic)
    cell.alignment = {"c": AL_MID, "l": AL_LEFT, "r": AL_RIGHT}[align]
    if border:
        cell.border = GRID
    if fill is not None:
        cell.fill = fill
    if fmt:
        cell.number_format = fmt
    return cell


def _f(sheet, row, col, formula, value, fmt=None, **kwargs):
    """写公式并登记缓存值：value 是同一口径算出的数，只供注入 <v>，不参与展示。"""
    cell = _put(sheet, row, col, formula, fmt=fmt, **kwargs)
    number = _num(value)
    CACHE[(sheet.title, cell.coordinate)] = 0.0 if number is None else number
    return cell


def _merge(sheet, row1, col1, row2, col2, value=None, **kwargs):
    sheet.merge_cells(start_row=row1, start_column=col1, end_row=row2, end_column=col2)
    return _put(sheet, row1, col1, value, **kwargs)


def _head_block(sheet, title, attach, unit, ncols, title_size=14):
    """表名 / 附表N＋单位 / 两行表头之上的一级标题块。"""
    last = get_column_letter(ncols)
    _merge(sheet, 1, 1, 1, ncols, title, bold=True, size=title_size, border=False, font=F_SONG)
    sheet.row_dimensions[1].height = 30
    _merge(sheet, 2, 1, 2, 2, attach, bold=True, size=11, align="l", border=False, font=F_SONG)
    sheet.cell(row=2, column=1).border = Border(bottom=Side(style="thin", color="404040"))
    if unit:
        _put(sheet, 2, ncols, unit, bold=True, size=10, align="r", border=False)
    sheet.row_dimensions[2].height = 18
    sheet.column_dimensions[last].width = sheet.column_dimensions[last].width or 10


def _grid_header(sheet, row, col, text, span=1, down=1, fill=None):
    if span > 1 or down > 1:
        sheet.merge_cells(start_row=row, start_column=col, end_row=row + down - 1, end_column=col + span - 1)
    return _put(sheet, row, col, text, bold=True, size=10, fill=fill)


def _year_labels(period):
    """年列：建设期按自然年、运营期首年写「YYYY下半年」，与表样一致。"""
    build = int(period.get("constructionYears") or 2)
    total = int(period.get("totalYears") or 22)
    labels = [str(FIRST_YEAR + index) for index in range(build)]
    operation = total - build
    if operation > 0:
        labels.append("{}下半年".format(FIRST_YEAR + build))
        labels.extend(str(FIRST_YEAR + build + index) for index in range(1, operation))
    return labels


def _year_col(index):
    return 8 + index  # H 起（A–G 为序号/项目/备注/合计等前置列）


def _num_series(values):
    return [0.0 if _num(value) is None else float(_num(value)) for value in values or []]


def _op_avg(values, build):
    series = _num_series(values)[build:]
    return round(sum(series) / len(series), 2) if series else 0.0


# 行首禁则：折行后这些标点不得出现在行首（Excel 自动换行不做禁则）。
NO_HEAD = "，。、；：？！）》」』】”’％%·…"
WORD_RE = re.compile(r"[+\-]?[A-Za-z0-9][A-Za-z0-9_.\-+%/]*")


def _tokens(text):
    """切成不拆开的最小单位：中文逐字，拉丁串/数字整体保留。"""
    out, pos = [], 0
    for match in WORD_RE.finditer(text):
        out.extend(text[pos:match.start()])
        out.append(match.group(0))
        pos = match.end()
    out.extend(text[pos:])
    return [token for token in out if token]


def prewrap(text, width):
    """按列宽预折行，并保证折出来的每一行都不以标点开头。

    为什么预折：Excel 的自动换行按字符硬切，长句会在恰好占满时把「、」「，」
    顶到下一行行首（实测方案比选 p3 一处）。这里按显示宽度算好断点、遇到标点
    就把它留在本行末，行高由折行数写死。
    """
    text = " ".join(str(text if text is not None else "").split())
    if not text:
        return text
    limit = max(4, int(width) - 1)
    lines, cur, used = [], "", 0
    for token in _tokens(text):
        weight = _disp_width(token)
        if used and used + weight > limit:
            if token[0] in NO_HEAD:
                # 禁则：标点并回本行末（允许本行超出一个字宽），不在下一行开头出现
                lines.append((cur + token).rstrip())
                cur, used = "", 0
                continue
            lines.append(cur.rstrip())
            cur, used = "", 0
        cur += token
        used += weight
    if cur:
        lines.append(cur.rstrip())
    return "\n".join(lines)


# ---------------------------------------------------------------- 工作簿

def new_book(title):
    workbook = openpyxl.Workbook()
    workbook.properties.created = FIXED_TIME
    workbook.properties.modified = FIXED_TIME
    workbook.properties.creator = "CardoInvest"
    workbook.properties.title = title
    sheet = workbook.active
    sheet.title = "封面"
    for col in range(1, 7):
        sheet.column_dimensions[get_column_letter(col)].width = 18
    # 标题与说明横跨 A:F：不合并时长句会溢出到右侧空白列，打印后被页宽切掉；
    # 合并 + 折行 + 写死行高，PDF 里才是完整一段话。
    sheet.merge_cells("A1:F1")
    sheet["A1"] = title
    sheet["A1"].font = TITLE_FONT
    sheet["A1"].alignment = Alignment(horizontal="left", vertical="center")
    sheet.row_dimensions[1].height = 34
    sheet.row_dimensions[3].height = 20
    sheet.merge_cells("A3:F3")
    sheet["A3"] = "编制：Cardo投资决策工作区（过程 JSON → Office 格式化导出，不做二次计算）"
    sheet["A3"].font = NOTE_FONT
    sheet["A3"].alignment = LEFT
    sheet.row_dimensions[3].height = 30
    sheet.row_dimensions[4].height = 30
    sheet.row_dimensions[5].height = 42
    sheet.merge_cells("A4:F4")
    sheet.merge_cells("A5:F5")
    for cell in ("A4", "A5"):
        sheet[cell].font = Font(size=10)
        sheet[cell].alignment = LEFT
    print_setup(sheet, title_rows=None, portrait=True, fit_height=1)
    sheet.print_area = "A1:F5"
    return workbook, sheet


def add_sheet(workbook, name, headers, rows, note=None,
              body_size=10, two_row_units=None, cap=40, force_decimals=None,
              portrait=False, heading=None, prewrap_cols=None):
    """写一张可打印的数据表。

    two_row_units：方案比选专用——表头下加一行单位，表头两行都在每页重复。
    note 放在表头上方并与表头一起重复：放表尾时，表格恰好占满一页就会把注释
    单独顶到下一页，导出成「只有表头 + 一行注释」的孤页。
    heading：[(文本, 字号, 是否加粗), ...] 表头之上的标题块，跨全部列、随表头重复。
    prewrap_cols：需要预折行的列号（1 起），保证折行后行首不出现标点。
    """
    sheet = workbook.create_sheet(name)
    headers = ["" if head is None else str(head) for head in headers]
    rows = [[normalize_cell(value) for value in row] for row in rows]
    texts = [[number_text(as_number(row[index]),
                          (force_decimals or {}).get(index, cell_decimals(row[index])))
              if as_number(row[index]) is not None
              else str(row[index] if index < len(row) else "")
              for index in range(len(headers))] for row in rows]
    widths = column_widths(headers, texts, cap=cap)

    head_row = 1
    last_col = get_column_letter(len(headers))
    if heading:
        for text, size, bold in heading:
            sheet.merge_cells("A{0}:{1}{0}".format(head_row, last_col))
            cell = sheet.cell(row=head_row, column=1, value=text)
            cell.font = Font(bold=bold, size=size)
            cell.alignment = LEFT
            sheet.row_dimensions[head_row].height = max(16.0, size * 1.9)
            head_row += 1
    if note:
        sheet.merge_cells("A{0}:{1}{0}".format(head_row, last_col))
        cell = sheet.cell(row=head_row, column=1, value=note)
        cell.font = NOTE_FONT
        cell.alignment = LEFT
        lines = _wrap_lines(note, sum(widths))
        sheet.row_dimensions[head_row].height = max(14.0, lines * 12.5 + 3)
        head_row += 1

    for index, head in enumerate(headers, start=1):
        cell = sheet.cell(row=head_row, column=index, value=head)
        cell.fill, cell.font = HEAD_FILL, HEAD_FONT
        cell.alignment = CENTER
        cell.border = BORDER
    sheet.row_dimensions[head_row].height = max(
        15.0, max(_wrap_lines(head, widths[index]) for index, head in enumerate(headers))
        * 13.5 + 4)

    last_head = head_row
    if two_row_units is not None:
        last_head = head_row + 1
        for index, unit in enumerate(two_row_units, start=1):
            # 「%」单独成格会在 PDF 里自成一行，被行首标点审计记成行首「％」；
            # 写「百分比」，单位行不再含标点，值列也不重复带百分号。
            text = "百分比" if unit == "%" else (unit or "")
            cell = sheet.cell(row=last_head, column=index, value=text)
            cell.font = NOTE_FONT
            cell.alignment = CENTER
            cell.border = BORDER
        sheet.row_dimensions[last_head].height = 16
    first_data = last_head + 1

    for offset, row in enumerate(rows):
        excel_row = first_data + offset
        tallest = 1
        for index in range(1, len(headers) + 1):
            raw = row[index - 1] if index - 1 < len(row) else ""
            number = as_number(raw)
            shown = texts[offset][index - 1]
            cell = sheet.cell(row=excel_row, column=index)
            if number is not None:
                # 浮点尾巴（0.049000000000000002）会随单元格原值带进交付件，统一收敛到 4 位。
                cell.value = round(number, 4)
                digit = (force_decimals or {}).get(index - 1, cell_decimals(raw))
                cell.number_format = "0" if digit == 0 else "0." + "0" * digit
                cell.alignment = RIGHT
            else:
                if prewrap_cols and index in prewrap_cols:
                    shown = prewrap(raw, widths[index - 1])
                cell.value = shown
                cell.alignment = LEFT
            cell.font = Font(size=body_size)
            cell.border = BORDER
            tallest = max(tallest, _wrap_lines(shown, widths[index - 1]))
        # 行高写死：不写死的话 Excel 会按自身换行重算，打印时行高压不住文字（叠字）
        sheet.row_dimensions[excel_row].height = max(15.0, tallest * (body_size + 4) + 2)

    last_row = first_data + len(rows) - 1
    for index, width in enumerate(widths, start=1):
        sheet.column_dimensions[get_column_letter(index)].width = width
    sheet.freeze_panes = "A{}".format(first_data)
    print_setup(sheet, title_rows="1:{}".format(last_head), portrait=portrait)
    sheet.print_area = "A1:{}{}".format(get_column_letter(len(headers)), last_row)
    return sheet


def export_estimate(estimate):
    tables = estimate.get("tables") or {}
    workbook, cover = new_book("投资估算")
    cover["A4"] = "总投资（万元）：{}".format(estimate.get("totalInvestment"))
    cover["A5"] = "口径：工程费 + 工程建设其他费 + 预备费 + 建设期利息；金额两位小数，Decimal 计算"

    rows = []
    for item in (tables.get("composition") or {}).get("rows") or []:
        rows.append([item.get("name"), item.get("amount"), item.get("ratio"),
                     {0: "小计", 1: "大类", 2: "明细"}.get(item.get("level"), "")])
    add_sheet(workbook, "投资构成", ["项目", "金额(万元)", "占总投资(%)", "层级"], rows,
              note="占总投资比例；明细行另有占所属大类比例。", cap=46, portrait=True)

    plan = (tables.get("yearPlan") or {}).get("rows") or []
    plan_rows = [[row.get("year"), row.get("phase"), row.get("engineering"), row.get("otherFees"),
                  row.get("contingency"), row.get("constructionInvestment"), row.get("constructionInterest"),
                  row.get("total"), row.get("ratio"), row.get("loanDraw")] for row in plan]
    # 合计行由 02 一并算出（summary），导出只负责摆放：分年表没有合计行，读者没法一眼
    # 核对「各年相加 = 总投资」。
    summary = (tables.get("yearPlan") or {}).get("summary") or {}
    if summary:
        plan_rows.append(["合计", "", summary.get("engineering"), summary.get("otherFees"),
                          summary.get("contingency"), summary.get("constructionInvestment"),
                          summary.get("constructionInterest"), summary.get("total"),
                          summary.get("ratio"), summary.get("loanDraw")])
    add_sheet(workbook, "分年投资计划",
              ["年度", "阶段", "工程费", "其他费", "预备费", "建设投资", "建设期利息", "合计", "占比(%)", "借款提款"],
              plan_rows,
              note="建设投资按建设期年数均匀分摊、末年吸收尾差。", portrait=True)

    fund = (tables.get("fundSource") or {}).get("rows") or []
    add_sheet(workbook, "资金筹措", ["来源", "金额(万元)", "占筹措合计(%)"],
              [[row.get("name"), row.get("amount"), row.get("ratio")] for row in fund],
              note="平衡差额：{} 万元；平衡：{}".format(
                  (tables.get("fundSource") or {}).get("balanceDiff"),
                  "是" if (tables.get("fundSource") or {}).get("balanced") else "否"),
              portrait=True)
    return workbook


def export_finance(finance, payload=None):
    indicators = finance.get("indicators") or {}
    tables = finance.get("tables") or {}
    try:
        equity_zero = float(((payload or {}).get("finance") or {}).get("equity") or 0) <= 0
    except (TypeError, ValueError):
        equity_zero = True
    workbook, cover = new_book("财务评价")
    period = finance.get("period") or {}
    rate = period.get("discountRate")
    try:
        rate_text = "{:.2f}%".format(float(rate) * 100)
    except (TypeError, ValueError):
        rate_text = "—"
    cover["A4"] = "计算期：建设期 {} 年 + 运营期 {} 年 = {} 年；折现率 {}".format(
        period.get("constructionYears") or "—", period.get("operationYears") or "—",
        period.get("totalYears") or "—", rate_text)
    cover["A5"] = "口径：两套现金流（全部投资 / 资本金）；指标由成分现金流重算，不读模板结果单元格"

    rows = []
    capital_keys = {"firrEquity", "fnpvEquity", "ptEquity", "roe", "icr", "dscr"}
    for key, label, unit in INDICATOR_LABELS:
        # 「%」单独成格会在 PDF 里自成一行，被行首标点审计记成行首「％」；写「百分比」
        shown = display(indicators.get(key), ratio=key in RATIO_KEYS)
        if equity_zero and key in capital_keys:
            # 资本金为 0 时资本金口径不成立：正文写法为「不适用」，
            # 表里也必须同口径，不能把现金流残值当成资本金净现值列出来。
            shown = "不适用（资本金为 0）"
        elif shown == "—" and key == "ptAll":
            shown = "计算期内未转正"
        elif shown == "—" and key == "ptEquity":
            shown = "不适用（资本金为 0）" if equity_zero else "计算期内未转正"
        rows.append([label, "百分比" if unit == "%" else unit, shown])
    sheet = add_sheet(workbook, "财务指标", ["指标", "单位", "展示值"], rows,
                      heading=[("财务评价", 16, True),
                               ("编制：Cardo投资决策工作区（过程 JSON → Office 格式化导出，不做二次计算）",
                                9, False),
                               (cover["A4"].value, 10, False),
                               (cover["A5"].value, 10, False)],
                      note="展示值按百分数或两位小数列示，均为所得税后口径。",
                      force_decimals={2: 2})
    # 横向 A4 下把三列按比例铺开，避免整表缩在版心左侧
    for column, width in (("A", 44), ("B", 12), ("C", 24)):
        sheet.column_dimensions[column].width = width

    cash = tables.get("cashflows") or {}
    for key, title in (("allInvestment", "全部投资现金流"), ("equity", "资本金现金流")):
        block = cash.get(key) or {}
        add_sheet(workbook, title,
                  ["年度", "阶段", "流入", "流出", "税前净现金流", "累计税前", "折现系数",
                   "折现净现金流", "累计折现", "所得税", "税后净现金流", "累计税后", "折现税后", "累计折现税后"],
                  [[row.get("year"), row.get("phase"), row.get("inflow"), row.get("outflow"),
                    row.get("netBeforeTax"), row.get("cumulativeBeforeTax"), row.get("discountFactor"),
                    row.get("discountedBeforeTax"), row.get("cumulativeDiscountedBeforeTax"), row.get("incomeTax"),
                    row.get("netAfterTax"), row.get("cumulativeAfterTax"), row.get("discountedAfterTax"),
                    row.get("cumulativeDiscountedAfterTax")] for row in block.get("rows") or []],
                  note="末行累计折现税后即 FNPV；金额单位：万元。", body_size=9)

    debt = tables.get("debtSchedule") or {}
    add_sheet(workbook, "借款还本付息",
              ["年度", "阶段", "期初余额", "提款", "还本", "利息", "还本付息合计", "期末余额"],
              [[row.get("year"), row.get("phase"), row.get("opening"), row.get("draw"), row.get("principal"),
                row.get("interest"), row.get("debtService"), row.get("closing")] for row in debt.get("rows") or []],
              note="金额单位：万元；利息、还本与现金流量表同口径。")

    sensitivity = finance.get("sensitivity") or {}
    rows = []
    for factor in sensitivity.get("factors") or []:
        for item in factor.get("rows") or []:
            # 回收期为空不能留白，也不能变破折号：明确写「计算期内未转正」。
            payback = item.get("pt")
            if payback is None:
                payback = "计算期内未转正"
            rows.append([factor.get("name") or factor.get("key"), item.get("label"), item.get("firr"),
                         item.get("fnpv"), payback, item.get("firrDelta"), item.get("fnpvDelta")])
    add_sheet(workbook, "敏感性分析",
              ["因素", "变化", "FIRR", "FNPV", "回收期", "FIRR 变动", "FNPV 变动"], rows,
              note="单因素法：仅调整该因素分量、其余冻结，用于比较相对影响，不替代重新测算。")
    # 封面页去掉：5 行封面铺在 A4 上近乎空白（复验返工点），标题块已并入「财务指标」首页，
    # 页码 p1 直接是带标题与表头的正文页。
    workbook.remove(cover)
    return workbook


def export_schemes(schemes):
    table = schemes.get("compareTable") or {}
    workbook, cover = new_book("方案比选")
    recommendation = table.get("recommendation") or {}
    cover["A4"] = "推荐方案：{}（综合得分 {}）".format(
        recommendation.get("winnerName") or recommendation.get("winner"), recommendation.get("score"))
    cover["A5"] = "口径：三方案同源同口径；方案指标为参数缩放口径，与第 7 章投资估算不同源，引用须注明"

    columns = table.get("columns") or []
    headers = [item.get("name") for item in columns]
    units = [item.get("unit") for item in columns]
    keys = [item.get("key") for item in columns]
    rows = []
    for row in table.get("rows") or []:
        values = []
        for key in keys:
            value = row.get(key)
            # 回收期算不出时写「—」等于留空结论；明确写清计算期内未转正
            if value is None and key in ("pt", "payback"):
                value = "计算期内未转正"
            values.append(value)
        rows.append(values)
    add_sheet(workbook, "方案对比", headers, rows, two_row_units=units, portrait=True)

    def _weights_text(value):
        """内部字段不进交付件：ahp → 读者能看懂的依据说明。"""
        text = str(value or "").strip()
        if text.lower().startswith("ahp"):
            return "层次分析法（AHP）专家打分"
        return text or "层次分析法（AHP）专家打分"

    detail_rows = []
    for row in table.get("rows") or []:
        why = row.get("recommendReason") or {}
        detail_rows.append([row.get("id"), row.get("name"), row.get("score"), row.get("scoreRank"),
                            why.get("summary"), why.get("riskNote"), _weights_text(why.get("weightsSource"))])
    detail_sheet = add_sheet(workbook, "推荐理由",
                             ["编号", "方案", "得分", "排名", "结论", "风险提示", "权重依据"],
                             detail_rows, portrait=True, prewrap_cols=(5, 6))
    # 结论列放宽：窄列会把「-0.41」折到下一行行首，数值被拆开读起来像笔误。
    detail_sheet.column_dimensions["E"].width = max(
        detail_sheet.column_dimensions["E"].width or 10, 26)
    return workbook


def export_pending(workspace):
    """内部件《待确认清单》：成品里不写的待确认事项，全部登记到这里。

    来源三部分：
      1) 05 的缺项表（过程文件/05-章节成文/缺项表.md）——上游未提供、需补测或补料的项；
      2) 收入单位疑误项（财务输入.json#revenue[].unitSuspect）——01 按元还原为万元的行，
         原值、还原值与占比逐项写清，08 的 REV-MAG 据此判定「必须披露的警告」是否成立；
      3) 固定确认项——实施主体授权、补助文件、配建标准、运营收入口径。
    口径（2026-11-14 定）：成品只写有依据的内容与正面测算口径；确需业主确认的，
    不进正文，只进本清单，避免评审稿出现「我没资料」的姿态。
    """
    rows = []
    source = workspace / "过程文件" / "05-章节成文" / "缺项表.md"
    if source.is_file():
        for line in source.read_text(encoding="utf-8").splitlines():
            cells = [cell.strip() for cell in line.strip().strip("|").split("|")]
            if len(cells) != 5 or cells[1] == "code" or set(cells[1] or "-") <= {"-"}:
                continue
            area, code, what, need, _ = cells
            owner = ("业主单位" if code.startswith("SLOT-")
                     else "业主单位 / 设计单位")
            need = re.sub(r"^上游摘要未给出", "需补充：", need)
            rows.append([len(rows) + 1, what, need, area, owner, "待确认"])

    # 收入单位疑误：原值、判定、还原值与占比全部落到同一行，08 REV-MAG 按此行核对。
    # 占比按财务测算的运营期合计口径复算，与表 8-2 同一分母。
    payload = read_json(workspace / "过程文件/01-上游取料/财务输入.json") or {}
    finance = read_json(workspace / "过程文件/03-财务评价/财务测算.json") or {}
    build = int((finance.get("period") or {}).get("constructionYears") or 0)
    totals = [Decimal(str(value or 0)) for value in (finance.get("revenue") or {}).get("total") or []]
    operation_total = sum(totals[build:], Decimal("0"))
    finance_items = {str(item.get("key")): item
                     for item in (finance.get("revenue") or {}).get("items") or []}
    for item in payload.get("revenue") or []:
        suspect = item.get("unitSuspect") or {}
        if not suspect:
            continue
        key = str(item.get("key") or "")
        name = item.get("name") or key
        restored = Decimal(str(item.get("annualAmount") or 0))
        yearly = [Decimal(str(value or 0))
                  for value in (finance_items.get(key) or {}).get("yearly") or []]
        item_total = sum(yearly[build:], Decimal("0"))
        share = (item_total / operation_total * 100) if operation_total > 0 else None
        content = ("提资原值 {:,.2f} 按「{}」表头填报，判定单位疑误（元填成万元）；"
                   "按元计列为 {:.2f} 万元/年，占运营期营业收入 {}；请提资方核实该行单位与金额").format(
            Decimal(str(suspect.get("originalAmount") or 0)),
            suspect.get("declaredUnit") or "万元", restored,
            "{:.2f}%".format(share) if share is not None else "—")
        rows.append([len(rows) + 1, "{} {}单位口径".format(key, name), content,
                     "表 8-2、8.1.3", "业主单位 / 提资方", "待确认"])

    for item in (
        ("实施主体授权文件", "项目业主单位对实施主体的书面确认与授权", "1.1、6.1", "业主单位"),
        ("政策性补助文件", "补助规模、比例与到位时间的批复文件", "7.2、8.3", "业主单位 / 财政主管部门"),
        ("停车与充电配置比例", "停车位与充电桩的配建比例（数量已按附表1/附表4 计列）", "4.1、5.2", "业主单位 / 设计单位"),
        ("运营收入确认值", "出租率、租金水平及增长率", "8.1、8.3", "业主单位 / 运营主体"),
    ):
        rows.append([len(rows) + 1] + list(item) + ["待确认"])

    workbook, _cover = new_book("待确认清单")
    add_sheet(workbook, "待确认清单",
              ["序号", "事项", "需确认内容", "影响章节/位置", "确认方", "状态"], rows,
              heading=[("待确认清单（内部件）", 16, True),
                       ("用途：登记成品正文不写的待确认事项，随内部件归档，不随建议书交付", 9, False)],
              note="成品正文只写有依据的内容与正面测算口径；确认结论回填至对应章节后再据以修订。",
              force_decimals={0: 0}, portrait=True)
    # 封面单独一页在 A4 上近乎空白，标题块已并入首页表头。
    workbook.remove(_cover)
    return workbook


def _signature_block(cover, start_row=7):
    """封面签章区：只留栏目名，不写占位符文字，空格供签章/手填。"""
    labels = ["编制单位", "编制人", "审核人", "批准人", "编制日期"]
    for offset, label in enumerate(labels):
        row = start_row + offset
        cell = cover.cell(row=row, column=1, value=label)
        cell.font = Font(bold=True, size=10)
        cell.alignment = LEFT
        cover.merge_cells(start_row=row, start_column=2, end_row=row, end_column=6)
        cover.cell(row=row, column=2).border = Border(bottom=THIN)
        cover.row_dimensions[row].height = 24
    cover.print_area = "A1:F{}".format(start_row + len(labels) - 1)


def _signature_sheet(workbook):
    sheet = workbook.create_sheet("签署")
    sheet["A1"] = "签署"
    sheet["A1"].font = TITLE_FONT
    sheet.column_dimensions["A"].width = 16
    sheet.column_dimensions["B"].width = 34
    labels = ["编制单位", "编制人", "审核人", "批准人", "编制日期"]
    for offset, label in enumerate(labels):
        row = 3 + offset
        sheet.cell(row=row, column=1, value=label).font = Font(bold=True, size=10)
        sheet.cell(row=row, column=1).border = BORDER
        sheet.cell(row=row, column=2).border = Border(bottom=THIN)
        sheet.row_dimensions[row].height = 26
    print_setup(sheet, title_rows="1:2", portrait=True, fit_height=1)
    sheet.print_area = "A1:B{}".format(2 + len(labels))
    return sheet


def _d(value):
    try:
        return Decimal(str(value if value not in (None, "") else 0))
    except Exception:  # noqa: BLE001
        return Decimal("0")


def _diff(left, right):
    return (left - right).quantize(Decimal("0.01"))


def _init_book(title):
    workbook = openpyxl.Workbook()
    workbook.properties.created = FIXED_TIME
    workbook.properties.modified = FIXED_TIME
    workbook.properties.creator = "CardoInvest"
    workbook.properties.title = title
    workbook.calculation.fullCalcOnLoad = True
    workbook.remove(workbook.active)
    return workbook


def _sign_block(sheet, start_row, label_col=1):
    """公文签署栏：左侧栏名，右侧留白下划线，手签用。"""
    row = start_row
    for label in ("编制单位", "编制人", "审核人", "批准人", "日期"):
        _put(sheet, row, label_col, label, bold=True, align="l", border=False)
        cell = _put(sheet, row, label_col + 1, None, align="l", border=False)
        cell.border = Border(bottom=Side(style="thin", color="404040"))
        sheet.row_dimensions[row].height = 24
        row += 1
    return row


def _estimate_notes(estimate):
    eng = _num(estimate.get("engineering")) or 0.0
    other = _num(estimate.get("otherFeeTotal")) or 0.0
    cont = _num(estimate.get("contingency")) or 0.0
    interest = _num(estimate.get("constructionInterest")) or 0.0
    total = _num(estimate.get("totalInvestment")) or 0.0
    rate = cont / (eng + other) if (eng + other) else 0.0
    funding = estimate.get("funding") or {}
    subsidy = _num(funding.get("subsidy")) or 0.0
    loan = _num(funding.get("loan")) or 0.0
    gap = round(total - subsidy - loan, 2)
    return [
        ("项目名称", "示例片区更新提升项目"),
        ("编制范围", "老旧街区更新提升，含市政管线综合更新、建筑外立面整治、街区环境整治、道路改造、"
                     "充电桩五项，不实施整体拆除重建、不新增建设用地"),
        ("编制依据", "《成都市城市更新条例》及现行建设工程计价依据；项目资料取本项目提资《财务参数》表；"
                     "文件依据见项目建议书第 1.4 节"),
        ("工程费用", "{} 个单项工程，合计 {:.2f} 万元；按综合造价指标计列，提资未分列建筑工程费、"
                     "安装工程费与设备及工器具购置费".format(len(estimate.get("engineeringRows") or []), eng)),
        ("工程建设其他费", "{} 项，合计 {:.2f} 万元；按提资财务参数表金额计列，计费基数与费率待初步设计"
                           "阶段复核".format(len(estimate.get("otherFees") or []), other)),
        ("基本预备费", "以工程费用与工程建设其他费之和为基数，费率 {:.2f}%，计 {:.2f} 万元".format(rate * 100, cont)),
        ("建设期利息", "资金筹措无借款，建设期利息 {:.2f} 万元".format(interest)),
        ("总投资", "工程费用 + 工程建设其他费 + 基本预备费 + 建设期利息 = {:.2f} 万元".format(total)),
        ("资金筹措", "政策性补助 {:.2f} 万元，专项债 {:.2f} 万元；与总投资差额 {:.2f} 万元，为补助取整形成，"
                     "在 1 万元容差内".format(subsidy, loan, gap)),
        ("计量单位", "万元，金额保留两位小数；占总投资比例按四舍五入保留两位百分数"),
        ("勾稽关系", "工程费用分项合计 = 一；其他费分项合计 = 二；一 + 二 + 三 + 四 = 六（项目总投资）；"
                     "各分项合计与分年投资计划合计一致"),
    ]


def _estimate_note_sheet(workbook, estimate):
    sheet = workbook.create_sheet("编制说明")
    sheet.column_dimensions["A"].width = 20
    sheet.column_dimensions["B"].width = 88
    _merge(sheet, 1, 1, 1, 2, "编制说明", bold=True, size=16, border=False, font=F_SONG)
    sheet.row_dimensions[1].height = 34
    band = PatternFill("solid", fgColor="D9E2F3")
    _put(sheet, 3, 1, "项目", bold=True, fill=band)
    _put(sheet, 3, 2, "内容", bold=True, fill=band)
    row = 4
    for name, text in _estimate_notes(estimate):
        _put(sheet, row, 1, name, bold=True, align="l")
        _put(sheet, row, 2, text, align="l", size=10)
        sheet.row_dimensions[row].height = 34 if len(text) > 44 else 22
        row += 1
    last = _sign_block(sheet, row + 2)
    print_setup(sheet, title_rows=None, portrait=True, fit_height=1)
    sheet.print_area = "A1:B{}".format(last - 1)


def _estimate_main_sheet(workbook, estimate):
    sheet = workbook.create_sheet("附表1工程总投资估算表")
    for key, width in (("A", 6), ("B", 40), ("C", 11), ("D", 11), ("E", 15), ("F", 11),
                       ("G", 12), ("H", 7), ("I", 9), ("J", 13), ("K", 34), ("L", 11)):
        sheet.column_dimensions[key].width = width
    _head_block(sheet, "工程总投资匡算表", "附表1", "单位：万元", 12, title_size=18)
    _grid_header(sheet, 3, 1, "序号", down=2)
    _grid_header(sheet, 3, 2, "工程或费用名称", down=2)
    _grid_header(sheet, 3, 3, "估算价值（万元）", span=5)
    for offset, text in enumerate(("建筑工程费", "安装工程费", "设备及工器具购置费", "其他费", "小\u3000计")):
        _grid_header(sheet, 4, 3 + offset, text)
    _grid_header(sheet, 3, 8, "单位", down=2)
    _grid_header(sheet, 3, 9, "技术经济指标", span=2)
    _grid_header(sheet, 4, 9, "工程量")
    _grid_header(sheet, 4, 10, "单位价值(元/单位)")
    _grid_header(sheet, 3, 11, "备注", down=2)
    _grid_header(sheet, 3, 12, "占总投资比例(%)", down=2)
    for row in (3, 4):
        sheet.row_dimensions[row].height = 26 if row == 3 else 30

    eng_total = _num(estimate.get("engineering")) or 0.0
    other_total = _num(estimate.get("otherFeeTotal")) or 0.0
    cont = _num(estimate.get("contingency")) or 0.0
    interest = _num(estimate.get("constructionInterest")) or 0.0
    total = _num(estimate.get("totalInvestment")) or 0.0
    rate = round(cont / (eng_total + other_total), 4) if (eng_total + other_total) else 0.0
    share = (lambda value: value / total) if total else (lambda value: 0.0)
    eng_rows = estimate.get("engineeringRows") or []
    eng_last = 5 + len(eng_rows)
    other_head = eng_last + 1
    fee_rows = estimate.get("otherFees") or []
    fee_first = other_head + 1
    fee_last = fee_first + len(fee_rows) - 1
    reserve = fee_last + 1
    static = reserve + 1
    finance = static + 1
    total_row = finance + 3

    _put(sheet, 5, 1, "一", bold=True)
    _put(sheet, 5, 2, "工程费", bold=True, align="l")
    _f(sheet, 5, 7, "=SUM(G6:G{})".format(eng_last), eng_total, FMT_AMT2, bold=True)
    _f(sheet, 5, 12, "=G5/$G${}".format(total_row), share(eng_total), FMT_PCT, bold=True)

    row = 6
    for index, item in enumerate(eng_rows):
        amount = _num(item.get("amount")) or 0.0
        _put(sheet, row, 1, "（{}）".format("一二三四五六七八九十"[index]))
        _put(sheet, row, 2, item.get("name"), align="l")
        _put(sheet, row, 7, amount, FMT_AMT2)
        _f(sheet, row, 12, "=G{0}/$G${1}".format(row, total_row), share(amount), FMT_PCT)
        _put(sheet, row, 11, "按综合造价指标计列，未分列建筑安装与设备", align="l", size=9)
        sheet.row_dimensions[row].height = 20
        row += 1

    _put(sheet, other_head, 1, "二", bold=True)
    _put(sheet, other_head, 2, "工程建设其他费", bold=True, align="l")
    _f(sheet, other_head, 6, "=SUM(F{}:F{})".format(fee_first, fee_last), other_total, FMT_AMT2, bold=True)
    _f(sheet, other_head, 7, "=F{}".format(other_head), other_total, FMT_AMT2, bold=True)
    _f(sheet, other_head, 12, "=G{0}/$G${1}".format(other_head, total_row), share(other_total), FMT_PCT, bold=True)
    _put(sheet, other_head, 11, "按提资财务参数表金额计列", align="l", size=9)
    sheet.row_dimensions[other_head].height = 20

    for index, item in enumerate(fee_rows):
        row = fee_first + index
        amount = _num(item.get("amount")) or 0.0
        _put(sheet, row, 1, str(index + 1))
        _put(sheet, row, 2, item.get("name"), align="l")
        _put(sheet, row, 6, amount, FMT_AMT2)
        _f(sheet, row, 7, "=F{}".format(row), amount, FMT_AMT2)
        _f(sheet, row, 12, "=G{0}/$G${1}".format(row, total_row), share(amount), FMT_PCT)
        sheet.row_dimensions[row].height = 18

    _put(sheet, reserve, 1, "三", bold=True)
    _put(sheet, reserve, 2, "基本预备费", bold=True, align="l")
    _f(sheet, reserve, 6, "=ROUND((G5+F{})*{},2)".format(other_head, rate), cont, FMT_AMT2, bold=True)
    _f(sheet, reserve, 7, "=F{}".format(reserve), cont, FMT_AMT2, bold=True)
    _f(sheet, reserve, 12, "=G{0}/$G${1}".format(reserve, total_row), share(cont), FMT_PCT, bold=True)
    _put(sheet, reserve, 11, "（工程费用 + 工程建设其他费）×{:.2f}%".format(rate * 100), align="l", size=9)

    _put(sheet, static, 1, "四", bold=True)
    _put(sheet, static, 2, "静态投资", bold=True, align="l")
    _f(sheet, static, 6, "=F{}+F{}".format(other_head, reserve), other_total + cont, FMT_AMT2, bold=True)
    _f(sheet, static, 7, "=G5+G{}+G{}".format(other_head, reserve), eng_total + other_total + cont,
       FMT_AMT2, bold=True)
    _f(sheet, static, 12, "=G{0}/$G${1}".format(static, total_row), share(eng_total + other_total + cont),
       FMT_PCT, bold=True)

    _put(sheet, finance, 1, "五", bold=True)
    _put(sheet, finance, 2, "建设期融资费", bold=True, align="l")
    _f(sheet, finance, 6, "=SUM(F{}:F{})".format(finance + 1, finance + 2), interest, FMT_AMT2, bold=True)
    _f(sheet, finance, 7, "=F{}".format(finance), interest, FMT_AMT2, bold=True)
    _f(sheet, finance, 12, "=G{0}/$G${1}".format(finance, total_row), share(interest), FMT_PCT, bold=True)
    _put(sheet, finance + 1, 1, "1")
    _put(sheet, finance + 1, 2, "专项债利息", align="l")
    _put(sheet, finance + 1, 6, 0, FMT_AMT2)
    _put(sheet, finance + 2, 1, "2")
    _put(sheet, finance + 2, 2, "债券发行费", align="l")
    _put(sheet, finance + 2, 6, 0, FMT_AMT2)

    _put(sheet, total_row, 1, "六", bold=True)
    _put(sheet, total_row, 2, "项目总投资", bold=True, align="l")
    _f(sheet, total_row, 7, "=G{}+G{}".format(static, finance), total, FMT_AMT2, bold=True)
    _f(sheet, total_row, 12, "=G{0}/$G${0}".format(total_row), 1.0, FMT_PCT, bold=True)
    for row in (reserve, static, finance, total_row):
        sheet.row_dimensions[row].height = 20

    # 辅助汇总区：7.1「总投资构成表」的数据源，位于打印区外
    base = total_row + 2
    _merge(sheet, base, 1, base, 4, "附：总投资构成汇总（不随本表打印）", bold=True, align="l",
           border=False, size=9)
    for offset, text in enumerate(("序号", "工程及费用名称", "合计（万元）", "占总投资比例")):
        _put(sheet, base + 1, offset + 1, text, bold=True, size=9)
    summary = (("一", "工程费", eng_total), ("二", "工程建设其他费", other_total),
               ("三", "基本预备费", cont), ("四", "建设期利息", interest))
    for index, (number, name, amount) in enumerate(summary):
        row = base + 2 + index
        _put(sheet, row, 1, number, size=9)
        _put(sheet, row, 2, name, align="l", size=9)
        _put(sheet, row, 3, amount, FMT_AMT2, size=9)
        _put(sheet, row, 4, share(amount), FMT_PCT, size=9)
    last = base + 2 + len(summary)
    _put(sheet, last, 1, "合计", size=9)
    _put(sheet, last, 2, "项目总投资", align="l", size=9)
    _f(sheet, last, 3, "=SUM(C{}:C{})".format(base + 2, last - 1), total, FMT_AMT2, size=9)
    _put(sheet, last, 4, 1.0, FMT_PCT, size=9)

    print_setup(sheet, title_rows="1:4", portrait=False, fit_height=0)
    sheet.page_setup.paperSize = sheet.PAPERSIZE_A3
    sheet.print_area = "A1:L{}".format(total_row)


def _estimate_fee_sheet(workbook, estimate):
    sheet = workbook.create_sheet("工程建设其他费计算汇总表")
    for key, width in (("A", 5.1), ("B", 29), ("C", 11.7), ("D", 25.4), ("E", 14.1), ("F", 23.1),
                       ("G", 14.1), ("H", 12), ("I", 5.7), ("J", 12), ("K", 45.6), ("L", 10)):
        sheet.column_dimensions[key].width = width
    heads = ("序号", "费用名称", "计费基数", "利用计费公式测算基准价", "专业调整系数",
             "工程复杂程度调整系数", "高程调整系数", "附加调整系数", "下浮", "最终收费", "备注", "取费费率")
    for index, text in enumerate(heads):
        _put(sheet, 1, index + 1, text, bold=True, size=9)
    sheet.row_dimensions[1].height = 42

    fees = {item.get("name"): (_num(item.get("amount")) or 0.0)
            for item in (estimate.get("otherFees") or [])}
    groups = [
        ("建设管理费", ["建设单位管理费", "建设工程监理费"]),
        ("建设项目前期咨询费", ["项目建议书编制费", "可行性研究报告编制费"]),
        ("勘察设计费", ["勘察费", "设计费"]),
        ("招标代理费", []),
        ("施工图审查费", []),
        ("工程造价咨询费", ["工程量清单及控制价编制费", "竣工结算审核费", "施工阶段全过程造价控制"]),
        ("水土保持方案编制费", []),
        ("水土保持补偿费", []),
        ("安全预评费", []),
        ("安全验收评价费", []),
        ("工程检测费", []),
        ("社会稳定风险评估", []),
        ("管探费", []),
        ("道路检测费", []),
        ("临时交通组织费", []),
        ("施工占道费", []),
    ]
    used = set()
    for group, children in groups:
        # 无子项的组本身就是一条费用行；它的名字必须记为已用，
        # 否则下面的补漏循环会把它再当成「未登记费用」追加一遍（实测导致 12 行重复）。
        if children:
            used.update(children)
        else:
            used.add(group)
    for name in fees:
        if name not in used:
            groups.append((name, []))

    row = 2
    top_rows = []
    serial = 0
    for group, children in groups:
        present = [name for name in children if name in fees]
        if not children:
            serial += 1
            _put(sheet, row, 1, str(serial))
            _put(sheet, row, 2, group, align="l")
            _put(sheet, row, 10, fees.get(group, 0.0), FMT_AMT2)
            _put(sheet, row, 11, "按提资财务参数表金额计列", align="l", size=9)
            top_rows.append(row)
            sheet.row_dimensions[row].height = 20
            row += 1
            continue
        if not present:
            continue
        serial += 1
        _put(sheet, row, 1, str(serial))
        _put(sheet, row, 2, group, align="l", bold=True)
        head_row = row
        first_child = row + 1
        _f(sheet, row, 10, "=SUM(J{}:J{})".format(first_child, first_child + len(present) - 1),
           sum(fees[name] for name in present), FMT_AMT2, bold=True)
        top_rows.append(head_row)
        row += 1
        for index, name in enumerate(present):
            _put(sheet, row, 1, "{}.{}".format(serial, index + 1))
            _put(sheet, row, 2, name, align="l")
            _put(sheet, row, 10, fees[name], FMT_AMT2)
            _put(sheet, row, 11, "按提资财务参数表金额计列", align="l", size=9)
            sheet.row_dimensions[row].height = 20
            row += 1

    _put(sheet, row, 1, "合计", bold=True)
    _put(sheet, row, 2, "工程建设其他费合计", bold=True, align="l")
    refs = "+".join("J{}".format(item) for item in top_rows)
    _f(sheet, row, 10, "={}".format(refs or "0"), _num(estimate.get("otherFeeTotal")) or 0.0,
       FMT_AMT2, bold=True)
    total_row = row

    note_row = total_row + 2
    _merge(sheet, note_row, 1, note_row, 12,
           "注：单位万元。本项目其他费按提资财务参数表金额计列；计费基数（C 列）与取费费率（L 列）"
           "提资未提供，待初步设计阶段按现行计价依据复核后填列。",
           align="l", border=False, size=9)
    sheet.row_dimensions[note_row].height = 30
    print_setup(sheet, title_rows="1:1", portrait=True, fit_height=1)
    sheet.print_area = "A1:L{}".format(note_row)


def _estimate_schedule_sheet(workbook, estimate):
    sheet = workbook.create_sheet("附表2-工程进度计划表")
    sheet.column_dimensions["A"].width = 9
    sheet.column_dimensions["B"].width = 9
    months = 12 * int(((estimate.get("tables") or {}).get("yearPlan") or {}).get("constructionYears") or 2)
    for index in range(months):
        sheet.column_dimensions[get_column_letter(3 + index)].width = 3.4
    last_col = 2 + months
    _merge(sheet, 1, 1, 1, last_col, "工程进度计划表", bold=True, size=16, border=False, font=F_SONG)
    sheet.row_dimensions[1].height = 30
    _put(sheet, 2, 1, "附表2", bold=True, size=11, align="l", border=False, font=F_SONG)
    _merge(sheet, 3, 1, 4, 2, "年份\n项目", bold=True)
    years = 2
    per_year = 12
    for index in range(years):
        start = 3 + index * per_year
        sheet.merge_cells(start_row=3, start_column=start, end_row=3, end_column=start + per_year - 1)
        _put(sheet, 3, start, str(FIRST_YEAR + index), bold=True)
        for month in range(per_year):
            _put(sheet, 4, start + month, month + 1, bold=True, size=9)
    sheet.row_dimensions[3].height = 20
    sheet.row_dimensions[4].height = 18

    tasks = [
        ("完成项目建议书、可行性研究报告、勘察设计、招投标等前期工作", 5, 0, 6),
        ("老旧街区改造、配套设施及其他工程施工", 7, 6, 15),
        ("竣工验收并交付使用", 9, 15, 18),
    ]
    bar = PatternFill("solid", fgColor="8EA9DB")
    for text, row, start, length in tasks:
        _merge(sheet, row, 1, row + 1, 2, text, align="l", size=9)
        for column in range(3 + start, 3 + start + length):
            if column > last_col:
                break
            for line in (row, row + 1):
                cell = _put(sheet, line, column, None, size=9)
                cell.fill = bar
        sheet.row_dimensions[row].height = 18
        sheet.row_dimensions[row + 1].height = 18
    print_setup(sheet, title_rows="1:4", portrait=False, fit_height=0)
    sheet.print_area = "A1:{}{}".format(get_column_letter(last_col), 10)
    sheet.oddFooter.center.text = "第 &P 页，共 &N 页"
    sheet.oddFooter.center.size = 9


def export_estimate_deliverable(estimate):
    """匡算表（成品版）：按提资同族模板出编制说明、附表1、其他费汇总、附表2。"""
    reset_cache()
    workbook = _init_book("匡算表")
    _estimate_note_sheet(workbook, estimate)
    _estimate_main_sheet(workbook, estimate)
    _estimate_fee_sheet(workbook, estimate)
    _estimate_schedule_sheet(workbook, estimate)
    return workbook


def export_finance_deliverable(finance, payload, estimate):
    """财务评价表（成品版）：按提资同族模板出 13 张表（含说明与指标表）。"""
    reset_cache()
    workbook = _init_book("财务评价表")
    period = finance.get("period") or {}
    _finance_note_sheet(workbook, finance, payload, estimate)
    _finance_capital_sheet(workbook, finance, estimate, period)
    _finance_revenue_sheet(workbook, finance, period)
    _finance_cost_sheet(workbook, finance, period)
    _finance_profit_sheet(workbook, finance, period)
    _finance_depreciation_sheet(workbook, finance, estimate, period)
    _finance_debt_sheet(workbook, finance, period)
    _finance_cashflow_sheet(workbook, finance, estimate, period)
    _finance_equity_sheet(workbook, finance, period)
    _finance_plan_sheet(workbook, finance, estimate, period)
    _finance_bep_sheet(workbook, finance, period)
    _finance_sensitivity_sheet(workbook, finance, period)
    _finance_indicator_sheet(workbook, finance, estimate, period)
    return workbook


def _finance_year_header(sheet, period, first_col=8, head_row=3, year_row=4, label="运营期"):
    """年列：H 起（或指定列）22 个年度，建设期用自然年、运营期首年标「下半年」。"""
    labels = _year_labels(period)
    sheet.merge_cells(start_row=head_row, start_column=first_col, end_row=head_row,
                      end_column=first_col + len(labels) - 1)
    _put(sheet, head_row, first_col, label, bold=True)
    for index, text in enumerate(labels):
        _put(sheet, year_row, first_col + index, text, bold=True, size=9)
    sheet.row_dimensions[head_row].height = 22
    sheet.row_dimensions[year_row].height = 24
    return labels


def _row_series(sheet, row, values, first_col=8, fmt=FMT_AMT1, formula=None, total_col=None):
    """写一行年列：formula(ref) 给公式（同时登记缓存），不给则写硬值。"""
    values = _num_series(values)
    for index, value in enumerate(values):
        col = first_col + index
        if formula is None:
            _put(sheet, row, col, value, fmt)
        else:
            _f(sheet, row, col, formula(get_column_letter(col)), value, fmt)
    if total_col:
        last = get_column_letter(first_col + len(values) - 1)
        head = get_column_letter(first_col)
        _f(sheet, row, total_col, "=SUM({0}{1}:{2}{1})".format(head, row, last),
           round(sum(values), 2), fmt, bold=True)
    return round(sum(values), 2)


def _side_note(sheet, first_col, count, note, row):
    sheet.merge_cells(start_row=row, start_column=first_col, end_row=row,
                      end_column=first_col + count - 1)
    _put(sheet, row, first_col, note, align="l", border=False, size=9)


def _finance_note_sheet(workbook, finance, payload, estimate):
    sheet = workbook.create_sheet("测算说明")
    sheet.column_dimensions["A"].width = 20
    sheet.column_dimensions["B"].width = 88
    _merge(sheet, 1, 1, 1, 2, "测算说明", bold=True, size=16, border=False, font=F_SONG)
    sheet.row_dimensions[1].height = 34
    period = finance.get("period") or {}
    indicators = finance.get("indicators") or {}
    equity = _num(((payload or {}).get("finance") or {}).get("equity")) or 0.0
    try:
        rate_text = "{:.2f}%".format(float(period.get("discountRate")) * 100)
    except (TypeError, ValueError):
        rate_text = "—"
    items = [
        ("项目名称", "示例片区更新提升项目"),
        ("计算期", "建设期 {} 年（{}—{}）+ 运营期 {} 年（{} 年下半年起算）= {} 年".format(
            period.get("constructionYears") or "—", FIRST_YEAR, FIRST_YEAR + 1,
            period.get("operationYears") or "—", FIRST_YEAR + 2, period.get("totalYears") or "—")),
        ("折现率", "{}（财务基准收益率，用于财务净现值）".format(rate_text)),
        ("评价口径", "所得税后；全部投资与资本金两套现金流并列，资本金为 0 时不适用资本金口径"),
        ("营业收入", "停车位、流动餐车租赁、外摆空间租赁、广告四项，增值税率 6%；按提资财务参数表"
                     "年均收入计列，量价参数待运营方案深化确定"),
        ("成本费用", "外购燃料及动力费、工资及福利费、管理费用、日常修理费、其他费用、大修费 6 项；"
                     "运营期首年按半年计"),
        ("折旧", "房屋建筑物折旧年限 30 年、残值率 5%、首年半年；无形资产摊销表不出（本项目无销售部分）"),
        ("税金及附加", "增值税 6%；城市维护建设税 7%、教育费附加 5%、印花税租赁收入 0.1%、"
                       "土地使用税 8 元/平方米；房产税未计征"),
        ("所得税", "企业所得税税率 25%"),
        ("资金筹措", "政策性补助 {} 万元，无政府专项债、无银行借款；附表7 按无借款情形列示".format(
            (estimate.get("funding") or {}).get("subsidy"))),
        ("资本金", "资本金为 {} 万元，资本金 FIRR、资本金回收期、资本金净利润率不适用".format(equity)
         if equity <= 0 else "资本金为 {} 万元".format(equity)),
        ("主要指标", "全部投资 FIRR {}，FNPV {} 万元，静态投资回收期 {} 年，总投资收益率 {}".format(
            indicators.get("firrAll"), indicators.get("fnpvAll"), indicators.get("ptAll"),
            indicators.get("roi"))),
        ("敏感性分析", "单因素法，投资、收入、经营成本各按 ±10% 两档测算，仅调整该因素分量、其余冻结"),
        ("数据来源", "过程文件/03-财务评价/财务测算.json；投资口径与匡算表附表1 同源"),
    ]
    band = PatternFill("solid", fgColor="D9E2F3")
    _put(sheet, 3, 1, "项目", bold=True, fill=band)
    _put(sheet, 3, 2, "说明", bold=True, fill=band)
    row = 4
    for name, text in items:
        _put(sheet, row, 1, name, bold=True, align="l")
        _put(sheet, row, 2, text, align="l", size=10)
        sheet.row_dimensions[row].height = 34 if len(text) > 44 else 22
        row += 1
    last = _sign_block(sheet, row + 2)
    print_setup(sheet, title_rows=None, portrait=True, fit_height=1)
    sheet.print_area = "A1:B{}".format(last - 1)


def _finance_capital_sheet(workbook, finance, estimate, period):
    sheet = workbook.create_sheet("附表3-项目总投资使用计划与资金筹措表")
    sheet.column_dimensions["A"].width = 6
    sheet.column_dimensions["B"].width = 32
    for key in ("C", "D", "E"):
        sheet.column_dimensions[key].width = 8
    for key in ("F", "G"):
        sheet.column_dimensions[key].width = 12
    for index in range(22):
        sheet.column_dimensions[get_column_letter(8 + index)].width = 9
    _head_block(sheet, "项目总投资使用计划与资金筹措表", "附表3", "单位：万元", 29, title_size=14)
    _grid_header(sheet, 3, 1, "序号", down=2)
    _grid_header(sheet, 3, 2, "项目", down=2)
    _grid_header(sheet, 3, 3, "备注", span=3)
    _grid_header(sheet, 3, 6, "拟定总投资/校核", down=2)
    _grid_header(sheet, 3, 7, "合计", down=2)
    _finance_year_header(sheet, period, 8, label="投资使用计划")

    rows = ((estimate.get("tables") or {}).get("yearPlan") or {}).get("rows") or []
    build = int(period.get("constructionYears") or 2)
    total_years = int(period.get("totalYears") or 22)
    invest = [0.0] * total_years
    for index, item in enumerate(rows[:build]):
        invest[index] = _num(item.get("constructionInvestment")) or 0.0
    subsidy_total = _num((estimate.get("funding") or {}).get("subsidy")) or 0.0
    subsidy = list(invest)
    if build:
        subsidy[build - 1] = round(subsidy_total - sum(subsidy[:build - 1]), 2)
    total = _num(estimate.get("totalInvestment")) or 0.0

    def line(row, serial, name, values, formula=None, fmt=FMT_AMT1, note=None, bold=False):
        _put(sheet, row, 1, serial, bold=bold)
        _put(sheet, row, 2, name, align="l", bold=bold)
        if note:
            _merge(sheet, row, 3, row, 5, note, align="l", size=9)
        _row_series(sheet, row, values, 8, fmt, formula, total_col=7)
        sheet.row_dimensions[row].height = 20

    _put(sheet, 5, 1, "1", bold=True)
    _put(sheet, 5, 2, "投资总额", bold=True, align="l")
    _put(sheet, 5, 6, total, FMT_AMT2, bold=True)
    _row_series(sheet, 5, invest, 8, FMT_AMT1,
                lambda ref: "={0}6+{0}7".format(ref), total_col=7)
    line(6, "1.1", "固定资产投资", invest)
    line(7, "1.2", "建设期融资费", [0.0] * total_years)
    _put(sheet, 8, 1, "2", bold=True)
    _put(sheet, 8, 2, "资金筹措", bold=True, align="l")
    _row_series(sheet, 8, [subsidy[i] for i in range(total_years)], 8, FMT_AMT1,
                lambda ref: "={0}9+{0}13".format(ref), total_col=7)
    _put(sheet, 9, 1, "2.1")
    _put(sheet, 9, 2, "财政资金", align="l")
    _f(sheet, 9, 6, "=G9", subsidy_total, FMT_AMT2)
    _f(sheet, 9, 4, "=G9/G8", (subsidy_total / total if total else 0.0), FMT_PCT)
    _row_series(sheet, 9, subsidy, 8, FMT_AMT1, total_col=7)
    line(10, "2.1.1", "用于建设投资", subsidy)
    line(11, "2.1.2", "用于支付建设期利息", [0.0] * total_years)
    line(12, "2.1.3", "用于支付债券发行费", [0.0] * total_years)
    line(13, "2.2", "申请政府专项债", [0.0] * total_years)
    line(14, "2.2.1", "用于建设投资", [0.0] * total_years)

    _put(sheet, 16, 2, "专项债利率", align="l")
    _put(sheet, 16, 6, 0.027, FMT_PCT)
    _side_note(sheet, 8, 22, "注：本项目无专项债、无银行借款，建设期融资费为 0；专项债利率仅列示参数，未发生计息。", 17)
    print_setup(sheet, title_rows="1:4", portrait=False, fit_height=0)
    sheet.page_setup.paperSize = sheet.PAPERSIZE_A4
    sheet.print_area = "A1:AC17"


def _finance_revenue_sheet(workbook, finance, period):
    sheet = workbook.create_sheet("附表4 -项目收入、增值税及附加表")
    sheet.column_dimensions["A"].width = 6
    sheet.column_dimensions["B"].width = 34
    for key in ("C", "D", "E"):
        sheet.column_dimensions[key].width = 10
    for key in ("F", "G"):
        sheet.column_dimensions[key].width = 13
    for index in range(22):
        sheet.column_dimensions[get_column_letter(8 + index)].width = 9
    _head_block(sheet, "项目收入、增值税、税金及附加表", "附表4", "单位：万元", 29, title_size=14)
    _grid_header(sheet, 3, 1, "序号", down=2)
    _grid_header(sheet, 3, 2, "项目", down=2)
    _grid_header(sheet, 3, 3, "备注", span=3, down=2)
    _grid_header(sheet, 3, 6, "拟定总投资/校核", down=2)
    _grid_header(sheet, 3, 7, "合计", down=2)
    _finance_year_header(sheet, period, 8, label="运营期")

    revenue = finance.get("revenue") or {}
    items = revenue.get("items") or []
    cost = finance.get("cost") or {}
    surcharges = cost.get("surcharges") or {}
    output_vat = _num_series(revenue.get("outputVat"))
    input_vat = _num_series(cost.get("inputVat"))
    years = int(period.get("totalYears") or 22)

    def fill(row, serial, name, values, formula=None, fmt=FMT_AMT1, note=None, bold=False):
        _put(sheet, row, 1, serial, bold=bold)
        _put(sheet, row, 2, name, align="l", bold=bold)
        if note is not None:
            _put(sheet, row, 3, note, align="l", size=9)
        _row_series(sheet, row, values, 8, fmt, formula, total_col=7)
        sheet.row_dimensions[row].height = 20

    totals = []
    for index in range(len(items)):
        totals.append(_num_series(items[index].get("yearly")))
    while len(totals) < 4:
        totals.append([0.0] * years)

    _put(sheet, 5, 1, "一", bold=True)
    _put(sheet, 5, 2, "收入合计", bold=True, align="l")
    _row_series(sheet, 5, [sum(totals[i][j] for i in range(4)) for j in range(years)], 8, FMT_AMT1,
                lambda ref: "={0}6+{0}7+{0}8+{0}9".format(ref), total_col=7)
    sheet.row_dimensions[5].height = 20
    for index in range(4):
        name = items[index].get("name") if index < len(items) else "—"
        fill(6 + index, str(index + 1), name, totals[index],
             note="增值税率 6%；按提资财务参数表年均收入计列，量价参数待运营方案深化确定")

    _put(sheet, 10, 1, "5")
    _put(sheet, 10, 2, "充电桩", align="l")
    _put(sheet, 10, 3, "提资未列此项收入；充电桩数量待确认，运营方案深化阶段确定", align="l", size=9)
    for index in range(years):
        _put(sheet, 10, 8 + index, None, FMT_AMT1)
    sheet.row_dimensions[10].height = 26

    _put(sheet, 11, 1, "二", bold=True)
    _put(sheet, 11, 2, "增值税、税金与附加", bold=True, align="l")
    _put(sheet, 11, 3, "增值税 + 税金及附加", align="l", size=9)
    _row_series(sheet, 11,
                [_num_series(cost.get("vatPayable"))[j] + _num_series(cost.get("surchargeTotal"))[j]
                 for j in range(years)], 8, FMT_AMT1,
                lambda ref: "={0}12+{0}15".format(ref), total_col=7)
    sheet.row_dimensions[11].height = 20
    fill(12, "1", "增值税", _num_series(cost.get("vatPayable")),
         formula=lambda ref: "=MAX(0,{0}13-{0}14)".format(ref))
    fill(13, "1.1", "销项税额", output_vat)
    fill(14, "1.2", "进项税额", input_vat,
         note="建设期按建安工程费 9%、其他费 6% 计，运营期按成本项税率计")
    fill(15, "2", "税金及附加", _num_series(cost.get("surchargeTotal")),
         formula=lambda ref: "=SUM({0}16:{0}20)".format(ref))
    fill(16, "2.1", "城市维护建设税", _num_series(surcharges.get("urban")), note="应纳增值税的 7%")
    fill(17, "2.2", "教育费附加", _num_series(surcharges.get("education")), note="应纳增值税的 5%")
    fill(18, "2.3", "印花税", _num_series(surcharges.get("stamp")), note="租赁收入的 0.1%")
    fill(19, "2.4", "土地使用税", _num_series(surcharges.get("landUse")), note="8 元/平方米")
    fill(20, "2.5", "房产税", _num_series(surcharges.get("property")),
         note="从租计征；本项目未计征")

    _side_note(sheet, 8, 22, "注：本表为附表5、附表6、附表8、附表9 的收入与税金入口；金额单位万元。", 22)
    print_setup(sheet, title_rows="1:4", portrait=False, fit_height=0)
    sheet.page_setup.paperSize = sheet.PAPERSIZE_A3
    sheet.print_area = "A1:AC22"


def _finance_cost_sheet(workbook, finance, period):
    sheet = workbook.create_sheet("附表5-总成本费用表")
    sheet.column_dimensions["A"].width = 6
    sheet.column_dimensions["B"].width = 36
    sheet.column_dimensions["C"].width = 13
    for index in range(22):
        sheet.column_dimensions[get_column_letter(4 + index)].width = 9
    _head_block(sheet, "项目总成本费用表", "附表5", "单位：万元", 25, title_size=14)
    _grid_header(sheet, 3, 1, "序号", down=2)
    _grid_header(sheet, 3, 2, "项目", down=2)
    _grid_header(sheet, 3, 3, "合计", down=2)
    _finance_year_header(sheet, period, 4, label="运营期")

    cost = finance.get("cost") or {}
    items = cost.get("items") or {}
    years = int(period.get("totalYears") or 22)
    series = {key: _num_series(items.get(key) if items.get(key) is not None else cost.get(key))
              for key in ("utilities", "labor", "management", "repair", "other", "overhaul",
                          "depreciation", "interest")}
    operating = _num_series(cost.get("operatingCost"))
    total_cost = _num_series(cost.get("totalCost"))

    labels = (
        ("utilities", "1", "外购燃料及动力费（运营收入的 3%）"),
        ("labor", "2", "员工工资及福利费"),
        ("management", "3", "管理费用"),
        ("repair", "4", "日常修理费"),
        ("other", "5", "其他费用"),
        ("overhaul", "6", "大修费"),
    )
    for row, (key, serial, name) in enumerate(labels, start=5):
        _put(sheet, row, 1, serial)
        _put(sheet, row, 2, name, align="l")
        _row_series(sheet, row, series[key], 4, FMT_AMT1, total_col=3)
        sheet.row_dimensions[row].height = 20

    _put(sheet, 11, 1, "7", bold=True)
    _put(sheet, 11, 2, "经营成本", bold=True, align="l")
    _row_series(sheet, 11, operating, 4, FMT_AMT1, lambda ref: "=SUM({}5:{}10)".format(ref, ref), total_col=3)
    _put(sheet, 12, 1, "8")
    _put(sheet, 12, 2, "折旧摊销", align="l")
    _row_series(sheet, 12, series["depreciation"], 4, FMT_AMT1, total_col=3)
    _put(sheet, 13, 1, "9")
    _put(sheet, 13, 2, "财务费用（利息支出）", align="l")
    _row_series(sheet, 13, series["interest"], 4, FMT_AMT1, total_col=3)
    _put(sheet, 14, 1, "10", bold=True)
    _put(sheet, 14, 2, "总成本费用", bold=True, align="l")
    _row_series(sheet, 14, total_cost, 4, FMT_AMT1,
                lambda ref: "={}11+{}12+{}13".format(ref, ref, ref), total_col=3)
    fixed = [series["labor"][j] + series["management"][j] + series["repair"][j] + series["other"][j]
             + series["depreciation"][j] + series["interest"][j] for j in range(years)]
    _put(sheet, 15, 1, "10.1")
    _put(sheet, 15, 2, "固定成本", align="l")
    _row_series(sheet, 15, fixed, 4, FMT_AMT1,
                lambda ref: "={0}6+{0}7+{0}8+{0}9+{0}12+{0}13".format(ref), total_col=3)
    _put(sheet, 16, 1, "10.2")
    _put(sheet, 16, 2, "可变成本", align="l")
    _row_series(sheet, 16, series["utilities"], 4, FMT_AMT1, lambda ref: "={}5".format(ref), total_col=3)
    for row in range(11, 17):
        sheet.row_dimensions[row].height = 20

    _side_note(sheet, 4, 22, "注：固定成本含工资及福利费、管理费用、日常修理费、其他费用、折旧摊销与利息支出；可变成本为外购燃料及动力费。", 18)
    print_setup(sheet, title_rows="1:4", portrait=False, fit_height=0)
    sheet.page_setup.paperSize = sheet.PAPERSIZE_A3
    sheet.print_area = "A1:Y18"


def _finance_profit_sheet(workbook, finance, period):
    sheet = workbook.create_sheet("附表6-项目利润与利润分配表")
    sheet.column_dimensions["A"].width = 6
    sheet.column_dimensions["B"].width = 32
    for key in ("C", "D", "E"):
        sheet.column_dimensions[key].width = 12
    for key in ("F", "G"):
        sheet.column_dimensions[key].width = 13
    for index in range(22):
        sheet.column_dimensions[get_column_letter(8 + index)].width = 9
    _head_block(sheet, "项目利润与利润分配表", "附表6", "单位：万元", 29, title_size=14)
    _grid_header(sheet, 3, 1, "序号", down=2)
    _grid_header(sheet, 3, 2, "项目", down=2)
    _grid_header(sheet, 3, 3, "备注", span=3, down=2)
    _grid_header(sheet, 3, 6, "拟定总投资/校核", down=2)
    _grid_header(sheet, 3, 7, "合计", down=2)
    _finance_year_header(sheet, period, 8, label="运营期")

    cost = finance.get("cost") or {}
    revenue = _num_series((finance.get("revenue") or {}).get("total"))
    vat = _num_series(cost.get("vatPayable"))
    surcharge = _num_series(cost.get("surchargeTotal"))
    total_cost = _num_series(cost.get("totalCost"))
    income_tax = _num_series(cost.get("incomeTax"))
    ebit = _num_series(cost.get("ebit"))
    ebitda = _num_series(cost.get("ebitda"))
    years = int(period.get("totalYears") or 22)
    profit = [round(revenue[j] - vat[j] - surcharge[j] - total_cost[j], 2) for j in range(years)]
    net = [round(profit[j] - income_tax[j], 2) for j in range(years)]
    retained = []
    carry = 0.0
    for j in range(years):
        available = round(net[j] + carry, 2)
        carry = available
        retained.append(round(available, 2))

    def line(row, serial, name, values, formula=None, note=None, bold=False):
        _put(sheet, row, 1, serial, bold=bold)
        _put(sheet, row, 2, name, align="l", bold=bold)
        if note:
            _put(sheet, row, 3, note, align="l", size=9)
        _row_series(sheet, row, values, 8, FMT_AMT1, formula, total_col=7)
        sheet.row_dimensions[row].height = 20

    line(5, "1", "营业收入", revenue, None, "含税收入，取自附表4", True)
    line(6, "2", "增值税、税金及附加", [round(vat[j] + surcharge[j], 2) for j in range(years)],
         lambda ref: "='附表4 -项目收入、增值税及附加表'!{0}12+'附表4 -项目收入、增值税及附加表'!{0}15"
         .format(ref))
    line(7, "3", "总成本费用", total_cost, lambda ref: "='附表5-总成本费用表'!{0}14".format(ref))
    line(8, "4", "利润总额", profit, lambda ref: "={0}5-{0}6-{0}7".format(ref), None, True)
    line(9, "5", "弥补以前年度亏损", [0.0] * years)
    line(10, "6", "应纳税所得额", profit, lambda ref: "={0}8-{0}9".format(ref))
    line(11, "7", "所得税", income_tax, lambda ref: "=ROUND({}10*25%,2)".format(ref))
    line(12, "8", "净利润", net, lambda ref: "={0}8-{0}11".format(ref), None, True)

    for row, serial, name, values, note in (
            (13, "9", "期初未分配利润", [0.0] + retained[:-1], None),
            (15, "11", "提取法定盈余公积金", [0.0] * years, "按公司章程提取，本表不预设"),
            (17, "13", "提取任意盈余公积金", [0.0] * years, "按公司章程提取，本表不预设")):
        _put(sheet, row, 1, serial)
        _put(sheet, row, 2, name, align="l")
        if note:
            _put(sheet, row, 3, note, align="l", size=9)
        _row_series(sheet, row, values, 8, FMT_AMT1, total_col=7)
        sheet.row_dimensions[row].height = 20
    for index in range(years):
        col = 8 + index
        if index == 0:
            continue
        _put(sheet, 13, col, None)
        _f(sheet, 13, col, "={}18".format(get_column_letter(col - 1)), retained[index - 1], FMT_AMT1)
    line(14, "10", "可供分配的利润", retained, lambda ref: "={0}12+{0}13".format(ref))
    line(16, "12", "可供投资者分配的利润", retained, lambda ref: "={0}14-{0}15".format(ref))
    line(18, "14", "未分配利润", retained, lambda ref: "={0}16-{0}17".format(ref))
    line(19, "15", "息税前利润（EBIT）", ebit,
         lambda ref: "={0}8+'附表5-总成本费用表'!{0}13".format(ref), None, True)
    line(20, "16", "息税折旧摊销前利润（EBITDA）", ebitda,
         lambda ref: "={0}19+'附表5-总成本费用表'!{0}12".format(ref))

    # 表样把两个收益率放在备注侧栏（C/D），便于评审一眼看到
    _put(sheet, 12, 3, "净资产收益率（ROE）", align="l", size=9)
    _put(sheet, 12, 4, "不适用（资本金为 0）", align="l", size=9)
    _put(sheet, 19, 3, "总投资收益率（ROI）", align="l", size=9)
    _f(sheet, 19, 4, "=G19/'附表3-项目总投资使用计划与资金筹措表'!G5/20",
       (_num(finance.get("indicators", {}).get("roi")) or 0.0), FMT_PCT)

    _side_note(sheet, 8, 22, "注：利润总额为负年度不缴纳所得税；盈余公积按公司章程提取，本表不预设比例。", 22)
    print_setup(sheet, title_rows="1:4", portrait=False, fit_height=0)
    sheet.page_setup.paperSize = sheet.PAPERSIZE_A3
    sheet.print_area = "A1:AC22"


def _finance_depreciation_sheet(workbook, finance, estimate, period):
    sheet = workbook.create_sheet("附表6-固定资产折旧费估算表")
    sheet.column_dimensions["A"].width = 6
    sheet.column_dimensions["B"].width = 34
    for key in ("C", "D"):
        sheet.column_dimensions[key].width = 13
    for index in range(22):
        sheet.column_dimensions[get_column_letter(4 + index)].width = 9
    _head_block(sheet, "固定资产折旧费估算表", "附表6", "单位：万元", 25, title_size=14)
    _grid_header(sheet, 3, 1, "序号", down=2)
    _grid_header(sheet, 3, 2, "项目", down=2)
    _grid_header(sheet, 3, 3, "合计", down=2)
    _finance_year_header(sheet, period, 4, label="运营期")

    depreciation = _num_series(((finance.get("cost") or {}).get("items") or {}).get("depreciation"))
    years = int(period.get("totalYears") or 22)
    annual = max(depreciation) if depreciation else 0.0
    original = _num(estimate.get("totalInvestment")) or round(annual * 19.5 / 0.95, 2)
    cumulative = []
    running = 0.0
    for value in depreciation:
        running = round(running + value, 2)
        cumulative.append(running)

    def scalar(row, serial, name, value, fmt=FMT_AMT2):
        _put(sheet, row, 1, serial)
        _put(sheet, row, 2, name, align="l")
        _put(sheet, row, 3, value, fmt, bold=True)
        sheet.row_dimensions[row].height = 20

    scalar(5, "1", "折旧年限（年）", 20, FMT_INT)
    scalar(6, "2", "残值率", 0.05, FMT_PCT)
    scalar(7, "3", "年折旧额", annual)
    scalar(8, "4", "固定资产原值", original)
    _put(sheet, 9, 1, "5")
    _put(sheet, 9, 2, "当期折旧额", align="l")
    _row_series(sheet, 9, depreciation, 4, FMT_AMT1, total_col=3)
    _put(sheet, 10, 1, "6")
    _put(sheet, 10, 2, "累计折旧", align="l")
    _row_series(sheet, 10, cumulative, 4, FMT_AMT1, total_col=None)
    _put(sheet, 11, 1, "7")
    _put(sheet, 11, 2, "固定资产净值", align="l")
    _row_series(sheet, 11, [round(original - value, 2) for value in cumulative], 4, FMT_AMT1)
    for row in (9, 10, 11):
        sheet.row_dimensions[row].height = 20
    _side_note(sheet, 4, 22, "注：年限平均法，折旧年限 20 年、残值率 5%，运营期首年按半年计提；原值取项目总投资。", 13)
    print_setup(sheet, title_rows="1:4", portrait=False, fit_height=0)
    sheet.page_setup.paperSize = sheet.PAPERSIZE_A3
    sheet.print_area = "A1:Y13"


def _finance_debt_sheet(workbook, finance, period):
    sheet = workbook.create_sheet("附表7-借款还本付息计划表")
    sheet.column_dimensions["A"].width = 6
    sheet.column_dimensions["B"].width = 36
    sheet.column_dimensions["C"].width = 14
    for index in range(22):
        sheet.column_dimensions[get_column_letter(4 + index)].width = 9
    _head_block(sheet, "借款还本付息计划表", "附表7", "单位：万元", 25, title_size=14)
    _grid_header(sheet, 3, 1, "序号", down=2)
    _grid_header(sheet, 3, 2, "项目", down=2)
    _grid_header(sheet, 3, 3, "合计", down=2)
    _finance_year_header(sheet, period, 4, label="运营期")

    years = int(period.get("totalYears") or 22)
    zeros = [0.0] * years
    rows = (
        ("1", "专项债", None),
        ("1.1", "期初借款余额", zeros),
        ("1.2", "当期新增借款", zeros),
        ("1.3", "当期应计利息", zeros),
        ("1.4", "当期还本付息", zeros),
        ("", "其中：还本", zeros),
        ("", "其中：付息", zeros),
        ("1.5", "期末借款余额", zeros),
        ("2", "可用于还本付息的资金", zeros),
        ("2.1", "息税折旧摊销前利润（EBITDA）", zeros),
        ("2.1.1", "息税前利润（EBIT）", zeros),
        ("2.1.2", "折旧摊销", zeros),
        ("2.2", "其他可用于还本付息资金", zeros),
        ("2.3", "当期剩余资金", zeros),
    )
    row = 5
    for serial, name, values in rows:
        _put(sheet, row, 1, serial or None, bold=not serial)
        _put(sheet, row, 2, name, align="l", bold=not serial)
        if values is None:
            _put(sheet, row, 3, 0.0, FMT_AMT2, bold=True)
        else:
            _row_series(sheet, row, values, 4, FMT_AMT1, total_col=3)
        sheet.row_dimensions[row].height = 20
        row += 1
    _put(sheet, 19, 1, "3", bold=True)
    _put(sheet, 19, 2, "计算指标", bold=True, align="l")
    _put(sheet, 20, 2, "偿债备付率", align="l")
    _put(sheet, 20, 3, "不适用（无借款）", align="l")
    _put(sheet, 21, 2, "借款偿还期", align="l")
    _put(sheet, 21, 3, "不适用（无借款）", align="l")
    for row in (19, 20, 21):
        sheet.row_dimensions[row].height = 20
    _side_note(sheet, 4, 22, "注：本项目资金筹措为政策性补助，无政府专项债、无银行借款，本表按表式列示并标注不适用。", 23)
    print_setup(sheet, title_rows="1:4", portrait=False, fit_height=0)
    sheet.page_setup.paperSize = sheet.PAPERSIZE_A3
    sheet.print_area = "A1:Y23"


def _cf_values(block, key, years):
    items = (block or {}).get(key) or []
    values = [(_num(item.get("value")) if isinstance(item, dict) else _num(item)) or 0.0 for item in items]
    values = values + [0.0] * max(0, years - len(values))
    return values[:years]


def _finance_cashflow_sheet(workbook, finance, estimate, period):
    sheet = workbook.create_sheet("附表8-项目投资现金流量表")
    sheet.column_dimensions["A"].width = 6
    sheet.column_dimensions["B"].width = 34
    sheet.column_dimensions["C"].width = 13
    for index in range(22):
        sheet.column_dimensions[get_column_letter(4 + index)].width = 9
    _head_block(sheet, "项目投资现金流量表", "附表8", "单位：万元", 25, title_size=14)
    _grid_header(sheet, 3, 1, "序号", down=2)
    _grid_header(sheet, 3, 2, "项目", down=2)
    _grid_header(sheet, 3, 3, "合计", down=2)
    _finance_year_header(sheet, period, 4, label="计算期")

    years = int(period.get("totalYears") or 22)
    flows = finance.get("cashflows") or {}
    block = flows.get("allInvestment") or {}
    revenue = _num_series((finance.get("revenue") or {}).get("total"))
    operating = _num_series((finance.get("cost") or {}).get("operatingCost"))
    vat = _num_series((finance.get("cost") or {}).get("vatPayable"))
    surcharge = _num_series((finance.get("cost") or {}).get("surchargeTotal"))
    tax = _num_series((finance.get("cost") or {}).get("adjustedIncomeTax"))
    inflow = _cf_values(block, "inflow", years)
    outflow = _cf_values(block, "outflow", years)
    net = [round(inflow[j] - outflow[j], 2) for j in range(years)]
    net_after = [round(net[j] - tax[j], 2) for j in range(years)]
    cumulative, cumulative_after = [], []
    running = running_after = 0.0
    for j in range(years):
        running = round(running + net[j], 2)
        running_after = round(running_after + net_after[j], 2)
        cumulative.append(running)
        cumulative_after.append(running_after)
    invest = [0.0] * years
    plan = ((estimate.get("tables") or {}).get("yearPlan") or {}).get("rows") or []
    for index, item in enumerate(plan):
        if index < years:
            invest[index] = _num(item.get("constructionInvestment")) or 0.0
    indicators = finance.get("indicators") or {}

    def line(row, serial, name, values, formula=None, bold=False, total=True):
        _put(sheet, row, 1, serial, bold=bold)
        _put(sheet, row, 2, name, align="l", bold=bold)
        _row_series(sheet, row, values, 4, FMT_AMT1, formula, total_col=3 if total else None)
        sheet.row_dimensions[row].height = 20

    residual = [round(inflow[j] - revenue[j], 2) for j in range(years)]

    def _a4(ref):
        """附表4 的年列自 H 起，本表自 D 起，跨表引用要右移 4 列。"""
        return get_column_letter(column_index_from_string(ref) + 4)

    line(5, "1", "现金流入", inflow, lambda ref: "={0}6+{0}7".format(ref), True)
    line(6, "1.1", "营业收入", revenue,
         lambda ref: "='附表4 -项目收入、增值税及附加表'!{}5".format(_a4(ref)))
    line(7, "1.2", "回收固定资产余值", residual, None)
    line(8, "2", "现金流出", outflow, lambda ref: "=SUM({0}9:{0}12)".format(ref), True)
    line(9, "2.1", "建设投资", invest, None)
    line(10, "2.2", "经营成本", operating, lambda ref: "='附表5-总成本费用表'!{0}11".format(ref))
    line(11, "2.3", "增值税、税金及附加", [round(vat[j] + surcharge[j], 2) for j in range(years)],
         lambda ref: "='附表4 -项目收入、增值税及附加表'!{0}12+'附表4 -项目收入、增值税及附加表'!{0}15".format(_a4(ref)))
    line(12, "2.4", "维持运营投资", [0.0] * years)
    line(13, "3", "所得税前净现金流量", net, lambda ref: "={0}5-{0}8".format(ref), True)
    line(14, "4", "累计所得税前净现金流量", cumulative, None, total=False)
    line(15, "5", "调整所得税", tax, None)
    line(16, "6", "所得税后净现金流量", net_after, lambda ref: "={0}13-{0}15".format(ref), True)
    line(17, "7", "累计所得税后净现金流量", cumulative_after, None, total=False)

    _put(sheet, 18, 2, "计算指标", bold=True, align="l")
    _put(sheet, 19, 2, "财务内部收益率（FIRR）", align="l")
    _f(sheet, 19, 3, "=IRR(D16:Y16)", _num(indicators.get("firrAll")), FMT_PCT, bold=True)
    _put(sheet, 20, 2, "财务净现值（FNPV）", align="l")
    _f(sheet, 20, 3, "=NPV(3.5%,D16:Y16)", _num(indicators.get("fnpvAll")), FMT_AMT2, bold=True)
    _put(sheet, 21, 2, "静态投资回收期（含建设期）", align="l")
    _put(sheet, 21, 3, _num(indicators.get("ptAll")), FMT_AMT2, bold=True)
    for row in (18, 19, 20, 21):
        sheet.row_dimensions[row].height = 20
    _side_note(sheet, 4, 22, "注：所得税后口径；折现率 3.5%；FIRR、FNPV 由本表净现金流量重算，与正文第 8 章同源。", 23)
    print_setup(sheet, title_rows="1:4", portrait=False, fit_height=0)
    sheet.page_setup.paperSize = sheet.PAPERSIZE_A3
    sheet.print_area = "A1:Y23"


def _finance_equity_sheet(workbook, finance, period):
    sheet = workbook.create_sheet("附表9-现金流量表 (资本金)")
    sheet.column_dimensions["A"].width = 6
    sheet.column_dimensions["B"].width = 34
    sheet.column_dimensions["C"].width = 13
    for index in range(22):
        sheet.column_dimensions[get_column_letter(4 + index)].width = 9
    _head_block(sheet, "项目资本金现金流量表", "附表9", "单位：万元", 25, title_size=14)
    _grid_header(sheet, 3, 1, "序号", down=2)
    _grid_header(sheet, 3, 2, "项目", down=2)
    _grid_header(sheet, 3, 3, "合计", down=2)
    _finance_year_header(sheet, period, 4, label="计算期")
    note = ("本项目资金筹措为政策性补助，资本金为 0，资本金现金流量表不适用；"
            "如后续引入社会资本，按本表式及《建设项目经济评价方法与参数》补算资本金 FIRR、FNPV 与回收期。")
    _merge(sheet, 5, 1, 7, 25, note, align="l", border=False, size=10)
    sheet.row_dimensions[5].height = 30
    print_setup(sheet, title_rows="1:4", portrait=False, fit_height=0)
    sheet.page_setup.paperSize = sheet.PAPERSIZE_A3
    sheet.print_area = "A1:Y7"


def _finance_plan_sheet(workbook, finance, estimate, period):
    sheet = workbook.create_sheet("附表8-财务计划现金流量表")
    sheet.column_dimensions["A"].width = 6
    sheet.column_dimensions["B"].width = 34
    sheet.column_dimensions["C"].width = 13
    for index in range(22):
        sheet.column_dimensions[get_column_letter(4 + index)].width = 9
    _head_block(sheet, "财务计划现金流量表", "附表8", "单位：万元", 25, title_size=14)
    _grid_header(sheet, 3, 1, "序号", down=2)
    _grid_header(sheet, 3, 2, "项目", down=2)
    _grid_header(sheet, 3, 3, "合计", down=2)
    _finance_year_header(sheet, period, 4, label="计算期")

    years = int(period.get("totalYears") or 22)
    revenue = _num_series((finance.get("revenue") or {}).get("total"))
    operating = _num_series((finance.get("cost") or {}).get("operatingCost"))
    vat = _num_series((finance.get("cost") or {}).get("vatPayable"))
    surcharge = _num_series((finance.get("cost") or {}).get("surchargeTotal"))
    tax = _num_series((finance.get("cost") or {}).get("incomeTax"))
    subsidy_total = _num((estimate.get("funding") or {}).get("subsidy")) or 0.0
    invest = [0.0] * years
    plan = ((estimate.get("tables") or {}).get("yearPlan") or {}).get("rows") or []
    for index, item in enumerate(plan):
        if index < years:
            invest[index] = _num(item.get("constructionInvestment")) or 0.0
    build = int(period.get("constructionYears") or 2)
    subsidy = list(invest)
    if build:
        subsidy[build - 1] = round(subsidy_total - sum(subsidy[:build - 1]), 2)

    def line(row, serial, name, values, formula=None, bold=False):
        _put(sheet, row, 1, serial, bold=bold)
        _put(sheet, row, 2, name, align="l", bold=bold)
        _row_series(sheet, row, values, 4, FMT_AMT1, formula, total_col=3)
        sheet.row_dimensions[row].height = 20

    line(5, "一", "资金来源", [round(subsidy[j] + revenue[j], 2) for j in range(years)],
         lambda ref: "=SUM({0}6:{0}9)".format(ref), True)
    line(6, "1", "财政资金", subsidy)
    line(7, "2", "政府专项债", [0.0] * years)
    line(8, "3", "营业收入", revenue)
    line(9, "4", "回收资产余值", [0.0] * years)
    line(10, "二", "资金运用",
         [round(invest[j] + operating[j] + vat[j] + surcharge[j] + tax[j], 2) for j in range(years)],
         lambda ref: "=SUM({0}11:{0}16)".format(ref), True)
    line(11, "1", "建设投资", invest)
    line(12, "2", "经营成本", operating)
    line(13, "3", "增值税及附加", [round(vat[j] + surcharge[j], 2) for j in range(years)])
    line(14, "4", "所得税", tax)
    line(15, "5", "借款还本", [0.0] * years)
    line(16, "6", "借款付息", [0.0] * years)
    net = [round(subsidy[j] + revenue[j] - (invest[j] + operating[j] + vat[j] + surcharge[j] + tax[j]), 2)
           for j in range(years)]
    cumulative, running = [], 0.0
    for value in net:
        running = round(running + value, 2)
        cumulative.append(running)
    line(17, "三", "净现金流量", net, lambda ref: "={0}5-{0}10".format(ref), True)
    _put(sheet, 18, 1, "四", bold=True)
    _put(sheet, 18, 2, "累计净现金流量", bold=True, align="l")
    _row_series(sheet, 18, cumulative, 4, FMT_AMT1)
    sheet.row_dimensions[18].height = 20
    _side_note(sheet, 4, 22, "注：本项目无借款，借款还本与付息为 0；财政资金与营业收入同口径计入资金来源。", 20)
    print_setup(sheet, title_rows="1:4", portrait=False, fit_height=0)
    sheet.page_setup.paperSize = sheet.PAPERSIZE_A3
    sheet.print_area = "A1:Y20"


def _bep_values(finance, period):
    cost = finance.get("cost") or {}
    surcharges = _num_series(cost.get("surchargeTotal"))
    vat = _num_series(cost.get("vatPayable"))
    revenue = _num_series((finance.get("revenue") or {}).get("total"))
    base = round(sum(revenue) - sum(vat) - sum(surcharges), 2)
    items = cost.get("items") or {}
    fixed = round(sum(_num_series(items.get("labor"))) + sum(_num_series(items.get("management")))
                 + sum(_num_series(items.get("repair"))) + sum(_num_series(items.get("other")))
                 + sum(_num_series(cost.get("depreciation"))) + sum(_num_series(cost.get("interest"))), 2)
    variable = round(sum(_num_series(items.get("utilities"))), 2)
    bep = round(fixed / (base - variable), 4) if (base - variable) else 0.0
    return base, fixed, variable, bep


def _finance_bep_sheet(workbook, finance, period):
    sheet = workbook.create_sheet("B4盈亏平衡分析")
    sheet.column_dimensions["A"].width = 12
    for key in ("B", "C", "D", "E"):
        sheet.column_dimensions[key].width = 18
    _head_block(sheet, "盈亏平衡分析", "附表B4", "单位：万元", 5, title_size=14)
    for index, text in enumerate(("变化率", "营业收入（税后）", "固定成本", "可变成本", "总成本费用")):
        _put(sheet, 3, index + 1, text, bold=True, size=10)
    sheet.row_dimensions[3].height = 34
    base, fixed, variable, bep = _bep_values(finance, period)
    for index in range(10):
        row = 4 + index
        rate = index / 10.0
        _put(sheet, row, 1, rate, FMT_PCT0)
        _put(sheet, row, 2, round(base * rate, 2), FMT_AMT2)
        _put(sheet, row, 3, fixed, FMT_AMT2)
        _put(sheet, row, 4, round(variable * rate, 2), FMT_AMT2)
        _put(sheet, row, 5, round(fixed + variable * rate, 2), FMT_AMT2)
        sheet.row_dimensions[row].height = 18
    _put(sheet, 15, 1, "盈亏平衡点", bold=True, align="l")
    _put(sheet, 15, 2, bep, FMT_PCT, bold=True)
    _merge(sheet, 15, 3, 15, 5, "固定成本 ÷（营业收入 − 可变成本）", align="l", border=False, size=9)
    _side_note(sheet, 1, 5, "注：盈亏平衡点按运营期合计口径测算；营业收入为扣除增值税及附加后的金额。", 17)
    print_setup(sheet, title_rows="1:3", portrait=True, fit_height=1)
    sheet.print_area = "A1:E17"


def _finance_sensitivity_sheet(workbook, finance, period):
    sheet = workbook.create_sheet("B5敏感性分析")
    sheet.column_dimensions["A"].width = 16
    for key in ("B", "C", "D", "E", "F", "G"):
        sheet.column_dimensions[key].width = 14
    _head_block(sheet, "敏感性分析", "附表B5", "单位：万元", 7, title_size=14)
    heads = ("敏感因素", "变化幅度", "内部收益率", "财务净现值", "敏感度系数", "临界点", "备注")
    for index, text in enumerate(heads):
        _put(sheet, 3, index + 1, text, bold=True, size=10)
    sheet.row_dimensions[3].height = 32
    sensitivity = finance.get("sensitivity") or {}
    base = sensitivity.get("base") or {}
    day = _num(base.get("firr")) or 0.0
    _put(sheet, 4, 1, "基本方案", align="l")
    _put(sheet, 4, 2, "0%")
    _put(sheet, 4, 3, day, FMT_PCT)
    _put(sheet, 4, 4, _num(base.get("fnpv")), FMT_AMT2)
    _put(sheet, 4, 5, "—")
    _put(sheet, 4, 6, "—")
    _put(sheet, 4, 7, "基准", align="l")
    row = 5
    for factor in sensitivity.get("factors") or []:
        for item in factor.get("rows") or []:
            delta = _num(item.get("deltaPercent")) or 0.0
            if abs(delta) < 1e-9:
                continue
            firr = _num(item.get("firr")) or 0.0
            coefficient = ((firr - day) / day / (delta / 100.0)) if day else 0.0
            _put(sheet, row, 1, factor.get("name"), align="l")
            _put(sheet, row, 2, "{:+.0f}%".format(delta))
            _put(sheet, row, 3, firr, FMT_PCT)
            _put(sheet, row, 4, _num(item.get("fnpv")), FMT_AMT2)
            _put(sheet, row, 5, round(coefficient, 2), FMT_AMT2)
            zero = factor.get("fnpvZeroDeltaPercent")
            _put(sheet, row, 6, "FNPV=0 在 {}% 变化时".format(zero) if zero else "测算区间内未出现",
                 align="l", size=9)
            _put(sheet, row, 7, "临界点" if zero else "", align="l", size=9)
            sheet.row_dimensions[row].height = 20
            row += 1

    row += 1
    _merge(sheet, row, 1, row, 7, "测算矩阵（单因素，其余因素冻结）", bold=True, align="l", border=False, size=10)
    row += 1
    for index, text in enumerate(("指标", "-10%", "0", "+10%")):
        _put(sheet, row, index + 1, text, bold=True, size=10)
    variations = sensitivity.get("variations") or [-10, 0, 10]
    for index, text in enumerate(variations):
        pass
    row += 1
    for label, key, fmt in (("内部收益率", "firr", FMT_PCT), ("财务净现值", "fnpv", FMT_AMT2),
                            ("投资回收期", "pt", FMT_AMT2)):
        _put(sheet, row, 1, label, align="l")
        for column, variation in enumerate(variations, start=2):
            value = None
            for factor in sensitivity.get("factors") or []:
                for item in factor.get("rows") or []:
                    if abs((_num(item.get("deltaPercent")) or 0.0) - float(variation)) < 1e-9:
                        value = item.get(key)
                        break
                if value is not None:
                    break
            _put(sheet, row, column, _num(value), fmt)
        sheet.row_dimensions[row].height = 20
        row += 1
    _side_note(sheet, 1, 7, "注：单因素法，仅调整该因素分量、其余冻结；临界点为 FNPV 归零时的因素变化幅度。", row)
    print_setup(sheet, title_rows="1:3", portrait=True, fit_height=1)
    sheet.print_area = "A1:G{}".format(row)


def _finance_indicator_sheet(workbook, finance, estimate, period):
    sheet = workbook.create_sheet("经济技术指标表")
    sheet.column_dimensions["A"].width = 9
    sheet.column_dimensions["B"].width = 34
    sheet.column_dimensions["C"].width = 10
    sheet.column_dimensions["D"].width = 16
    sheet.column_dimensions["E"].width = 18
    _head_block(sheet, "主要经济指标汇总表", "", "", 5, title_size=16)
    for index, text in enumerate(("序号", "项目名称", "单位", "指标", "备注")):
        _put(sheet, 3, index + 1, text, bold=True, size=10, fill=BAND_FILL)
        sheet.cell(row=3, column=index + 1).font = Font(name=F_TNR, size=10, bold=True, color="FFFFFF")
    sheet.row_dimensions[3].height = 24

    indicators = finance.get("indicators") or {}
    cost = finance.get("cost") or {}
    tables = finance.get("tables") or {}
    years = int(period.get("totalYears") or 22)
    build = int(period.get("constructionYears") or 2)
    total = _num(estimate.get("totalInvestment")) or 0.0
    full = tables.get("fullCostStructure") or {}
    operation = tables.get("costStructure") or {}
    revenue_struct = tables.get("revenueStructure") or {}
    profit = [_num_series((finance.get("revenue") or {}).get("total"))[j]
              - _num_series(cost.get("vatPayable"))[j]
              - _num_series(cost.get("surchargeTotal"))[j]
              - _num_series(cost.get("totalCost"))[j] for j in range(years)]
    net = [profit[j] - _num_series(cost.get("incomeTax"))[j] for j in range(years)]
    base, fixed, variable, bep = _bep_values(finance, period)

    rate = FMT_PCT
    rows = [
        ("1", "项目总投资", "万元", total, "", FMT_AMT2),
        ("1.1", "建设投资", "万元", _num(estimate.get("constructionInvestment")), "", FMT_AMT2),
        ("1.2", "建设期利息", "万元", _num(estimate.get("constructionInterest")), "", FMT_AMT2),
        ("1.3", "流动资金", "万元", 0.0, "本项目不计流动资金", FMT_AMT2),
        ("2", "成本", "—", None, "", None),
        ("2.1", "总成本费用", "万元", _num(full.get("avgAnnual")), "运营期平均", FMT_AMT2),
        ("2.2", "经营成本费用", "万元", _num(operation.get("avgAnnual")), "运营期平均", FMT_AMT2),
        ("3", "收入及利润", "—", None, "", None),
        ("3.1", "营业收入", "万元", _num(revenue_struct.get("avgAnnual")), "运营期平均", FMT_AMT2),
        ("3.2", "利润总额", "万元", _op_avg(profit, build), "运营期平均", FMT_AMT2),
        ("3.3", "净利润", "万元", _op_avg(net, build), "运营期平均", FMT_AMT2),
        ("4", "效益指标", "—", None, "", None),
        ("4.1", "项目投资内部收益率（税后）", "%", _num(indicators.get("firrAll")), "所得税后", rate),
        ("4.2", "项目静态投资回收期（税后）", "年", _num(indicators.get("ptAll")), "含建设期", FMT_AMT2),
        ("4.3", "项目总投资收益率", "%", _num(indicators.get("roi")), "ROI", rate),
        ("4.4", "项目资本金净利润率", "%", "不适用", "资本金为 0", FMT_PCT),
        ("4.5", "项目资本金内部收益率", "%", "不适用", "资本金为 0", FMT_PCT),
        ("5", "借款偿还期", "年", "不适用", "本项目无借款", FMT_AMT2),
        ("6", "盈亏平衡点", "—", bep, "固定成本 ÷（营业收入 − 可变成本）", FMT_PCT),
    ]
    row = 4
    for serial, name, unit, value, note, fmt in rows:
        _put(sheet, row, 1, serial, bold=not name.startswith(("项目", "成本", "收入", "效益")))
        _put(sheet, row, 2, name, align="l", bold=len(serial) == 1)
        _put(sheet, row, 3, unit)
        if value is None:
            _put(sheet, row, 4, "—")
        else:
            _put(sheet, row, 4, value, fmt)
        _put(sheet, row, 5, note, align="l", size=9)
        sheet.row_dimensions[row].height = 20
        row += 1
    _side_note(sheet, 1, 5, "注：指标与附表8、附表6、附表3 同源；资本金为 0，资本金口径指标不适用。", row)
    print_setup(sheet, title_rows="1:3", portrait=True, fit_height=1)
    sheet.print_area = "A1:E{}".format(row)


def export_deliverable_tables(workspace):
    workspace = Path(workspace)
    estimate = read_json(workspace / "过程文件/02-匡算/投资估算.json")
    finance = read_json(workspace / "过程文件/03-财务评价/财务测算.json")
    payload = read_json(workspace / "过程文件/01-上游取料/财务输入.json") or {}
    if not estimate or not finance:
        print("[export_office] 成品表缺失过程产物：投资估算/财务测算未就绪")
        return []
    outputs = []
    estimate_book = export_estimate_deliverable(estimate)
    estimate_path = workspace / "成果文件/成品版/匡算表/匡算表.xlsx"
    save(estimate_book, estimate_path, CACHE)
    outputs.append({"file": "匡算表.xlsx", "path": estimate_path.relative_to(workspace).as_posix(),
                    "sheets": estimate_book.sheetnames, "sha256": sha256_file(estimate_path)})
    finance_book = export_finance_deliverable(finance, payload, estimate)
    finance_path = workspace / "成果文件/成品版/财务评价表/财务评价表.xlsx"
    save(finance_book, finance_path, CACHE)
    outputs.append({"file": "财务评价表.xlsx", "path": finance_path.relative_to(workspace).as_posix(),
                    "sheets": finance_book.sheetnames, "sha256": sha256_file(finance_path)})
    target = workspace / "过程文件/06-渲染docx/成品表导出清单.json"
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text(json.dumps({"schema": "cardo-invest/deliverable-tables/v1",
                                  "generatedAt": None, "exports": outputs},
                                 ensure_ascii=False, indent=2, sort_keys=True), encoding="utf-8")
    print("[export_office] 成品表：匡算表.xlsx、财务评价表.xlsx")
    return outputs


def _sheet_xml_map(payload):
    """workbook.xml + rels → {sheet 名: zip 内路径}。

    openpyxl 写出的 Relationship 属性顺序是 Target/Type/Id，不能要求 Id 在前，
    必须逐个标签取出 Id 与 Target 再配对。
    """
    try:
        book = payload["xl/workbook.xml"].decode("utf-8")
        rels = payload["xl/_rels/workbook.xml.rels"].decode("utf-8")
    except KeyError:
        return {}
    rel_map = {}
    for tag in re.findall(r"<Relationship\b[^>]*/>", rels):
        rid = re.search(r'Id="([^"]+)"', tag)
        target = re.search(r'Target="([^"]+)"', tag)
        if rid and target:
            rel_map[rid.group(1)] = target.group(1)
    out = {}
    for tag in re.findall(r"<sheet\b[^>]*/>", book):
        name = re.search(r'name="([^"]*)"', tag)
        rid = re.search(r'r:id="([^"]+)"', tag)
        if not name or not rid:
            continue
        target = rel_map.get(rid.group(1), "").lstrip("/")
        if not target:
            continue
        out[_xml_unescape(name.group(1))] = target if target.startswith("xl/") else "xl/" + target
    return out


CELL_XML = re.compile(r'<c r="([A-Z]+[0-9]+)"([^>]*)>(.*?)</c>', re.S)


def _inject_values(xml, mapping):
    """给公式单元格补 <v> 缓存值，data_only 回读才拿得到数（openpyxl 不写缓存）。

    openpyxl 写公式时落的是空壳 `<f>…</f><v></v>`，必须先认这个空壳，
    否则「已含 <v>」的判据会把每个公式都跳过。
    """
    def repl(match):
        ref = match.group(1)
        if ref not in mapping:
            return match.group(0)
        body = match.group(3)
        if "<f>" not in body:
            return match.group(0)
        body = re.sub(r"<v\s*/>|<v></v>", "", body)
        if "<v>" in body:
            return match.group(0)
        attrs = re.sub(r'\s+t="[^"]*"', "", match.group(2))
        return '<c r="{}"{}>{}<v>{}</v></c>'.format(ref, attrs, body, _cell_number(mapping[ref]))
    return CELL_XML.sub(repl, xml)


def _cell_number(value):
    number = float(value)
    if abs(number) < 1e-12:
        return "0"
    return "{:.10g}".format(number)


def normalize_zip(path, cache=None):
    """固定 xlsx 的 zip 头时间戳与 docProps 修改时间，保证两次导出字节一致（内容不变）。

    cache={(sheet 名, 单元格): 数值}：顺带把公式缓存值写进 sheet XML。
    """
    path = Path(path)
    with zipfile.ZipFile(path) as source:
        items = [(info.filename, source.read(info.filename)) for info in source.infolist()]
    if cache:
        by_file = {}
        for (title, ref), value in cache.items():
            target = _sheet_xml_map(dict(items)).get(title)
            if target:
                by_file.setdefault(target, {})[ref] = value
        if by_file:
            patched = []
            for name, data in items:
                if name in by_file:
                    data = _inject_values(data.decode("utf-8"), by_file[name]).encode("utf-8")
                patched.append((name, data))
            items = patched
    fixed_core = re.compile(rb"(<dcterms:modified[^>]*>)[^<]*(</dcterms:modified>)")
    temporary = path.with_name(path.name + ".tmp")
    with zipfile.ZipFile(temporary, "w", zipfile.ZIP_DEFLATED) as target:
        for name, data in items:
            if name == "docProps/core.xml":
                data = fixed_core.sub(rb"\g<1>2026-01-01T00:00:00Z\g<2>", data)
            info = zipfile.ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            info.external_attr = 0o600 << 16
            target.writestr(info, data)
    os.replace(temporary, path)


def save(workbook, path, cache=None):
    path.parent.mkdir(parents=True, exist_ok=True)
    if path.exists():
        path.unlink()
    workbook.save(path)
    normalize_zip(path, cache)


def main(argv=None):
    parser = argparse.ArgumentParser(description="cardo-invest Office 导出（过程 JSON → xlsx）")
    parser.add_argument("--workspace", default=".")
    parser.add_argument("--deliverables", action="store_true",
                        help="只生成成品版两张表（匡算表、财务评价表），供 06 步调用")
    args = parser.parse_args(argv)
    workspace = Path(args.workspace).resolve()
    if args.deliverables:
        outputs = export_deliverable_tables(workspace)
        if len(outputs) != 2:
            return 2
        for item in outputs:
            print("  {} → {}".format(item["file"], item["path"]))
        return 0

    out_dir = workspace / "成果文件"
    out_dir.mkdir(parents=True, exist_ok=True)

    estimate = read_json(workspace / "过程文件/02-匡算/投资估算.json")
    finance = read_json(workspace / "过程文件/03-财务评价/财务测算.json")
    schemes = read_json(workspace / "过程文件/04-三方案比选/方案比选.json")
    payload = read_json(workspace / "过程文件/01-上游取料/财务输入.json") or {}
    missing = [name for name, payload in (("投资估算", estimate), ("财务测算", finance), ("方案比选", schemes))
               if not payload]
    if missing:
        print("[export_office] 缺过程产物：{}（先跑 02/03/04）".format(missing))
        return 2

    exports = []
    for title, builder, payload in (("投资估算.xlsx", export_estimate, estimate),
                                    ("财务评价.xlsx", lambda data, p=payload: export_finance(data, p), finance),
                                    ("方案比选.xlsx", export_schemes, schemes)):
        workbook = builder(payload)
        # 一成果一子目录：成果文件/内部件/<件名>/<件名>.xlsx
        path = out_dir / "内部件" / title.replace(".xlsx", "") / title
        save(workbook, path)
        exports.append({"file": title, "sheets": workbook.sheetnames,
                        "rows": sum(workbook[name].max_row for name in workbook.sheetnames),
                        "sha256": sha256_file(path)})

    pending_name = "待确认清单.xlsx"
    pending_workbook = export_pending(workspace)
    pending_path = out_dir / "内部件" / "待确认清单" / pending_name
    save(pending_workbook, pending_path)
    exports.append({"file": pending_name, "sheets": pending_workbook.sheetnames,
                    "rows": sum(pending_workbook[name].max_row for name in pending_workbook.sheetnames),
                    "sha256": sha256_file(pending_path)})

    # 表 10-1 的附件 5—6 是第 9 章引用的 AHP 源表。它们只躺在「输入文件」里时，
    # 成品说「已附」而交付目录里没有，读者 / 专家翻不到（R5 B-5）。
    # 这里真附进 成果文件/内部件/AHP/，导出清单同步记账；08 的闸门核实物存在。
    attachments = []
    for name, pattern in (("AHP权重表.xlsx", "AHP权重表*.xlsx"),
                          ("AHP方案评分结论.xlsx", "AHP方案评分结论*.xlsx")):
        # 输入件实际名带「_成品版」/「_溯源版」后缀，精确名匹配会全部落空，
        # 成品仍写「已附」而目录不存在。按前缀通配，并优先取成品版。
        candidates = sorted((workspace / "输入文件").rglob(pattern),
                            key=lambda path: (0 if "成品版" in path.name else 1, str(path)))
        source = next((candidate for candidate in candidates if candidate.is_file()), None)
        if source is None:
            continue
        attached = out_dir / "内部件" / "AHP" / name
        attached.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(source, attached)
        attachments.append({"file": name,
                            "path": attached.relative_to(workspace).as_posix(),
                            "source": source.relative_to(workspace).as_posix(),
                            "sha256": sha256_file(attached)})

    deliverable = out_dir / "成品版" / "项目投资决策建议书" / "项目投资决策建议书.docx"
    if not deliverable.is_file():
        deliverable = out_dir / "项目投资决策建议书.docx"
    manifest = {
        "schema": SCHEMA_LIST,
        "generatedAt": None,
        "exports": exports,
        "attachments": attachments,
        "docx": {"file": deliverable.name, "exists": deliverable.is_file(),
                 "sha256": sha256_file(deliverable) if deliverable.is_file() else None},
        "note": "四张 xlsx 只做格式化，不做二次计算；缺数写「—」不补 0。attachments 为表 10-1 附件 5—6 的实物副本。",
    }
    target = workspace / "过程文件/08-验证与边界/Office导出清单.json"
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text(json.dumps(manifest, ensure_ascii=False, indent=2, sort_keys=True), encoding="utf-8")
    print("[export_office] 内部件：{}".format("、".join(item["file"] for item in exports)))
    return 0


if __name__ == "__main__":
    sys.exit(main())
