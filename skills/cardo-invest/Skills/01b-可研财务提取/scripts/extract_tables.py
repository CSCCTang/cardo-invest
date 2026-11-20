"""可研/匡算附表（xlsx）→ 财务参数抽取。

目标文件形态：可研报告配套的「附表」工作簿，例如
  附表1工程总投资估算表 / 附表3-项目总投资使用计划与资金筹措表 /
  附表4-项目收入、增值税及附加表 / 附表5-总成本费用表

只从原表抄数，不补数。每个值记 {file, sheet, cell/row, quote} 出处。

表格形态（实测某街匡算 xlsx）：
  附表1  A=序号 B=费用名称 C=建筑工程费 D=安装工程费 E=设备及工器具购置费
         F=其他费 G=小计 H=单位 I=技术经济指标
         末尾有一张「投资构成汇总」：A=序号 B=名称 C=合计(万元) D=占总投资比例
  附表3  A=序号 B=项目 C=备注 ... G=合计 H=建设期年份
"""

from __future__ import annotations

import re
import sys
from decimal import Decimal
from pathlib import Path

# 运行时是 python313._pth 模式，脚本目录不进 sys.path：兄弟模块必须显式加进来
sys.path.insert(0, str(Path(__file__).resolve().parent))

from extract_feasibility import clean_line, fmt, to_decimal  # noqa: E402

# ---------------------------------------------------------------- 表名识别

ESTIMATE_SHEETS = ("附表1工程总投资估算表", "工程总投资匡算表", "工程总投资估算表",
                    "项目总投资估算表", "总投资估算表", "投资估算表")
FUNDING_SHEETS = ("附表3-项目总投资使用计划与资金筹措表", "项目总投资使用计划与资金筹措表",
                  "资金筹措表", "投资使用计划与资金筹措表")
REVENUE_SHEETS = ("附表4-项目收入、增值税及附加表", "项目收入、增值税及附加表",
                   "收入、增值税及附加表", "收入表")
COST_SHEETS = ("附表5-总成本费用表", "总成本费用表", "总成本表", "成本费用表")
METRIC_SHEETS = ("附表8-项目投资现金流量表", "附表8-财务计划现金流量表",
                 "项目投资现金流量表", "现金流量表")

# 附表1 里的大类行（序号列的值）
_CATEGORY = {
    "一": "engineering",
    "二": "otherFees",
    "三": "contingency",
    "四": "staticInvestment",
    "五": "constructionInterest",
    "六": "total",
}
# 汇总表里的名称 → 大类
_SUMMARY_MAP = (
    ("建筑安装工程费", "engineering"),
    ("工程建设其他费", "otherFees"),
    ("预备费", "contingency"),
    ("建设期利息", "constructionInterest"),
    ("估算总投资", "total"),
    ("项目总投资", "total"),
)


def _num(value):
    """单元格 → Decimal；非数值/空 → None。"""
    if value is None:
        return None
    if isinstance(value, (int, float, Decimal)):
        try:
            return Decimal(str(value)).quantize(Decimal("0.01"))
        except Exception:
            return None
    text = str(value).strip().replace(",", "")
    if not text or text in ("-", "—", "/", "不计取", "暂估"):
        return None
    try:
        return Decimal(text).quantize(Decimal("0.01"))
    except Exception:
        return None


def _find_sheet(workbook, candidates):
    """按候选名精确/包含匹配工作表，返回 (name, worksheet) 或 (None, None)。"""
    names = workbook.sheetnames
    for candidate in candidates:
        if candidate in names:
            return candidate, workbook[candidate]
    for candidate in candidates:
        for name in names:
            if candidate in name:
                return name, workbook[name]
    return None, None


def _cell(row, index):
    try:
        return row[index]
    except (IndexError, TypeError):
        return None


def _quote(values, limit=90):
    parts = [clean_line(str(v)) for v in values if v is not None and clean_line(str(v))]
    return " | ".join(parts)[:limit]


# ---------------------------------------------------------------- 附表1：投资构成汇总

def read_summary(worksheet, file_label, sheet_name):
    """读附表1 末尾的「投资构成汇总」两列表（名称 + 合计万元）。

    返回 {"engineering": Decimal, ...} + {"_rows": {key: {value, source}}}
    """
    result = {}
    rows = list(worksheet.iter_rows(values_only=True))
    for row in rows:
        if not row or len(row) < 3:
            continue
        seq = clean_line(str(_cell(row, 0) or ""))
        name = clean_line(str(_cell(row, 1) or ""))
        value = _num(_cell(row, 2))
        if not name or value is None:
            continue
        for key, label in _SUMMARY_MAP:
            if label in name and label not in result:
                result[label] = value
                result.setdefault("_rows", {})[label] = {
                    "value": value,
                    "source": {
                        "file": file_label, "sheet": sheet_name,
                        "row": _row_index(row), "quote": _quote(row),
                    },
                }
                break
    return result


def _row_index(row):
    """openpyxl 的 values_only 行没有行号，用 enumerate 补（在调用处传）。"""
    return row.get("_rowno") if isinstance(row, dict) else None


def read_estimate_detail(worksheet, file_label, sheet_name):
    """读附表1 明细：按大类收集「(一)(二)…/1/1.1」层级的费用项与金额。

    层级判定：序号列形如 一/二/（一）/1/1.1 —— 只取「(一)」与「1」两级作为明细项，
    更深的 1.1 属于子项，避免重复计数。
    """
    items = {"engineering": [], "otherFees": [], "contingency": []}
    category = None
    rows = list(worksheet.iter_rows(values_only=True))
    for index, row in enumerate(rows, start=1):
        if not row or len(row) < 7:
            continue
        seq = clean_line(str(_cell(row, 0) or ""))
        name = clean_line(str(_cell(row, 1) or ""))
        if not name:
            continue
        # 大类行
        if seq in ("一", "二", "三"):
            category = _CATEGORY.get(seq)
            continue
        if seq in ("四", "五", "六"):
            category = None
            continue
        # 只取「(一)」与「1」两级
        is_top = bool(re.fullmatch(r"[（(][一二三四五六七八九十]+[）)]", seq))
        is_item = bool(re.fullmatch(r"\d+", seq))
        if not (is_top or is_item):
            continue
        # 金额：工程费取 G 列（小计）；其他费取 F 列（其他费）。
        # 建设管理费一行 F=537.59、G=477.07（下浮后小计），而附表1「工程建设其他费」
        # 合计 2878.53 与预备费 2161.1 都是按 F 列口径算的；取 G 会少 60.52、
        # 进而让总投资与资金筹措差 66.57，匡算配不平。
        amount = _num(_cell(row, 5 if category == "otherFees" else 6))
        if amount is None or amount == 0:
            continue
        if category not in items:
            continue
        items[category].append({
            "name": name,
            "amount": amount,
            "level": "top" if is_top else "item",
            "source": {
                "file": file_label, "sheet": sheet_name,
                "row": index, "quote": _quote(row),
            },
        })
    # 工程费同时存在「(一)」大类与「1」明细两级：任一有大类行，就只留大类，
    # 否则父子同列会把同一笔钱加两遍（本项目会从 18732.46 涨到 45921.50）。
    for key, rows in items.items():
        tops = [x for x in rows if x["level"] == "top"]
        if tops:
            items[key] = tops
    return items


def pick_by_total(rows, target):
    """按「附表的合计值」从候选行里挑一组，使金额之和正好等于合计数。

    多张表拼在一起时，同一笔费既有汇总项（建设管理费 537.59）又有子项
    （建设单位管理费 277.72 + 监理 259.87）。靠名字猜父子不可靠
    （勘察费/工程检测费同为 149.86），改用附表的权威合计做精确匹配，
    优先取条目少的那组（汇总项），避免重复计费。
    """
    if not rows or not target:
        return rows
    cents = [int(round(float(r["amount"]) * 100)) for r in rows]
    need = int(round(float(target) * 100))
    order = sorted(range(len(rows)), key=lambda i: -cents[i])
    best = None
    budget = [200000]

    def dfs(pos, remain, chosen):
        nonlocal best
        budget[0] -= 1
        if budget[0] <= 0:
            return
        if remain == 0:
            if best is None or len(chosen) < len(best):
                best = list(chosen)
            return
        if pos >= len(order) or remain < 0:
            return
        if best is not None and len(chosen) >= len(best):
            return
        for k in range(pos, len(order)):
            idx = order[k]
            if cents[idx] > remain:
                continue
            chosen.append(idx)
            dfs(k + 1, remain - cents[idx], chosen)
            chosen.pop()
            if best is not None and len(best) <= 2:
                return

    dfs(0, need, [])
    if not best:
        return rows
    return [r for idx, r in enumerate(rows) if idx in set(best)]


def read_contingency_rate(worksheet, file_label, sheet_name):
    """从预备费行备注里取费率，如「（工程费+工程建设其他费-土地费）×10%」。"""
    for index, row in enumerate(worksheet.iter_rows(values_only=True), start=1):
        if not row or len(row) < 3:
            continue
        seq = clean_line(str(_cell(row, 0) or ""))
        name = clean_line(str(_cell(row, 1) or ""))
        if "预备费" not in name:
            continue
        note = ""
        for value in row[2:]:
            note = clean_line(str(value or ""))
            if note:
                break
        match = re.search(r"(\d+(?:\.\d+)?)\s*%", note)
        if match:
            rate = to_decimal(match.group(1))
            if rate is not None:
                return {"rate": rate / Decimal(100), "source": {
                    "file": file_label, "sheet": sheet_name, "row": index,
                    "quote": _quote(row)}}
    return None


# ---------------------------------------------------------------- 附表3：资金筹措

def read_funding(worksheet, file_label, sheet_name):
    """读资金筹措：资本金/借款/补助/利率。

    识别「财政资金」「专项债」「银行贷款」等来源名：
      财政资金 → 补助（subsidy）
      专项债 / 银行贷款 / 借款 → 借款（loan）
    资本金单独识别。
    """
    out = {}
    rows = list(worksheet.iter_rows(values_only=True))
    for index, row in enumerate(rows, start=1):
        if not row or len(row) < 3:
            continue
        seq = clean_line(str(_cell(row, 0) or ""))
        name = clean_line(str(_cell(row, 1) or ""))
        note = clean_line(str(_cell(row, 2) or ""))
        # 金额列：找第一个非空数值列（合计列在中部，逐年列在后面）
        amount = None
        for value in row[3:]:
            candidate = _num(value)
            if candidate is not None:
                amount = candidate
                break
        source = {"file": file_label, "sheet": sheet_name, "row": index, "quote": _quote(row)}

        if "资本金" in name and amount is not None and "capital" not in out:
            out["capital"] = {"amount": amount, "source": source}
            continue
        if amount is None:
            continue
        if re.search(r"专项债|银行贷款|借款|贷款", name) and "loan" not in out:
            out["loan"] = {"amount": amount, "source": source}
            continue
        if re.search(r"财政资金|政府投资|补助|自筹", name) and "subsidy" not in out:
            out["subsidy"] = {"amount": amount, "source": source}
            continue
        # 利率行
        if "利率" in name or "利率" in note:
            for value in row[3:]:
                rate = _num(value)
                if rate is not None and rate < 1:
                    out["loanRate"] = {"rate": rate, "source": source}
                    break
    return out


# ---------------------------------------------------------------- 附表4/5：收入与成本

def read_revenue(worksheet, file_label, sheet_name):
    """读附表4 收入表：项目名 + **年均**收入。

    附表4 的「合计」列是**全运营期合计**，不是年值：早先把合计当成年收入
    （27965.2 万元/年），FIRR 被抬到 123%，08 的合理区间闸门才拦住。
    年均收入 = 合计 ÷ 运营期年数；只收带数字序号的明细行，跳过「一 收入合计」
    父行与单价/数量/使用率等测算参数行，读到「二 增值税、税金与附加」为止。
    """
    items = []
    rows = list(worksheet.iter_rows(values_only=True))
    if not rows:
        return []
    total_col = None
    for row in rows[:12]:
        for index, cell in enumerate(row):
            if clean_line(str(cell or "")).replace(" ", "") == "合计":
                total_col = index
                break
        if total_col is not None:
            break
    if total_col is None:
        return []
    # 运营期年数：年表头整年计 1、半年计 0.5
    years = Decimal("0")
    for row in rows[:12]:
        texts = list(row)[total_col + 1:]
        matched = 0
        for cell in texts:
            if isinstance(cell, (int, float)) and 2000 <= float(cell) <= 2100:
                matched += 1
            elif re.fullmatch(r"20\d{2}(上|下)?半年?", clean_line(str(cell or ""))):
                matched += 1
        if matched >= 3:
            for cell in texts:
                text = clean_line(str(cell or ""))
                if isinstance(cell, (int, float)) and 2000 <= float(cell) <= 2100:
                    years += Decimal("1")
                elif re.fullmatch(r"20\d{2}", text):
                    years += Decimal("1")
                elif re.fullmatch(r"20\d{2}(上|下)半年", text):
                    years += Decimal("0.5")
            break
    if years <= 0:
        years = Decimal("1")
    started = False
    for index, row in enumerate(rows, start=1):
        if not row or len(row) <= total_col:
            continue
        seq = clean_line(str(row[0] or ""))
        name = clean_line(str(row[1] or ""))
        if "收入合计" in name:
            started = True
            continue
        if not started:
            continue
        if seq in ("二", "三") and ("增值税" in name or "税金" in name):
            break
        # 明细行才带数字序号（1 / 1.1）；空白序号是单价/数量/使用率等参数行
        if not re.fullmatch(r"\d+(?:\.\d+)*", seq) or not name:
            continue
        total = _num(row[total_col])
        if total is None or total == 0:
            continue
        items.append({
            "name": name,
            "annualAmount": (total / years).quantize(Decimal("0.01")),
            "totalAmount": total,
            "source": {"file": file_label, "sheet": sheet_name, "row": index,
                       "quote": _quote(row)},
        })
    return items


def read_cost(worksheet, file_label, sheet_name):
    """读附表5 总成本费用表：尽量抽职工人数/薪酬；抽不到就返回空。"""
    out = {}
    rows = list(worksheet.iter_rows(values_only=True))
    for index, row in enumerate(rows, start=1):
        if not row or len(row) < 2:
            continue
        name = clean_line(str(_cell(row, 0) or ""))
        values = [_num(v) for v in row[1:]]
        values = [v for v in values if v is not None]
        if not values:
            continue
        source = {"file": file_label, "sheet": sheet_name, "row": index, "quote": _quote(row)}
        if ("人数" in name or "定员" in name or "职工" in name) and "crew" not in out:
            out["crew"] = {"value": values[0], "source": source}
        elif ("工资" in name or "薪酬" in name) and "salaryWan" not in out:
            out["salaryWan"] = {"value": values[0], "source": source}
    return out


# ---------------------------------------------------------------- 汇总入口

def extract_from_workbook(path, file_label=None):
    """从一个 xlsx 里抽全部财务信息。找不到任何附表则返回 None。"""
    try:
        import openpyxl
    except ImportError:
        return None
    label = file_label or path.name
    try:
        workbook = openpyxl.load_workbook(path, data_only=True, read_only=False)
    except Exception:
        return None

    out = {"file": label, "summary": {}, "detail": {}, "funding": {},
           "contingencyRate": None, "revenue": [], "cost": {}, "sheets": {}}

    name, sheet = _find_sheet(workbook, ESTIMATE_SHEETS)
    if sheet is not None:
        out["sheets"]["estimate"] = name
        out["summary"] = read_summary(sheet, label, name)
        out["detail"] = read_estimate_detail(sheet, label, name)
        out["contingencyRate"] = read_contingency_rate(sheet, label, name)

    name, sheet = _find_sheet(workbook, FUNDING_SHEETS)
    if sheet is not None:
        out["sheets"]["funding"] = name
        out["funding"] = read_funding(sheet, label, name)

    name, sheet = _find_sheet(workbook, REVENUE_SHEETS)
    if sheet is not None:
        out["sheets"]["revenue"] = name
        out["revenue"] = read_revenue(sheet, label, name)

    name, sheet = _find_sheet(workbook, COST_SHEETS)
    if sheet is not None:
        out["sheets"]["cost"] = name
        out["cost"] = read_cost(sheet, label, name)

    if not out["sheets"]:
        return None
    return out
