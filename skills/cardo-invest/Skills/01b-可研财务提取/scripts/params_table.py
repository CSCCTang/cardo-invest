"""财务参数.xlsx 的模板管理与用户数据纠错。

两件事：
1. **模板由本技能管理** —— 空白模板随技能包走（assets/inputs/财务参数-空白.xlsx），
   01b 每次运行都会把它复制到 输入文件/财务参数-空白.xlsx，
   用户可以直接填、填完另存为 输入文件/财务参数.xlsx。
2. **对用户填的非标准数据做纠错** —— 用户手填难免有单位错、量纲错、
   百分数写成 8 而不是 0.08、千分位、货币符号、全角字符、列名顺序不同等。
   这里统一归一，并把「改了什么」记进纠错台账，让用户看得见。

原则：纠错只做**确定性可判定**的修正（单位换算、百分号、千分位、全角、
列名别名、空值归一）；**不做**任何猜测性修正（不猜用户意图、不补缺失值）。
"""

from __future__ import annotations

import re
import shutil
from decimal import Decimal
from pathlib import Path

# ---------------------------------------------------------------- 模板管理

TEMPLATE_NAME = "财务参数-空白.xlsx"
BLANK_NAME = "财务参数-空白.xlsx"
FILLED_NAME = "财务参数.xlsx"

# 9 张表的工作表名（与 01 取料的 find_sheet 匹配口径一致）
SHEETS = ["项目", "工程费", "其他费", "预备费", "建设期利息",
          "收入", "成本", "资金结构", "可研财务指标"]


def template_path(package_root):
    """技能包内的空白模板路径。"""
    return Path(package_root) / "assets" / "inputs" / TEMPLATE_NAME


def ensure_blank_template(workspace, package_root):
    """确保 输入文件/ 里有空白模板（用户可填的入口）。

    返回 (模板路径, 是否新建)
    """
    source = template_path(package_root)
    target = Path(workspace) / "输入文件" / BLANK_NAME
    target.parent.mkdir(parents=True, exist_ok=True)
    if target.exists():
        return target, False
    if not source.is_file():
        return None, False
    shutil.copyfile(source, target)
    return target, True


# ---------------------------------------------------------------- 纠错

# 百分数写法：8 / 8% / 0.08 都可能出现 → 统一为小数
_PERCENT_CELLS = {
    ("预备费", "预备费率"),
    ("建设期利息", "年利率"),
    ("成本", "年增长率"),
    ("资金结构", "借款年利率"),
    ("可研财务指标", "FNPV折现率"),
}
_PERCENT_SHEETS = ("收入",)

# 单位别名 → 统一键名（用户可能把表头写成同义词）
_COLUMN_ALIASES = {
    "金额(万元)": ("金额(万元)", "金额", "投资(万元)", "金额（万元）"),
    "年均收入(万元)": ("年均收入(万元)", "年均收入", "年收入(万元)"),
    "增值税率": ("增值税率", "税率", "增值税率(%)"),
    "年增长率": ("年增长率", "增长率"),
    "编号": ("编号", "代码", "序号"),
    "名称": ("名称", "工程名称", "费用名称", "收入项", "项目名称"),
}


def normalize_text(value):
    """全角转半角、去空白、统一常见符号。"""
    if value is None:
        return ""
    text = str(value)
    table = str.maketrans({
        "０": "0", "１": "1", "２": "2", "３": "3", "４": "4",
        "５": "5", "６": "6", "７": "7", "８": "8", "９": "9",
        "．": ".", "，": ",", "－": "-", "＋": "+", "％": "%",
        "（": "(", "）": ")", "：": ":", "　": " ",
    })
    text = text.translate(table)
    return re.sub(r"\s+", " ", text).strip()


def parse_number(value, percent=False):
    """把单元格解析成 Decimal。

    支持：4301.76 / 4,301.76 / 4301.76万元 / ￥4301.76 / 8% / 8
    percent=True 时：带 % 或数值 >1 → 视为百分数（8 → 0.08）
    返回 (Decimal, 纠错说明列表)
    """
    fixes = []
    text = normalize_text(value)
    if not text:
        return None, fixes
    if text in ("-", "—", "/", "不计取", "暂估", "N/A", "n/a"):
        return None, fixes

    had_percent = "%" in text
    text = text.replace("%", "").replace("￥", "").replace("¥", "")
    text = text.replace("万元", "").replace("元", "").replace("个", "")
    text = text.replace(",", "").replace(" ", "")
    match = re.fullmatch(r"-?\d+(\.\d+)?", text)
    if not match:
        return None, fixes
    try:
        number = Decimal(text)
    except Exception:
        return None, fixes
    if percent:
        if had_percent:
            if number > 1:
                number = number / Decimal(100)
                fixes.append("百分号写法：{}% → {}".format(
                    number * 100, number))
        elif number > 1:
            number = number / Decimal(100)
            fixes.append("按百分数解读：{} → {}".format(
                number * 100, number))
    return number, fixes


# 百分比字段：按「字段名」判定，而不是写死表名
_PERCENT_FIELDS = ("预备费率", "年利率", "借款年利率", "年增长率", "折现率",
                   "增值税率", "费率", "资本金比例", "还款比例")
# 万元字段：带「万元」或「(万元)」的金额字段
_AMOUNT_FIELDS = ("金额", "投资", "收入", "成本", "支出", "补助", "借款",
                  "资本金", "基数", "回收期", "FNPV", "总投资", "利润", "费用")


def parse_field_number(label, value):
    """按**字段语义**解析数值：自动识别百分比与金额单位。

    给 01 取料侧用 —— 用户手填的表里字段名千奇百怪
    （「预备费率」「年利率(%)」「金额(万元)」「总投资」…），
    按字段名判断比按表名判断更贴近用户的填写直觉。

    返回 (Decimal|None, 纠错说明列表)
    """
    fixes = []
    text = normalize_text(value)
    if not text:
        return None, fixes
    key = normalize_text(label)
    # 百分数字段：字段名含「率」「%」且不是「回收期」
    is_percent = ("%" in key or "率" in key) and "回收期" not in key
    if is_percent:
        # 明确写了 (%) 的，值可能写 8 或 0.08，两种都接受
        return parse_number(value, percent=True)
    return parse_number(value, percent=False)


def canonical_header(text):
    """表头归一：返回标准列名，或 None（不是已知列）。"""
    key = normalize_text(text)
    if not key:
        return None
    for standard, aliases in _COLUMN_ALIASES.items():
        if key in aliases:
            return standard
    return None


def fix_sheet_rows(sheet_name, rows):
    """对一张表做纠错归一。

    rows: [[cell, ...], ...]（不含表头）
    返回 (归一后的 rows, 纠错说明列表)
    """
    fixes = []
    header = [canonical_header(cell) for cell in (rows[0] if rows else [])]
    body = rows[1:] if rows else []

    out = []
    for row in body:
        new_row = list(row)
        for index, cell in enumerate(row):
            column = header[index] if index < len(header) else None
            if column is None:
                new_row[index] = normalize_text(cell)
                continue
            percent = (sheet_name, column) in _PERCENT_CELLS or \
                (sheet_name in _PERCENT_SHEETS and column == "增值税率")
            number, notes = parse_number(cell, percent=percent)
            if notes:
                fixes.append("{}!{} 第{}行：{}".format(
                    sheet_name, column, len(out) + 2, "；".join(notes)))
            if number is not None:
                new_row[index] = float(number)
            else:
                new_row[index] = normalize_text(cell)
        out.append(new_row)
    return out, fixes


def validate_budget_balance(engineering, other_fees, contingency, interest, total,
                            engineering_total=None, other_fees_total=None,
                            contingency_total=None):
    """资金平衡自检。

    优先用附表**大类行的权威值**互校：
        工程费用 + 工程建设其他费用 + 预备费 + 建设期利息 = 项目总投资
    没有大类行时，退回用抽到的明细之和与总投资比。

    返回 (是否通过, 说明)
    """
    if total is None:
        return True, "无总投资值，跳过平衡校验"
    parts = []
    if engineering_total is not None:
        parts.append(("工程费用", engineering_total))
    elif engineering:
        parts.append(("工程费明细", sum(engineering, Decimal(0))))
    if other_fees_total is not None:
        parts.append(("工程建设其他费用", other_fees_total))
    elif other_fees:
        parts.append(("其他费明细", sum(other_fees, Decimal(0))))
    if contingency_total is not None:
        parts.append(("预备费", contingency_total))
    elif contingency:
        parts.append(("预备费明细", sum(contingency, Decimal(0))))
    if interest:
        parts.append(("建设期利息", sum(interest, Decimal(0))))
    if not parts:
        return True, "无明细值，跳过平衡校验"
    subtotal = sum((value for _label, value in parts), Decimal(0))
    diff = abs(total - subtotal)
    detail = " + ".join("{} {}".format(label, value) for label, value in parts)
    if diff <= Decimal("1"):
        return True, "{} = {}，与总投资 {} 一致".format(detail, subtotal, total)
    return False, "{} = {}，与总投资 {} 相差 {} 万元".format(detail, subtotal, total, diff)


def project_name_from_dir(dir_name):
    """从**项目目录名**提取项目正式名称。

    项目目录名是提资方自己的命名，最权威 —— 比 PDF 文件名可靠
    （实测某企业片区的目录里放着别条道路的可研，按文件名取会得到
    「某路（A路-B路）市政道路工程项目」，项目名直接填错）。
    """
    import re as _re
    stem = dir_name.strip()
    stem = _re.sub(r"^\s*\d+(?:[\.\-]\d+)*[\.\-]?\s*", "", stem)
    stem = _re.sub(r"[（(][^）)]*(项目建议书|可行性研究报告|可研报告|附专家|盖章|审定|修订稿)[^）)]*[）)]", "", stem)
    stem = _re.sub(r"[（(]\s*\d{6,8}\s*[）)]", "", stem)      # 去 (20261022)
    stem = _re.sub(r"(项目建议书|可行性研究报告|可研报告)", "", stem)   # 可出现在中间
    stem = _re.sub(r"\d{6,8}$", "", stem)                            # 去尾部日期 20261022
    stem = stem.strip(" -·")
    return stem or dir_name.strip()
