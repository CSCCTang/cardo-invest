"""01 上游取料（前段）：把 Office 输入与上游机读 JSON 装配成「引擎项目」。

输入（相对工作区根，可被工作区参数.json 覆盖）
  输入文件/                   投递区：用户可直接放文件、文件夹或 zip；脚本递归发现并按文件名关键词归类
  过程文件/00-总控/上游机读件/  上游过程 JSON 归档（权重/评分/现状指标/事实库/槽位），不属于用户输入

输出
  过程文件/00-总控/引擎项目/  引擎约定布局：04_成果报告/ + 05_过程/{AHP,低空解析,投资决策}/
  过程文件/01-上游取料/取料清单.json

规则
  - 只做格式装配，不补数、不推算：缺什么就跳过并在清单里登记，交给 collect_inputs 出 gaps；
  - 财务参数.xlsx 是唯一必需输入，缺失 → exit 3（待用户补，不拿经验值顶替）；
  - AHP 不重算：只把上游权重/评分转成契约键，权重值原样保留；
  - 路径一律相对工作区根，不写盘符。

用法：python prepare_inputs.py --workspace . [--scenario full|gaps] [--no-timestamp]
退出码：0 装配完成；2 工作区/输入目录不合法；3 缺必需财务参数（待补）。
"""
from __future__ import annotations

import argparse
import json
import os
import shutil
import sys
import tempfile
from decimal import Decimal, InvalidOperation
from pathlib import Path

SCHEMA_LIST = "cardo-invest/intake-manifest/v1"

# 上游 AHP 方案名 → 本包标准三方案（04 步 方案参数表.yaml 的 modes）的对应序。
# 上游 CardoAHP 产出的是「方案一/二/三」这类中性名，未必能对上标准三方案；
# 对不上的按「排名表里的名次」保序（见 parse_ahp_scores_office）。
# 注意：此表只决定 S1/S2/S3 的编号先后，不改方案名与得分。
SCHEME_ORDER = {
    "拆除重建": 0, "方案一": 0, "全面改造": 0, "整体拆除重建": 0,
    "微更新": 1, "方案二": 1, "渐进式更新": 1, "有机更新": 1,
    "功能置换": 2, "方案三": 2, "业态置换": 2, "功能更新": 2,
}


# ---------------------------------------------------------------- 基础 IO

def read_json(path):
    try:
        return json.loads(Path(path).read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return None


def normalize_zip(path):
    """把 docx/xlsx 的 zip 头时间戳固定，保证两次跑字节一致（内容不变）。"""
    import zipfile

    path = Path(path)
    with zipfile.ZipFile(path) as source:
        items = [(info.filename, source.read(info.filename)) for info in source.infolist()]
    temporary = path.with_name(path.name + ".tmp")
    with zipfile.ZipFile(temporary, "w", zipfile.ZIP_DEFLATED) as target:
        for name, data in items:
            info = zipfile.ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            info.external_attr = 0o600 << 16
            target.writestr(info, data)
    os.replace(temporary, path)


def copy_status_report(source, target):
    """搬上游现状报告；只剔除上游自带的『（数据来源：…）』机读标注行（不改编正文、不改数）。"""
    shutil.copy2(source, target)
    try:
        import docx
    except ImportError:
        return -1
    document = docx.Document(str(target))
    removed = 0
    for paragraph in list(document.paragraphs):
        if paragraph.text.strip().startswith("（数据来源："):
            element = paragraph._element
            element.getparent().remove(element)
            removed += 1
    if removed:
        document.save(str(target))
        normalize_zip(target)
    return removed


def _docx_text(path):
    """轻量提取 docx 正文（不依赖 python-docx）：段落标记断行后去标签。"""
    import re as _re
    import zipfile as _zip
    try:
        with _zip.ZipFile(path) as archive:
            xml = archive.read("word/document.xml").decode("utf-8", "replace")
    except (OSError, KeyError, _zip.BadZipFile):
        return ""
    xml = xml.replace("</w:p>", "\n")
    return _re.sub(r"<[^>]+>", "", xml)


OWN_MARKERS = ("示例街道", "示例社区A")
# 参考件与本项目同属一个街道/区，出现共享地名属正常：既不能作为「这是本项目」的证据，
# 也不能进指纹（否则我方正文合法出现「示例路」会被误判串入）。
SHARED_MARKERS = ("示例路", "某区", "某市", "示例社区C")
# 同类报告成品常见的结构词：命中 ≥3 且不含本项目专名，才判为「形式参考」
STRUCT_MARKERS = ("投资决策建议书", "项目建议书", "可行性研究报告", "建设必要性",
                  "投资估算", "财务评价", "资金筹措", "运营模式")
GENERIC_NAMES = ("本项目", "该项目", "老旧小区", "老旧街区", "该片区", "本项目片区", "居住小区",
                 "城市社区", "示范社区", "完整社区", "小区", "社区", "片区",
                 "所在社区", "依托街道", "所在街道", "两个街道", "各街道", "相关街道")
# 以动词/介词开头的「××街道」不是专名，是句子片段（依托街道、施工前由街道…）
PROSE_PREFIXES = ("依托", "分散", "施工", "所在", "配合", "协调", "会同", "组织", "涉及",
                  "位于", "属于", "纳入", "交由", "由", "在", "与", "和", "及", "对", "为",
                  "是", "将", "按", "经", "向", "从", "把", "被", "让", "使", "各", "两", "该")


def register_form_reference(workspace, index, project_name, no_timestamp=False):
    """从投递区识别「明显不是本项目、但是同类报告成品」 → 隔离为形式参考（C7）。

    识别不看固定文件名：结构词命中 ≥3 且本项目专名 0 命中才判定；
    docx/pdf 都在候选内（pdf 不做文本抽取，只在登记里标「需人工核对」）。
    只登记骨架与允许学习的项；同时产出专名/数值指纹，供 08 做零串入守卫。
    """
    import re as _re
    candidates = []
    for item in index.files:
        if item.suffix.lower() not in (".docx", ".pdf"):
            continue
        if any(key in item.name for key in ("现状分析", "财务参数")):
            continue
        text = _docx_text(item) if item.suffix.lower() == ".docx" else ""
        if not text:
            continue
        own_hits = sum(text.count(marker) for marker in OWN_MARKERS)
        struct_hits = sum(1 for marker in STRUCT_MARKERS if marker in text)
        if own_hits or struct_hits < 3:
            continue
        headings = [line.strip() for line in text.splitlines()
                    if _re.match(r"^第[一二三四五六七八九十]+[章节]|^\d{1,2}(\.\d{1,2})?\s+\S{2,20}$",
                                 line.strip())]
        candidates.append({
            "file": rel(item, workspace),
            "type": "形式参考",
            "verdict": {"结构词命中": struct_hits, "本项目专名命中": own_hits},
            "skeleton": headings[:60],
            "allow": ["章节骨架", "句式", "表格样式", "依据排列方式"],
            "deny": ["其他项目名", "地名", "街道社区名", "任何数值"],
            "text": text,
        })
    if not candidates:
        return

    def _numbers(text):
        found = set()
        for match in _re.finditer(r"\d{1,3}(?:,\d{3})+\.\d{2}|\d{4,}\.\d{2}", text):
            found.add(match.group(0))
        return sorted(found)

    def _proper_names(text):
        """只取「××街道 / ××社区」这类真专名＋参考件项目名。

        不用「小区/院落/广场/路/街/巷」做后缀：参考件里大量是「城镇老旧小区」「临街院落」
        这类普通名词，扫进指纹会把本方正文的正常表述误判成串入（实测 271 个里大半是碎片）。
        """
        found = set()
        # 按标点/书名号切段后整段匹配，避免「二是配合街道」「为对示例街道」这类半句碎片
        for segment in _re.split(r"[，。、；：（）()《》〈〉\s]", text):
            segment = segment.strip("　 “”\"'")
            if not _re.match(r"^[\u4e00-\u9fa5]{2,5}(?:街道|社区)$", segment):
                continue
            if any(marker in segment for marker in OWN_MARKERS + SHARED_MARKERS):
                continue
            if segment in GENERIC_NAMES:
                continue
            if any(segment.startswith(prefix) for prefix in PROSE_PREFIXES):
                continue
            found.add(segment)
        for match in _re.finditer(
                r"([\u4e00-\u9fa5A-Za-z0-9（）()]{6,40}?项目)[）)]?\s*(?:项目建议书|可行性研究报告)", text):
            title = match.group(1).strip("（()） 　")
            if any(word in title for word in ("印发", "关于", "通知", "发展改革委", "街道", "社区")):
                continue
            if len(title) >= 8 and not any(marker in title for marker in OWN_MARKERS):
                found.add(title)
        return sorted(found)

    registry = {"schema": "cardo-invest/form-reference/v1", "projectName": project_name,
                "items": [], "rule": "只学章节骨架、句式、表格样式、依据排列；其他项目名/地名/街道社区名/数值零进入正文"}
    fingerprints = {"schema": "cardo-invest/form-reference-fingerprint/v1",
                    "source": [item["file"] for item in candidates],
                    "properNouns": [], "numbers": []}
    for item in candidates:
        text = item.pop("text")
        registry["items"].append(item)
        fingerprints["properNouns"] = sorted(set(fingerprints["properNouns"]) | set(_proper_names(text)))
        fingerprints["numbers"] = sorted(set(fingerprints["numbers"]) | set(_numbers(text)))
    write_json(workspace / "过程文件/01-上游取料/形式参考登记.json", registry, no_timestamp)
    write_json(workspace / "过程文件/01-上游取料/形式参考指纹.json", fingerprints, no_timestamp)
    print("[prepare_inputs] 识别形式参考 {} 份（专名 {} 个、数值 {} 个已入指纹，不做正文素材）".format(
        len(candidates), len(fingerprints["properNouns"]), len(fingerprints["numbers"])))


def write_json(path, payload, no_timestamp=False):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    text = json.dumps(payload, ensure_ascii=False, indent=2, sort_keys=True)
    if not no_timestamp:
        text += "\n"
    path.write_text(text, encoding="utf-8")


def rel(path, base):
    try:
        return Path(path).resolve().relative_to(Path(base).resolve()).as_posix()
    except ValueError:
        return Path(path).name


def pick_rank(path):
    """取件优先级：交付口径优先于过程稿，空白模板排在最后。

    整包倒入上游成果目录时，同名件会同时出现「成品版」与「溯源版」；
    早先按 (名称长度, 路径) 排序，`内部件` 的字母序在 `成品版` 之前，
    结果取到溯源版。这里显式定档：
      0 成品版 → 1 依据数据 → 2 内部件 → 3 其他散件 → 4 溯源/标注版 → 9 空白/模板/样例
    同档内再按名称长度、路径排序，保证结果稳定可复现。
    """
    text = "/".join(Path(path).parts).lower()
    if any(token in text for token in ("空白", "模板", "样例")):
        return 9
    if "成品版" in text:
        return 0
    if any(token in text for token in ("溯源", "标注版")):
        return 4
    if "依据数据" in text:
        return 1
    if "内部件" in text:
        return 2
    return 3


class InputIndex:
    """投递区索引：递归扫描文件/目录，自动解包 zip，按文件名关键词归类。

    用户可以直接放文件、文件夹或压缩包，不要求按目录分类；这里只做发现，不改原件。
    """

    ZIP_SUFFIXES = (".zip",)

    def __init__(self, root, unpack_root):
        self.root = Path(root)
        self.unpack_root = Path(unpack_root)
        self.files = []
        self.archives = []
        self.notes = []

    def build(self):
        self.unpack_root.mkdir(parents=True, exist_ok=True)
        for path in sorted(self.root.rglob("*")) if self.root.is_dir() else []:
            if path.is_file() and path.suffix.lower() in self.ZIP_SUFFIXES:
                self.archives.append(path)
                target = self.unpack_root / path.stem
                try:
                    self._unpack(path, target)
                except Exception as exc:  # 解包失败只登记，不中断装配
                    self.notes.append("解包失败：{}（{}）".format(path.name, exc))
                    continue
        for base in (self.root, self.unpack_root):
            if not base.is_dir():
                continue
            for path in sorted(base.rglob("*")):
                if path.is_file() and path.suffix.lower() not in self.ZIP_SUFFIXES:
                    self.files.append(path)
        return self

    def _unpack(self, path, target):
        import zipfile

        if target.exists():
            shutil.rmtree(target)
        target.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(path) as archive:
            for info in archive.infolist():
                if info.is_dir():
                    continue
                name = info.filename.replace("\\", "/")
                if name.startswith("__MACOSX/") or "/__MACOSX/" in name:
                    continue
                dest = (target / name).resolve()
                if target.resolve() not in dest.parents:
                    continue
                dest.parent.mkdir(parents=True, exist_ok=True)
                with archive.open(info) as src, open(dest, "wb") as out:
                    shutil.copyfileobj(src, out)
        self.notes.append("已解包：{} → {}".format(path.name, target.name))

    def find(self, keywords, exts, required_all=False):
        hits, exts = [], tuple(item.lower() for item in exts)
        for path in self.files:
            if path.suffix.lower() not in exts:
                continue
            name = path.name.lower()
            if required_all:
                matched = all(item.lower() in name for item in keywords)
            else:
                matched = any(item.lower() in name for item in keywords)
            if matched:
                hits.append(path)
        hits.sort(key=lambda item: (pick_rank(item), len(item.name), str(item)))
        return hits

    def first(self, keywords, exts, required_all=False):
        hits = self.find(keywords, exts, required_all)
        return hits[0] if hits else None


# 非标准写法的等价符号（用户手填表里常见）
def _record_fix(label, original, fixed, note):
    """登记一条纠错（让用户看得见「改了什么、为什么改」）。"""
    if len(_CORRECTIONS) < 200:
        _CORRECTIONS.append({
            "field": str(label or ""),
            "original": str(original),
            "fixed": str(fixed),
            "note": note,
        })


_NUM_NORMALIZE = {
    "，": ",", "．": ".", "。": ".", "－": "-", "＋": "+", "％": "%",
    "￥": "", "¥": "", "　": " ",
    "０": "0", "１": "1", "２": "2", "３": "3", "４": "4",
    "５": "5", "６": "6", "７": "7", "８": "8", "９": "9",
}
# 可跳过的单位后缀（值仍按原单位解释）
_NUM_UNIT_SUFFIX = ("万元", "亿元", "元", "个", "套", "盏", "处", "台", "座", "间",
                     "年", "月", "天", "小时", "人", "㎡", "m2", "m²", "km", "吨")


_CORRECTIONS = []       # 本次解析的纠错记录（用户可见：改了什么、为什么改）

# 收入表单位疑误阈值：单一收入项年均收入 ≥ 10,000（万元）＝ 1 亿元/年，
# 按「元填成万元」还原（÷10000）。只换单位，不改业务口径。
REVENUE_UNIT_ANOMALY = 10000


def num(value, label=""):
    """Excel 单元格 → 数值（Decimal 解析，整数落 int，其余落 float）；空 → None。

    容错（用户手填表常见不规范写法，**不能静默丢数据**）：
      * 全角数字/符号 → 半角（「５３００」→「5300」）
      * 千分位（「4,301.76」）
      * 带单位（「4301.76万元」「5300 元」）
      * 带货币符号（「￥5300」）
      * 百分号（「8%」→ 0.08；仅在 label 含「率」或「%」时转换）
      * 前后空白与中文括号
    label 传入字段名时，按字段语义判断是否为百分比字段
    （「预备费率」「年利率」写 8 视为 8%；「回收期」不转换）。
    """
    if value is None or (isinstance(value, str) and not value.strip()):
        return None
    if isinstance(value, (int, float)):
        value = repr(value) if isinstance(value, float) else str(value)
    text = str(value).strip()
    for src, dst in _NUM_NORMALIZE.items():
        text = text.replace(src, dst)
    text = text.replace(",", "").replace(" ", "")
    percent = False
    if "%" in text:
        percent = True
        text = text.replace("%", "")
    for suffix in _NUM_UNIT_SUFFIX:
        if text.endswith(suffix):
            text = text[: -len(suffix)]
            break
    if not text:
        return None
    try:
        dec = Decimal(text)
    except InvalidOperation:
        return None
    # 百分比归一：
    #   * 带 % → 一定 ÷100（「8%」→ 0.08）
    #   * 字段名是百分比字段且值 > 1 → ÷100（「预备费率」写 8 → 0.08）
    #   * 「回收期」等非百分比字段不转换
    if percent:
        dec = dec / Decimal(100)
        _record_fix(label, value, dec, "百分号写法：{} → {}".format(value, dec))
    elif label:
        key = str(label)
        if ("率" in key or "%" in key) and "回收期" not in key and dec > 1:
            dec = dec / Decimal(100)
            _record_fix(label, value, dec, "按百分数字段解读：{} → {}".format(value, dec))
    if "e" in text.lower():
        return float(dec)
    return int(dec) if dec == dec.to_integral_value() else float(dec)


def text(value):
    return "" if value is None else str(value).strip()


# ---------------------------------------------------------------- xlsx 读取

def load_sheets(path):
    import openpyxl

    workbook = openpyxl.load_workbook(path, data_only=True, read_only=True)
    sheets = {}
    for sheet in workbook.worksheets:
        rows = []
        for row in sheet.iter_rows(values_only=True):
            if row is None:
                continue
            cells = list(row)
            if not any(text(cell) for cell in cells):
                continue
            rows.append(cells)
        sheets[sheet.title] = rows
    workbook.close()
    return sheets


def find_sheet(sheets, *keywords):
    for name, rows in sheets.items():
        for keyword in keywords:
            if keyword in name:
                return rows
    return None


def table_rows(rows, required):
    """取「表头 + 数据行」：定位首行为表头（含 required 里任一列名），返回 (列索引, 数据行)。"""
    if not rows:
        return {}, []
    header_index = 0
    for index, row in enumerate(rows):
        labels = [text(cell) for cell in row]
        if any(label in required for label in labels):
            header_index = index
            break
    header = [text(cell) for cell in rows[header_index]]
    columns = {label: position for position, label in enumerate(header) if label}
    return columns, rows[header_index + 1:]


def cell(row, columns, *names):
    for name in names:
        position = columns.get(name)
        if position is not None and position < len(row):
            value = row[position]
            if value is not None:
                return value
    return None


# ---------------------------------------------------------------- 财务参数.xlsx → cardo/invest-params/v1

def parse_params(path):
    del _CORRECTIONS[:]
    sheets = load_sheets(path)
    kv = {}
    for name, rows in sheets.items():
        for row in rows:
            if len(row) >= 2 and text(row[0]) and num(row[1]) is None:
                kv.setdefault(text(row[0]), text(row[1]))

    project_rows = find_sheet(sheets, "项目") or []
    project = {}
    for row in project_rows:
        if len(row) >= 2:
            project[text(row[0])] = row[1]

    engineering = []
    columns, data = table_rows(find_sheet(sheets, "工程费") or [], {"编号", "名称", "金额"})
    for row in data:
        name = text(cell(row, columns, "名称", "工程名称"))
        amount = num(cell(row, columns, "金额(万元)", "金额", "投资(万元)"))
        if not name or amount is None:
            continue
        item = {"key": text(cell(row, columns, "编号", "代码")) or name,
                "name": name, "amount": amount}
        unit = text(cell(row, columns, "单位"))
        quantity = num(cell(row, columns, "工程量", "数量"))
        if unit:
            item["unit"] = unit
        if quantity is not None:
            item["quantity"] = quantity
        engineering.append(item)

    other_fees = []
    columns, data = table_rows(find_sheet(sheets, "其他费") or [], {"编号", "名称", "金额", "取费基数"})
    for row in data:
        name = text(cell(row, columns, "名称", "费用名称"))
        amount = num(cell(row, columns, "金额(万元)", "金额"))
        basis = num(cell(row, columns, "取费基数(万元)", "取费基数", "基数(万元)"))
        rate = num(cell(row, columns, "费率", "费率(%)"), "费率")
        if not name or (amount is None and (basis is None or rate is None)):
            continue
        item = {"key": text(cell(row, columns, "编号", "代码")) or name, "name": name}
        if amount is not None:
            item["amount"] = amount
        else:
            item["basis"] = basis
            item["rate"] = rate
        other_fees.append(item)

    contingency_rows = find_sheet(sheets, "预备费") or []
    contingency = None
    for row in contingency_rows:
        if len(row) >= 2 and "费率" in text(row[0]):
            contingency = num(row[1], text(row[0]))
    if contingency is None:
        for row in contingency_rows:
            for value in row[1:]:
                if num(value, text(row[0])) is not None:
                    contingency = num(value, text(row[0]))
                    break

    interest_rows = find_sheet(sheets, "建设期利息") or []
    interest = {}
    for row in interest_rows:
        if len(row) < 2:
            continue
        label = text(row[0])
        if "融资费用" in label or "利息额" in label:
            interest["amount"] = num(row[1])
        elif "借款" in label:
            interest["loanAmount"] = num(row[1])
        elif "利率" in label:
            interest["loanRate"] = num(row[1], "年利率")
        elif "年数" in label or "建设期" in label:
            interest["buildYears"] = num(row[1])
        elif "半额" in label:
            interest["halfYear"] = text(row[1]) in ("是", "true", "TRUE", "1", "Y", "y")

    revenue = []
    columns, data = table_rows(find_sheet(sheets, "收入") or [], {"编号", "名称", "单价"})
    for row in data:
        name = text(cell(row, columns, "名称", "收入项"))
        annual = num(cell(row, columns, "年均收入(万元)", "年均收入"))
        unit_suspect = None
        # 单位疑误：表头口径是「万元」，个别行按「元」误填会使数值放大约 1 万倍。
        # 判据只看量级——单一收入项年均收入 ≥ 10,000（即万元口径下已超 1 亿元/年），
        # 在街区更新这类小体量项目里不成立；按「元」还原为万元，并登记纠错台账，
        # 原表不改、正文可追溯。
        if annual is not None and annual >= REVENUE_UNIT_ANOMALY:
            fixed = round(annual / 10000, 2)
            _record_fix("收入/{}".format(name or "未命名"), annual, fixed,
                        "《收入》表按「万元」表头填写，该行数值量级为元（{:.2f} 万元/年 ⇒ {:.2f} 亿元/年，"
                        "量级不成立），按元÷10000 还原为万元；原表未改".format(annual, annual / 10000))
            unit_suspect = {"declaredUnit": "万元", "originalAmount": annual,
                            "restoredAmount": fixed, "restoredCaliber": "元计列为万元"}
            annual = fixed
        price = num(cell(row, columns, "单价", "单价(元)"))
        quantity = num(cell(row, columns, "数量"))
        if not name or (annual is None and (price is None or quantity is None)):
            continue
        item = {"key": text(cell(row, columns, "编号", "代码")) or name, "name": name}
        if unit_suspect:
            # 单位疑误要随收入项带到 05/08：正文脚注与内部件《待确认清单》都按此落点，
            # 不允许只在纠错台账里留一句、成品里静默改数。
            item["unitSuspect"] = unit_suspect
        if annual is not None:
            item["annualAmount"] = annual
        else:
            item["price"] = price
            item["quantity"] = quantity
            days = num(cell(row, columns, "计费天数", "天数"))
            months = num(cell(row, columns, "计费月数", "月数"))
            if days is not None:
                item["days"] = days
            if months is not None:
                item["months"] = months
        growth = num(cell(row, columns, "年增长率", "增长率"), "年增长率")
        vat_rate = num(cell(row, columns, "增值税率", "税率"), "增值税率")
        if growth is not None:
            item["growth"] = growth
        if vat_rate is not None:
            item["vatRate"] = vat_rate
        # 量价/出租率口径是正文「测算依据」列的唯一来源，不清洗、不改写。
        basis_text = text(cell(row, columns, "测算依据", "量价依据"))
        if basis_text:
            item["basisText"] = basis_text
        revenue.append(item)

    cost_rows = find_sheet(sheets, "成本") or []
    labor = {}
    for row in cost_rows:
        if len(row) < 2:
            continue
        label = text(row[0])
        if "人数" in label or "班组" in label:
            labor["crew"] = num(row[1])
        elif "薪酬" in label:
            labor["salaryWan"] = num(row[1])
        elif "福利" in label:
            labor["welfareFactor"] = num(row[1], "福利系数")
        elif "增长" in label:
            labor["growthRate"] = num(row[1], "年增长率")
    cost = {"labor": labor} if labor else {}

    finance_rows = find_sheet(sheets, "资金结构") or []
    finance = {}
    for row in finance_rows:
        if len(row) < 2:
            continue
        label = text(row[0])
        value = row[1]
        if "申报上限" in label:
            # 申报口径只用于披露「申报额度 − 投资需求」差额，不进入现金流。
            finance["declaredLoan"] = num(value)
        elif "申报口径" in label:
            finance["declaredFunding"] = num(value)
        elif "资本金" in label:
            finance["equity"] = num(value)
        elif "补助" in label:
            finance["subsidy"] = num(value)
        elif "借款期限" in label:
            finance["loanTerm"] = num(value)
        elif "宽限" in label:
            finance["graceYears"] = num(value)
        elif "还款方式" in label:
            finance["repaymentMethod"] = text(value)
        elif "利率" in label:
            finance["loanRate"] = num(value, "借款年利率")
        elif "评价投资基数" in label:
            finance["evaluationInvestment"] = num(value)
        elif "借款" in label:
            finance["loan"] = num(value)

    reference = {}
    reference_rows = find_sheet(sheets, "可研财务指标") or []
    for row in reference_rows:
        if len(row) < 2:
            continue
        label, value = text(row[0]), num(row[1])
        mapping = {
            "全部投资FIRR": "firrAll", "全部投资FNPV": "fnpvAll", "全部投资回收期": "ptAll",
            "资本金FIRR": "firrEquity", "资本金FNPV": "fnpvEquity", "资本金回收期": "ptEquity",
            "FNPV折现率": "discountRate",
        }
        for key, target in mapping.items():
            if label.startswith(key):
                # 只收有值的项：空单元格代表「未提供可研对照」，不留 None 占位
                # （否则下游拿到的是「键存在但值为 null」的污染状态，
                #   既分不清"没填可研指标"与"填了但算不出"，也会让口径判定走偏）
                if value is not None:
                    reference[target] = value
                break
    if reference:
        finance["referenceIndicators"] = reference

    # 项目字段：**取不到就不写这个键**（不是写 None）。
    # 原因（实测）：下游 03 财务评价 `int(payload["project"]["constructionYears"])`
    # 在值为 None 时直接抛 TypeError，整条链在 03 崩掉。
    # 键缺失 → 下游用默认值 → 链路不中断；缺项由缺项表登记（不补数）。
    _project = {
        "name": text(project.get("项目名称")) or text(kv.get("项目名称")),
        "direction": text(project.get("项目方向")) or text(kv.get("项目方向")),
        "area": num(project.get("改造面积(㎡)")),
        "constructionYears": num(project.get("建设期(年)")),
        "operationYears": num(project.get("运营期(年)")),
        "discountRate": num(project.get("折现率"), "折现率"),
        "landTaxArea": num(project.get("土地使用税计税面积(㎡)")),
    }
    _project = dict((key, value) for key, value in _project.items()
                     if value is not None and value != "")
    payload = {
        "schema": "cardo/invest-params/v1",
        "project": _project,
        "estimate": {
            "engineering": engineering,
            "otherFees": other_fees,
            "otherFeeTotal": None,   # 契约形状：未给合计由下游汇总，不补数
            "contingency": {"rate": contingency},
            "interest": interest,
        },
        "revenue": revenue,
        "cost": cost,
        "tax": {},
        "finance": finance,
        "profit": {},
        "schemes": {},
    }
    # 纠错台账：用户填的非标准写法被归一后，把「原值 → 归一值」记下来，
    # 让用户看得见改了什么（而不是静默改掉）。
    if _CORRECTIONS:
        payload["corrections"] = list(_CORRECTIONS)
    return payload


# ---------------------------------------------------------------- 槽位.xlsx → cardo/invest-slots/v1

SLOT_KEYS = {
    "功能定位": "functionalPosition",
    "业态": "businessTypes",
    "实施时序": "schedule",
    "组织架构": "organization",
    "运营模式": "operationModel",
    "改造思路": "renovationIdea",
}


def split_list(value):
    parts = []
    for chunk in str(value).replace("，", "、").replace("；", "；").replace(";", "；").split("、"):
        chunk = chunk.strip()
        if chunk:
            parts.append(chunk)
    for chunk in list(parts):
        if "；" in chunk:
            parts.remove(chunk)
            parts.extend([piece.strip() for piece in chunk.split("；") if piece.strip()])
    return parts


def parse_slots(path):
    sheets = load_sheets(path)
    rows = find_sheet(sheets, "槽位") or []
    slots = {}
    for row in rows:
        if len(row) < 2:
            continue
        key = text(row[0])
        value = text(row[1])
        if key not in SLOT_KEYS or not value:
            continue
        slots[SLOT_KEYS[key]] = split_list(value) if SLOT_KEYS[key] in ("businessTypes", "schedule") else value
    return slots



def parse_ahp_scores_office(path):
    """读《AHP方案评分结论.xlsx》：排序与评级 + 指标得分明细。只抄数，不合成、不重算。"""
    sheets = load_sheets(path)
    ranking_rows = find_sheet(sheets, "排序与评级", "排序")
    detail_rows = find_sheet(sheets, "指标得分明细", "指标得分")
    if not ranking_rows or not detail_rows:
        return None

    columns, data = table_rows(ranking_rows, {"名次", "方案", "综合得分"})
    ranking, summary = [], None
    for row in data:
        name = text(cell(row, columns, "方案"))
        composite = num(cell(row, columns, "综合得分"))
        if not name:
            continue
        if composite is None:                       # 「结论」行：整段摘要
            summary = text(cell(row, columns, "结论")) or next(
                (text(value) for value in row[1:] if text(value)), None)
            continue
        ranking.append({
            "name": name,
            "composite": composite,
            "label": text(cell(row, columns, "评级")),
            "action": text(cell(row, columns, "投资建议")),
            "coverage": text(cell(row, columns, "权重覆盖率")),
            "note": text(cell(row, columns, "备注")),
        })
    if len(ranking) < 3:
        return None

    scheme_names = [item["name"] for item in ranking]
    columns2, data2 = table_rows(detail_rows, {"准则", "指标", "指标ID", "综合权重"})
    detail = []
    for row in data2:
        indicator_id = text(cell(row, columns2, "指标ID"))
        indicator = text(cell(row, columns2, "指标"))
        if not indicator_id or not indicator:
            continue
        detail.append({
            "id": indicator_id,
            "name": indicator,
            "criterion": text(cell(row, columns2, "准则")),
            "globalWeight": num(cell(row, columns2, "综合权重")),
            "scores": {name: num(cell(row, columns2, name + "得分")) for name in scheme_names},
            "contributions": {name: num(cell(row, columns2, name + "贡献")) for name in scheme_names},
        })
    if not detail:
        return None

    by_name = {item["name"]: item for item in ranking}
    # 方案排序：先按 SCHEME_ORDER 归位（上游方案名与本包标准三方案的对应），
    # 排不动的按「排名表里的名次」保序 —— 排序只决定 S1/S2/S3 的编号先后，
    # 不改动任何一个方案的名称与得分。
    name_rank = {item["name"]: index for index, item in enumerate(ranking)}
    ordered = sorted(scheme_names,
                      key=lambda name: (SCHEME_ORDER.get(name, 99), name_rank.get(name, 99)))
    schemes = []
    for index, name in enumerate(ordered, start=1):
        item = by_name[name]
        schemes.append({
            "id": "S{}".format(index),
            "name": name,
            "composite": item["composite"],
            "rating": {"label": item["label"], "action": item["action"]},
            "scores": [{
                "id": row["id"],
                "name": row["name"],
                "criterion": row["criterion"],
                "criterionId": row["id"].split(".")[0],
                "globalWeight": row["globalWeight"],
                "score": row["scores"].get(name),
            } for row in detail],
        })

    top_name = max(ranking, key=lambda item: item["composite"])["name"]
    top_drivers = sorted(
        [{"id": row["id"], "name": row["name"], "score": row["scores"].get(top_name),
          "contribution": row["contributions"].get(top_name)}
         for row in detail if row["contributions"].get(top_name) is not None],
        key=lambda item: item["contribution"], reverse=True)[:3]
    return {"schemes": schemes, "ranking": ranking, "summary": summary, "topDrivers": top_drivers}


def parse_consistency_from_conclusion(path):
    """从《AHP权重与评分结论.docx》的一致性表读 CR/λmax（只读上传结论，不复算）。"""
    if not path or not Path(path).is_file():
        return None
    try:
        import docx
    except ImportError:
        return None
    try:
        document = docx.Document(str(path))
    except Exception:
        return None
    for table in document.tables:
        header = [item.text.strip() for item in table.rows[0].cells]
        if "CR" not in header or "λmax" not in header:
            continue
        position = {name: index for index, name in enumerate(header)}
        for row in table.rows[1:]:
            cells = [item.text.strip() for item in row.cells]
            if not cells or cells[0] != "criteria":
                continue
            return {
                "cr": num(cells[position["CR"]]),
                "n": num(cells[position["n"]]) if "n" in position else None,
                "lambdaMax": num(cells[position["λmax"]]),
                "ri": num(cells[position["RI"]]) if "RI" in position else None,
                "pass": cells[position["判定"]] == "通过" if "判定" in position else None,
                "method": "upstream-office",
                "note": "取自上游《AHP权重与评分结论.docx》一致性表；本工作区不复算",
            }
    return None


def parse_ahp_weights_finished(weights_path, scores_path=None):
    """读《AHP权重表_成品版.xlsx》的「权重表」页：准则/指标权重 + 一致性检验。

    成品版与溯源版的表结构不同（成品版面向交付，无「导出值」机读页）：
      「权重表」页三段：一、准则层权重；二、指标层权重（局部×准则=综合）；
      三、一致性检验（判断矩阵/阶数/λmax/CI/RI/CR/判定）。
    指标ID 从成品版《AHP方案评分结论》的「指标得分明细」按指标名对齐。
    """
    sheets = load_sheets(weights_path)
    rows = find_sheet(sheets, "权重表")
    if not rows:
        return None

    scores = parse_ahp_scores_office(scores_path) if scores_path and Path(scores_path).is_file() else None
    id_of_name = {}
    if scores:
        for item in scores["schemes"][0]["scores"]:
            id_of_name[item["name"]] = item["id"]

    criteria, indicators, consistency = [], [], None
    section = None
    for row in rows:
        labels = [text(value) for value in row]
        title = "".join(labels)
        if "准则层权重" in title:
            section = "criteria"
            continue
        if "指标层权重" in title:
            section = "indicators"
            continue
        if "一致性检验" in title:
            section = "consistency"
            continue
        if section == "criteria":
            name, weight = text(row[0]) if row else "", num(row[1]) if len(row) > 1 else None
            if not name or name == "准则" or weight is None:
                continue
            criteria.append({"key": name, "id": None, "name": name, "weight": weight})
        elif section == "indicators":
            criterion, name = text(row[0]) if row else "", text(row[1]) if len(row) > 1 else ""
            local = num(row[2]) if len(row) > 2 else None
            composite = num(row[3]) if len(row) > 3 else None
            if not criterion or not name or criterion == "准则" or local is None:
                continue
            indicator_id = id_of_name.get(name) or name
            indicators.append({
                "criterion": criterion,
                "criterionWeight": next((item["weight"] for item in criteria if item["name"] == criterion), None),
                "indicator": name,
                "indicatorId": indicator_id,
                "subWeight": local,
                "compositeWeight": composite,
            })
        elif section == "consistency":
            name = text(row[0]) if row else ""
            if not name or name == "判断矩阵" or len(row) < 7:
                continue
            cr = num(row[5])
            if cr is None:
                continue
            candidate = {
                "cr": cr,
                "n": num(row[1]),
                "lambdaMax": num(row[2]),
                "ri": num(row[4]),
                "pass": text(row[6]) == "通过",
                "method": "upstream-office",
                "note": "取自上游《AHP权重表（成品版）》一致性检验表；本工作区不复算",
            }
            # 成品版按判断矩阵逐行列示，没有「准则层」汇总行：取 CR 最大的一个作为整体口径。
            if consistency is None or cr > consistency["cr"]:
                consistency = candidate
            elif not candidate["pass"]:
                consistency["pass"] = False
    if not criteria or not indicators:
        return None
    return {
        "schema": "cardo/ahp-weights/v1",
        "source": [{"kind": "upstream-office", "id": "ahp-weights-xlsx",
                    "path": Path(weights_path).name}],
        "criteria": criteria,
        "criteriaWeights": [{"key": item["key"], "name": item["name"], "weight": item["weight"],
                             "geometricMean": None} for item in criteria],
        "indicatorWeights": indicators,
        "consistency": consistency or {
            "cr": None, "method": "upstream-office",
            "note": "上游已完成一致性检验（详见《AHP权重与评分结论》），本工作区不复算",
        },
        "temporary": False,
        "provenance": "准则权重与指标权重取自《AHP权重表（成品版）》「权重表」页；"
                      "判断矩阵表由权重派生展示（上游一致性已通过）",
        "note": "AHP 成果由用户在 AHP 工作区运行后拷入 输入文件/；本工作区只读，不重算。",
    }


def parse_ahp_weights_office(weights_path, scores_path=None, conclusion_path=None):
    """读《AHP权重表.xlsx》导出值（准则/局部/综合权重），指标名与准则名从评分结论明细对齐。"""
    sheets = load_sheets(weights_path)
    if not find_sheet(sheets, "导出值"):
        # 成品版没有「导出值」页，改读「权重表」页（交付口径优先）。
        return parse_ahp_weights_finished(weights_path, scores_path)
    rows = find_sheet(sheets, "导出值") or []
    values = {}
    for row in rows:
        if len(row) < 2:
            continue
        key, value = text(row[0]), num(row[1])
        if key and value is not None:
            values[key] = value
    criterion_weights = {key.split(".", 1)[1]: value for key, value in values.items()
                         if key.startswith("准则权重.")}
    local_weights = {key.split(".", 1)[1]: value for key, value in values.items()
                     if key.startswith("指标局部.")}
    global_weights = {key.split(".", 1)[1]: value for key, value in values.items()
                      if key.startswith("指标综合.")}
    if not criterion_weights or not local_weights:
        return None

    scores = parse_ahp_scores_office(scores_path) if scores_path and Path(scores_path).is_file() else None
    detail_names, criterion_names, criterion_of = {}, {}, {}
    if scores:
        for item in scores["schemes"][0]["scores"]:
            prefix = item["id"].split(".")[0]
            detail_names[item["id"]] = item["name"]
            criterion_names[prefix] = item.get("criterion") or prefix
            criterion_of[prefix] = item.get("criterion") or prefix

    criteria = []
    for criterion_id, weight in criterion_weights.items():
        name = criterion_names.get(criterion_id, criterion_id)
        criteria.append({"key": name, "id": criterion_id, "name": name, "weight": weight})

    indicator_weights = []
    for indicator_id, sub_weight in local_weights.items():
        criterion_id = indicator_id.split(".")[0]
        indicator_weights.append({
            "criterion": criterion_of.get(criterion_id, criterion_id),
            "criterionWeight": criterion_weights.get(criterion_id),
            "indicator": detail_names.get(indicator_id, indicator_id),
            "indicatorId": indicator_id,
            "subWeight": sub_weight,
            "compositeWeight": global_weights.get(indicator_id),
        })

    consistency = parse_consistency_from_conclusion(conclusion_path) or {
        "cr": None, "method": "upstream-office",
        "note": "上游已完成一致性检验（详见《AHP权重与评分结论.docx》），本工作区不复算",
    }
    return {
        "schema": "cardo/ahp-weights/v1",
        "source": [{"kind": "upstream-office", "id": "ahp-weights-xlsx",
                    "path": rel(weights_path, Path(weights_path).parent.parent.parent)}],
        "criteria": criteria,
        "criteriaWeights": [{"key": item["key"], "name": item["name"], "weight": item["weight"],
                             "geometricMean": None} for item in criteria],
        "indicatorWeights": indicator_weights,
        "consistency": consistency,
        "temporary": False,
        "provenance": "准则权重与指标权重取自用户拷入的 AHP 成果导出值；判断矩阵未随导出值提供，"
                      "表 3 为按权重派生的展示矩阵（上游一致性已通过，详见结论附件）",
        "note": "AHP 成果由用户在 AHP 工作区运行后拷入 输入文件/（任意层级/压缩包均可）；本工作区只读，不重算。",
    }


# ---------------------------------------------------------------- AHP 权重/评分（上游机读 JSON → 契约键）

def normalize_weights(raw):
    """上游权重 → 引擎契约：只校验与改键，不动数值，不重算。

    - key 用准则中文名（引擎的默认指标映射与方案评分都按中文准则名取）；
    - criteriaMatrix / criteriaWeights / indicatorWeights / consistency 原样透传；
    - 上游只给 criteria（无指标层）时透传其余缺省，由取料脚本按默认子权重派生并标注。
    """
    if not isinstance(raw, dict) or not isinstance(raw.get("criteria"), list):
        return None
    criteria = []
    for item in raw.get("criteria") or []:
        try:
            weight = float(Decimal(str(item.get("weight"))))
        except (InvalidOperation, TypeError, ValueError):
            return None
        name = text(item.get("name")) or text(item.get("key")) or text(item.get("id"))
        criteria.append({
            "key": text(item.get("key")) or name,
            "id": text(item.get("id")) or None,
            "name": name,
            "weight": weight,
        })
    if not criteria or sum(item["weight"] for item in criteria) <= 0:
        return None
    kept = {
        "schema": "cardo/ahp-weights/v1",
        "criteria": criteria,
        "temporary": bool(raw.get("temporary")),
        "source": raw.get("source") if isinstance(raw.get("source"), list) else [],
    }
    for key in ("criteriaMatrix", "criteriaWeights", "indicatorWeights", "consistency", "note"):
        if raw.get(key) is not None:
            kept[key] = raw[key]
    return kept


# ---------------------------------------------------------------- 主流程

def main(argv=None):
    parser = argparse.ArgumentParser(description="cardo-invest 上游取料（Office/机读 JSON → 引擎项目）")
    parser.add_argument("--workspace", default=".", help="工作区根（含 输入文件/ 过程文件/）")
    parser.add_argument("--scenario", choices=("full", "gaps"), default="full",
                        help="full=按输入装配；gaps=刻意不上游（验证缺项分支）")
    parser.add_argument("--no-timestamp", action="store_true")
    args = parser.parse_args(argv)

    workspace = Path(args.workspace).resolve()
    params_file = workspace / "工作区参数.json"
    config = read_json(params_file) or {}
    if not (workspace / "输入文件").is_dir() or not (workspace / "过程文件").is_dir():
        print("[prepare_inputs] 不是有效工作区（缺 输入文件/ 或 过程文件/）：{}".format(workspace))
        return 2

    paths = config.get("paths") or {}
    input_root = workspace / paths.get("inputRoot", "输入文件")
    upstream_json = workspace / paths.get("upstreamJson", "过程文件/00-总控/上游机读件")
    params_office = workspace / paths.get("financialParams", "输入文件/财务参数.xlsx")
    staging = workspace / paths.get("stagingProject", "过程文件/00-总控/引擎项目")

    # 投递区发现：递归扫描 + 自动解包；不要求用户按目录分类
    index = InputIndex(input_root, workspace / "过程文件/00-总控/输入解包").build()
    if params_office and not Path(params_office).is_file():
        found_params = index.first(["财务参数"], [".xlsx"])
        if found_params is not None:
            params_office = found_params

    # 清理旧产物：先删后建，避免上次跑剩的中间件混进本次。
    # **但保留「写作任务/分章初稿_待回填」**：那里面是 agent/大模型写好的章节正文，
    # 属于人工/模型投入的成果，删了等于白干一轮（用户要重新成文）。
    filled_backup = None
    filled_dir = staging / "05_过程" / "投资决策" / "写作任务" / "分章初稿_待回填"
    if filled_dir.is_dir():
        filled_backup = Path(tempfile.mkdtemp(prefix="cardo-filled-"))
        try:
            shutil.copytree(filled_dir, str(filled_backup / "分章初稿_待回填"))
        except OSError:
            shutil.rmtree(str(filled_backup), ignore_errors=True)
            filled_backup = None
    if staging.exists():
        shutil.rmtree(staging)
    for sub in ("04_成果报告", "05_过程/AHP", "05_过程/低空解析", "05_过程/投资决策"):
        (staging / sub).mkdir(parents=True, exist_ok=True)
    if filled_backup is not None:
        try:
            shutil.copytree(str(filled_backup / "分章初稿_待回填"), str(filled_dir))
        except OSError:
            pass
        finally:
            shutil.rmtree(str(filled_backup), ignore_errors=True)

    project_name = text(config.get("projectName")) or "某某片区改造项目"
    placed, missing = [], []
    full = args.scenario == "full"

    def note(action, source, target, detail=""):
        placed.append({"action": action, "source": source, "target": target, "detail": detail})

    def skip(code, what, need, suggestion=""):
        missing.append({"code": code, "what": what, "need": need, "suggestion": suggestion})

    # 0) 项目索引
    (staging / "00_项目索引.md").write_text(
        "# {}\n\n> 由 CardoInvest/01-上游取料 自动装配的引擎项目（过程目录，勿手改）。\n".format(project_name),
        encoding="utf-8")

    # 1) 现状分析报告（docx 只搬不改）
    report = None
    if full:
        report = index.first(["现状分析报告"], [".docx"]) or index.first(["现状分析"], [".docx"])
    if report is not None and report.is_file():
        removed = copy_status_report(report, staging / "04_成果报告" / "现状分析报告.docx")
        note("copy", rel(report, workspace), "04_成果报告/现状分析报告.docx",
             "docx 原样搬运；剔除上游『（数据来源：…）』标注行 {} 处".format(removed) if removed > 0
             else "docx 原样搬运（无标注行或未装 python-docx）")
    else:
        skip("UPSTREAM-STATUS", "现状分析报告（docx）", "输入文件/ 放 *现状分析报告*.docx",
             "如果你有 CardoStatus 跑出来的《现状分析报告》，放进 输入文件/；没有就先出一版缺口稿，把缺的地方逐条登记")

    # 2) AHP 权重：用户从 AHP 成果文件拷入的 xlsx 优先；机读 JSON 仅作兼容兜底
    weights = None
    ahp_scores_office = index.first(["AHP方案评分结论"], [".xlsx"]) or index.first(["评分结论"], [".xlsx"])
    ahp_weights_office = index.first(["AHP权重表"], [".xlsx"]) or index.first(["权重表"], [".xlsx"])
    ahp_conclusion_office = (index.first(["AHP权重与评分结论"], [".docx"])
                             or index.first(["评分结论"], [".docx"]))
    if full and ahp_weights_office is not None and ahp_weights_office.is_file():
        weights = parse_ahp_weights_office(ahp_weights_office, ahp_scores_office, ahp_conclusion_office)
    if weights:
        write_json(staging / "05_过程/AHP/权重表.json", weights, args.no_timestamp)
        note("convert", rel(ahp_weights_office, workspace), "05_过程/AHP/权重表.json",
             "读“导出值”表：准则/局部/综合权重；指标名对齐评分结论明细（不重算）")
    elif full and (upstream_json / "权重表.json").is_file():
        weights = normalize_weights(read_json(upstream_json / "权重表.json"))
        if weights:
            write_json(staging / "05_过程/AHP/权重表.json", weights, args.no_timestamp)
            note("convert", rel(upstream_json / "权重表.json", workspace), "05_过程/AHP/权重表.json",
                 "兼容路径：读上游机读 JSON（xlsx 未提供时）")
    else:
        skip("UPSTREAM-AHP", "AHP 权重表（cardo/ahp-weights/v1）", "输入文件/ AHP权重表.xlsx",
             "先在 CardoAHP 工作区跑完再拷三份 Office；暂无 AHP 时用默认权重并标 temporary，不编造专家评分")
        skip("UPSTREAM-AHP-INVALID" if full else "UPSTREAM-AHP", "AHP 权重不可用 → 引擎用默认权重并标临时",
             "上游 AHP 工作区导出 权重表.json")

    # 3) AHP 评分：xlsx 优先解析为契约 JSON；机读 JSON 兜底
    scores_payload = None
    if full and ahp_scores_office is not None and ahp_scores_office.is_file():
        parsed = parse_ahp_scores_office(ahp_scores_office)
        if parsed:
            scores_payload = {
                "schema": "cardo/ahp-scores/v1",
                "source": [{"kind": "upstream-office", "id": "ahp-scores-xlsx",
                            "path": rel(ahp_scores_office, workspace)}],
                "scoring": {"scale": 10, "min": 0, "max": 10},
                "schemes": parsed["schemes"],
                "ranking": parsed["ranking"],
                "conclusion": {"summary": parsed["summary"], "topDrivers": parsed["topDrivers"]},
                "note": "自《AHP方案评分结论.xlsx》抄录：排序/评级/逐指标得分与贡献；本工作区不合成、不重算。",
            }
    if scores_payload:
        write_json(staging / "05_过程/AHP/评分表.json", scores_payload, args.no_timestamp)
        note("convert", rel(ahp_scores_office, workspace), "05_过程/AHP/评分表.json",
             "读“排序与评级/指标得分明细”两表（只抄数）")
    elif full and (upstream_json / "评分表.json").is_file():
        shutil.copy2(upstream_json / "评分表.json", staging / "05_过程/AHP/评分表.json")
        note("copy", rel(upstream_json / "评分表.json", workspace), "05_过程/AHP/评分表.json",
             "兼容路径：上游评分机读 JSON 原样")
    else:
        skip("UPSTREAM-AHP-SCORES", "AHP 方案评分表", "输入文件/ AHP方案评分结论.xlsx",
             "如果你有 CardoAHP 跑出来的权重与评分，放进 输入文件/；没有则第 9 章只列权重与占位，不编造评分")

    # 4) 低空现状指标：只认 json；xlsx 仅人读留档、不解析（不许静默降级）
    metrics_json = index.first(["现状指标"], [".json"])
    metrics_xlsx = index.first(["现状指标"], [".xlsx"])
    if metrics_json is not None and metrics_json.is_file():
        shutil.copy2(metrics_json, staging / "05_过程/低空解析/现状指标.json")
        note("copy", rel(metrics_json, workspace), "05_过程/低空解析/现状指标.json",
             "用户复制进来的现状指标 JSON（实测指标来源）")
    elif full and (upstream_json / "现状指标.json").is_file():
        shutil.copy2(upstream_json / "现状指标.json", staging / "05_过程/低空解析/现状指标.json")
        note("copy", rel(upstream_json / "现状指标.json", workspace), "05_过程/低空解析/现状指标.json",
             "兼容路径：旧交接位 上游机读件/现状指标.json")
    elif metrics_xlsx is not None:
        skip("UPSTREAM-OSGB-XLSX-ONLY",
             "低空现状指标：已检测到《现状指标.xlsx》；该文件不在本包输入范围（本包只解析 JSON），不参与解析",
             "如需入库，向 CardoOSGB 要 过程文件/06-2.5D网格指标/现状指标.json，复制进 输入文件/",
             "xlsx 留作人读存档，不作解析输入")
    else:
        skip("UPSTREAM-OSGB", "低空现状指标（cardo/lowair-indicators/v1）",
             "向 CardoOSGB 要 过程文件/06-2.5D网格指标/现状指标.json，复制进 输入文件/",
             "没有则第 3 章指标表缺失；不要用公开遥感/统计值硬填")

    # 5) 现状事实库：输入文件优先（用户从 Status 复制），旧交接位兼容
    facts_dir = staging / "05_过程/现状事实库"
    facts_dir.mkdir(parents=True, exist_ok=True)
    facts_file = index.first(["事实库"], [".json"])
    status_facts = workspace / "过程文件/05-事实抽取与事实库/facts/事实库.json"
    facts_source = upstream_json / "现状事实库"
    if facts_file is not None and facts_file.is_file():
        shutil.copy2(facts_file, facts_dir / facts_file.name)
        note("copy", rel(facts_file, workspace), "05_过程/现状事实库/",
             "用户复制进来的现状事实库 JSON")
    elif full and status_facts.is_file():
        shutil.copy2(status_facts, facts_dir / status_facts.name)
        note("copy", rel(status_facts, workspace), "05_过程/现状事实库/",
             "同工作区 Status 原生事实库（用户未复制时的兜底）")
    elif full and facts_source.is_dir() and any(facts_source.glob("*.json")):
        for item in sorted(facts_source.glob("*.json")):
            shutil.copy2(item, facts_dir / item.name)
        note("copy", rel(facts_source, workspace), "05_过程/现状事实库/", "事实库 JSON")
    else:
        skip("UPSTREAM-FACTS", "现状事实库（cardo/facts，占位契约）",
             "向 CardoStatus 要 过程文件/05-事实抽取与事实库/facts/事实库.json，复制进 输入文件/",
             "缺则第 4 章问题证据只能按推断处理（标注为推断，需现场核实）")

    # 6) 手动槽位（JSON 优先，xlsx 兜底）
    slots = None
    slots_source = None
    if full:
        if (upstream_json / "槽位.json").is_file():
            raw = read_json(upstream_json / "槽位.json") or {}
            slots = raw.get("slots") if isinstance(raw.get("slots"), dict) else raw
            slots_source = upstream_json / "槽位.json"
        else:
            slots_xlsx = index.first(["槽位"], [".xlsx"])
            if slots_xlsx is not None:
                slots = parse_slots(slots_xlsx)
                slots_source = slots_xlsx
    if slots:
        write_json(staging / "05_过程/投资决策/槽位.json",
                   {"schema": "cardo/invest-slots/v1", "slots": slots}, args.no_timestamp)
        note("convert", rel(slots_source, workspace), "05_过程/投资决策/槽位.json", "手工槽位 {}".format(len(slots)))
    else:
        skip("INPUT-SLOTS", "手动槽位（功能定位/业态/时序/组织/运营）", "输入文件/ 槽位.xlsx",
             "可由项目负责人填写；暂缺时登记缺口，正文写「资料缺失」，不按经验推定")

    # 8) 投递区 Office 留档（AHP 权重/评分/结论、低空指标等）：只登记，不解析、不重算
    if full:
        consumed = {str(report.resolve()) if report is not None else "", str(Path(params_office).resolve())}
        for item in index.files:
            if str(item.resolve()) in consumed:
                continue
            if item.suffix.lower() not in (".docx", ".doc", ".pdf", ".xlsx", ".xls", ".pptx"):
                continue
            note("reference", rel(item, workspace), "（人读留档，不参与计算）", "投递区原件")
        for archive in index.archives:
            note("reference", rel(archive, workspace), "（压缩包已解包登记）", "投递区压缩包")

    # 8.5) 形式参考件隔离（C7）：识别同类报告成品 → 只登记骨架，不进提资本体
    register_form_reference(workspace, index, project_name, args.no_timestamp)

    # 9) 财务参数（唯一必需输入）
    if not Path(params_office).is_file():
        print("[prepare_inputs] 缺必需输入：{}（取费/收入/税率/资金结构）→ 待补，不补数".format(
            rel(params_office, workspace)))
        write_json(workspace / "过程文件/01-上游取料/取料清单.json", {
            "schema": SCHEMA_LIST, "projectName": project_name, "scenario": args.scenario,
            "placed": placed, "missing": missing + [{
                "code": "INPUT-PARAMS", "what": "财务参数（取费标准/收入口径/税率/资金结构）",
                "need": "输入文件/ 放 财务参数.xlsx",
                "suggestion": "按估算附表填工程费/其他费/收入/成本/资金结构；这是唯一必需输入，缺失时停下等补齐，不补数"}],
            "status": "need-input", "notes": index.notes,
        }, args.no_timestamp)
        return 3
    params = parse_params(params_office)
    if not params["estimate"]["engineering"]:
        print("[prepare_inputs] 财务参数.xlsx 未读到工程费明细（表「工程费」缺 编号/名称/金额）→ 待补")
        return 3
    if not params["project"]["name"]:
        params["project"]["name"] = project_name
    write_json(staging / "05_过程/投资决策/财务参数.json", params, args.no_timestamp)
    note("convert", rel(params_office, workspace), "05_过程/投资决策/财务参数.json",
         "工程费 {} 项 / 其他费 {} 项 / 收入 {} 项".format(
             len(params["estimate"]["engineering"]), len(params["estimate"]["otherFees"]), len(params["revenue"])))

    # 10) 现状指标归属核对：cardo/lowair-indicators/v1 本身不带项目名，
    # 只能用「可选的项目名字段」和「范围面积 vs 改造面积」两个信号核对。
    # 对不上只提示、不阻断——换来换去张冠李戴比缺项更难发现（实测踩过：别的项目指标混进来）。
    warnings = []
    metrics_path = staging / "05_过程/低空解析/现状指标.json"
    if metrics_path.is_file():
        try:
            metrics = read_json(metrics_path)
        except Exception:  # noqa: BLE001
            metrics = {}
        own_name = str(params["project"].get("name") or "").strip()
        up_name = str((metrics or {}).get("projectName")
                      or ((metrics or {}).get("project") or {}).get("name") or "").strip()
        own_area = params["project"].get("area")
        up_area = ((metrics or {}).get("scope") or {}).get("areaM2")
        reason = ""
        if up_name and own_name and up_name not in own_name and own_name not in up_name:
            reason = "低空指标标注的项目是「{}」，与本项目「{}」不一致".format(up_name, own_name)
        elif up_area and own_area:
            try:
                ratio = float(up_area) / float(own_area)
            except (TypeError, ValueError, ZeroDivisionError):
                ratio = 1.0
            # 口径不同：scope.areaM2 是模型覆盖范围（bbox），改造面积是红线内面积，
            # 正常会差一些；只有差到「3 倍以上或 1/3 以下」才当疑似张冠李戴，避免天天误报。
            if ratio > 3.0 or ratio < 1.0 / 3.0:
                reason = ("低空指标覆盖范围 {:.0f} ㎡，是财务参数改造面积 {:.0f} ㎡ 的 {:.1f} 倍，"
                          "量级对不上").format(float(up_area), float(own_area), ratio)
        if reason:
            text_warn = ("{}：第 3 章表 3-1 会照此引用，疑似不是同一项目。"
                         "若确非本项目，请向 CardoOSGB 重新解析本项目航拍，"
                         "把 过程文件/06-2.5D网格指标/现状指标.json 复制进 输入文件/".format(reason))
            warnings.append({"code": "UPSTREAM-OSGB-PROJECT-MISMATCH", "what": text_warn})
            index.notes.append("提示：" + text_warn)
            print("[prepare_inputs] 提示：" + text_warn)

    # 10b) 指标口径登记（字段级）：影像纹理采样得到的绿化/道路覆盖是派生代理值。
    # 08 缺口词闸门据此决定「影像纹理采样推定」这一技术表述是否豁免——
    # 只认字段上的 derived:true，没有纹理诊断块就不登记，闸门照常拦。
    texture = (((metrics or {}).get("diagnostics") or {}).get("texture")
               if metrics_path.is_file() else None)
    if isinstance(texture, dict):
        write_json(workspace / "过程文件/01-上游取料/低空指标口径.json", {
            "schema": "cardo/invest-lowair-caliber/v1",
            "source": "过程文件/00-总控/引擎项目/05_过程/低空解析/现状指标.json#diagnostics.texture",
            "texture": {
                "derived": True,
                "grade": "derived-proxy",
                "note": texture.get("note") or "绿地覆盖、道路覆盖为影像纹理采样代理值，非平面用地面积",
            },
        }, args.no_timestamp)

    manifest = {
        "schema": SCHEMA_LIST,
        "projectName": params["project"]["name"],
        "scenario": args.scenario,
        "placed": placed,
        "missing": missing,
        "warnings": warnings,
        "status": "ok",
        "notes": index.notes,
    }
    write_json(workspace / "过程文件/01-上游取料/取料清单.json", manifest, args.no_timestamp)
    print("[prepare_inputs] 引擎项目装配完成：{} 项落位 / {} 项缺口（scenario={}）".format(
        len(placed), len(missing), args.scenario))
    for gap in missing:
        if gap.get("suggestion"):
            print("  - {}：{}（建议：{}）".format(gap["code"], gap["what"], gap["suggestion"]))
    return 0


if __name__ == "__main__":
    sys.exit(main())
