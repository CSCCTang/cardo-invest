#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""check_evidence.py -- 07 依据核对：分章稿的每段出处与每个数值是否真的查得到。

只认《依据对照表.json》里的登记结果，不重新猜来源：
  E-1 段段有源   每个正文段/表格/图注都有来源行（无源块 = 0）
  E-2 出处可查   每条引用的文件与键都能落到实际文件上（未解析 = 0）
  E-3 数值登记   出现数值的块必须是有源块（数值不能裸奔）
  E-4 未解析留痕 对照表把未解析项单列，不允许"写得出处但查不到"
  E-5 正文洁净   成果 docx 正文 0 处「来源：/（数据来源：…）」来源行（出处只进对照表）
  E-6 对照表回指 对照表行与正文块一一对应，且每行的落地文件真实存在
  E-7 数值同源   正文每个数值都要在登记来源里找到同一个值；派生值（占比/合计）必须
                 能由两个已登记数值复算，并登记成带双上游指针的派生行；
                 成品 docx 里的数值必须与登记表一致（改稿不改表 = 拦住）

用法：
    python check_evidence.py --project-dir <工作区> [--format json --out 结果.json]
退出码：0 全过；1 有不通过项。
"""

import argparse
import json
import os
import re

SOURCE_PREFIX = ("来源：", "来源:", "（数据来源", "(数据来源", "数据来源：", "数据来源:")

# 数值：带千分位、小数、可选百分号；单位另算（中文单位贴着数字也算同一枚数值）
NUM_TOKEN = re.compile(r"\d[\d,]*(?:\.\d+)?\s*%?")
# 逐值回指要比符号：正文「-10,973.00 万元」与字段 -10973 必须同号，
# 不能因为 NUM_TOKEN 不带负号就把正负当成同一个值（那样「值相等」也验不出方向性错误）。
SIGNED_TOKEN = re.compile(r"-?[\d,]+(?:\.\d+)?\s*%?")
# 交叉引用（表号/图号/章节号）、标准号与文号不是数据，核对前先剔掉，
# 否则「见 5.3」「表 7-1」「GB 50500—2024」会被当成无源数值。
CROSS_RE = re.compile(
    r"(?:表|图)\s*\d+\s*[-–‑]\s*\d+"
    r"|第\s*\d+(?:\.\d+)?\s*[章节条款项]"
    # 「见 9 章」「见 5.3 节」这类不带「第」的章号也是交叉引用，不是数据。
    r"|(?<![\d.,])\d{1,2}(?:\.\d+)?\s*[章节](?!\s*[、,，])"
    r"|\d+\.\d+(?=\s*(?:节|章|条))"
    r"|(?:见|详见|参见|对照|与|据|按|向|作为)\s*\d{1,2}\.\d{1,2}(?:\.\d{1,3})*(?:\s*[、,，]\s*\d{1,2}\.\d{1,2}(?:\.\d{1,3})*)*"
    # 「1.1 项目基本信息、7.1 投资估算」这类节号列表：一位小数、两位以内整数，
    # 且后面不跟单位或百分号——数据值一般带单位或两位小数，不会被误伤。
    # 「8.1.5.1」这类三级以上编号整体剔除，否则会被截成「5.1」误判成无源数值。
    r"|(?<![\d.,])\d+(?:\.\d+){2,}(?![\d%])"
    r"|(?<![\d.,])(?:[1-9]|1[0-2])\.\d(?![\d%.])(?!\s*(?:万|亿|元|㎡|个|处|人|户|家|栋|米|倍|年))"
    r"|(?:按|据)\s*\d+(?:\.\d+)?\s*%"
    r"|(?:GB|GB/T|CJJ|JGJ|CJ/T|DG|DB|JG)\s*[/T]*\s*\d+(?:\s*[—–-]\s*\d{4})?"
    r"|〔\s*\d{4}\s*〕\s*\d+\s*号"
    r"|(?:19|20)\d{2}\s*年"
    # 坐标系编号与摘要算法不是数据：「EPSG:4544」「sha256」曾被当成无源数值。
    r"|EPSG\s*[:：]?\s*\d{4,5}"
    r"|sha-?\d+"
    # 月日与「公告第N号」也是引用不是数据：沿革句「自 5 月 28 日起施行」「公告第35号」
    # 曾被判成两个裸奔数值（27、35），其实它们的出处就是依据库条目本身。
    r"|\d{1,2}\s*月\s*\d{1,2}\s*日"
    r"|公告\s*第\s*\d+\s*号")

TOL_DIGITS = 6


def _decimals(raw):
    """展示精度：只数数值本身的小数位，单位/空格不算。

    「5.8年」曾按「.8年」数出 2 位小数 → 容差被压到 0.005，5.8 与 5.83 被误判不一致。
    """
    body = str(raw).replace(",", "").rstrip("%").strip()
    match = re.search(r"\d+(?:\.(\d+))?", body)
    return len(match.group(1) or "") if match else 0


def numbers_in_text(text):
    """→ [(原文, 数值, 是否百分比)]；交叉引用先剔除。"""
    text = CROSS_RE.sub(" ", text or "")
    out = []
    for match in NUM_TOKEN.finditer(text):
        raw = match.group(0).strip()
        percent = raw.endswith("%")
        try:
            value = float(raw.rstrip("%").replace(",", "").strip())
        except ValueError:
            continue
        out.append((raw, value, percent))
    return out


def _close(value, target, raw=""):
    """按原文显示精度比对：78.8 与 78.79 视为一致（0.05 容差）。

    没有小数位的 token（「3年」「1」「100%」）容差必须是 1e-6：早先一律给 0.5，
    「3年」能配上 2.65，「4年」能配上 4.38，逐数明细因此错绑 49 行。
    """
    if not _decimals(raw):
        return abs(value - float(target)) <= 1e-6
    tol = 0.5 * 10 ** (-_decimals(raw))
    return abs(value - float(target)) <= max(tol, 1e-6)


# ---------------------------------------------------------------- 来源取值

def _source_error(src, root):
    """来源行现场可解析性：文件要在盘上，JSON 的「文件#键」要真能取到节点。"""
    ref = str(src.get("ref") or "")
    # 「标准值:」是 AHP 查表值（RI 表 / CR 判据）的登记来源，不对应工作区文件。
    if ref.startswith("标准值:"):
        return ""
    real = str(src.get("path") or "")
    if real and not os.path.isabs(real):
        real = os.path.join(root, real)
    path_part = ref.split("#")[0].strip()
    pointer = ref.partition("#")[2].strip() if "#" in ref else ""
    if not real and path_part:
        real = path_part if os.path.isabs(path_part) else os.path.join(root, path_part)
    if not real or not os.path.isfile(real):
        return "来源文件不存在：{}".format(ref)
    if pointer and real.lower().endswith(".json"):
        try:
            with open(real, encoding="utf-8") as handle:
                node = json.load(handle)
        except (OSError, ValueError):
            return "来源文件读不出：{}".format(ref)
        if not _dig(node, pointer)[1]:
            return "键取不到：{}".format(ref)
    return ""


def _dig(node, pointer):
    """按 a/b/c 或 a.b.c 取节点；取不到返回 (None, False)。"""
    for key in [part for part in re.split(r"[/.]", pointer or "") if part]:
        if isinstance(node, list):
            try:
                node = node[int(key)]
            except (ValueError, IndexError):
                return None, False
        elif isinstance(node, dict) and key in node:
            node = node[key]
        else:
            return None, False
    return node, True


def _walk_numbers(node, out):
    if isinstance(node, dict):
        for value in node.values():
            _walk_numbers(value, out)
    elif isinstance(node, list):
        for value in node:
            _walk_numbers(value, out)
    elif isinstance(node, bool):
        return
    elif isinstance(node, (int, float)):
        out.append(float(node))
    elif isinstance(node, str):
        for match in NUM_TOKEN.finditer(node):
            raw = match.group(0).strip()
            try:
                value = float(raw.rstrip("%").replace(",", "").strip())
            except ValueError:
                continue
            out.append(value / 100 if raw.endswith("%") else value)


def _file_numbers(path, pointer, cache):
    """取一个来源文件里的全部数值（JSON 按指针取值域；xlsx 取单元格；docx/md 取文本）。"""
    key = (os.path.abspath(path), pointer or "")
    if key in cache:
        return cache[key]
    values = []
    lower = path.lower()
    try:
        if lower.endswith(".json"):
            with open(path, encoding="utf-8") as handle:
                node = json.load(handle)
            if pointer:
                node, ok = _dig(node, pointer)
                if not ok:
                    node = None
            if node is not None:
                _walk_numbers(node, values)
        elif lower.endswith((".xlsx", ".xlsm")):
            import openpyxl

            workbook = openpyxl.load_workbook(path, data_only=True, read_only=True)
            for sheet in workbook.worksheets:
                for row in sheet.iter_rows(values_only=True):
                    for cell in row:
                        if isinstance(cell, bool) or cell is None:
                            continue
                        if isinstance(cell, (int, float)):
                            values.append(float(cell))
                        elif isinstance(cell, str):
                            for match in NUM_TOKEN.finditer(cell):
                                raw = match.group(0).strip()
                                try:
                                    values.append(float(raw.rstrip("%").replace(",", "").strip()))
                                except ValueError:
                                    continue
            workbook.close()
        elif lower.endswith(".docx"):
            import docx

            document = docx.Document(path)
            texts = [para.text for para in document.paragraphs]
            for table in document.tables:
                for row in table.rows:
                    for cell in row.cells:
                        texts.append(cell.text)
            for text in texts:
                for match in NUM_TOKEN.finditer(text):
                    raw = match.group(0).strip()
                    try:
                        values.append(float(raw.rstrip("%").replace(",", "").strip()))
                    except ValueError:
                        continue
        else:
            with open(path, encoding="utf-8", errors="ignore") as handle:
                for match in NUM_TOKEN.finditer(handle.read()):
                    raw = match.group(0).strip()
                    try:
                        values.append(float(raw.rstrip("%").replace(",", "").strip()))
                    except ValueError:
                        continue
    except (OSError, ValueError, KeyError):
        values = []
    cache[key] = values
    return values


def _expand(values):
    """数值池扩展：0.788 → 同时认 78.8（正文按百分数写，来源按小数存）。"""
    out = set()
    for value in values:
        out.add(round(value, TOL_DIGITS))
        if 0 < abs(value) <= 2:
            out.add(round(value * 100, TOL_DIGITS))
        if abs(value) >= 1:
            out.add(round(value / 100, TOL_DIGITS))
    return out


def _file_counts(path, cache):
    """JSON 里各列表的长度：正文「共 16 项指标」这类计数是派生值，得能对上原始清单。"""
    key = ("counts", os.path.abspath(path))
    if key in cache:
        return cache[key]
    counts = set()
    if path.lower().endswith(".json"):
        try:
            with open(path, encoding="utf-8") as handle:
                stack = [json.load(handle)]
            while stack:
                node = stack.pop()
                if isinstance(node, list):
                    counts.add(len(node))
                    stack.extend(node)
                elif isinstance(node, dict):
                    stack.extend(node.values())
        except (OSError, ValueError):
            counts = set()
    cache[key] = counts
    return counts


def _file_ratios(path, cache):
    """结构化占比：同层「分项 ÷ 合计 ×100」与「列表项 ÷ 合计 ×100」。

    正文的占比、构成比都从这里来；不拿任意两个数相除去凑——
    否则随便改一个百分比也可能"碰巧"等于某两个数之比，拦不住。
    """
    key = ("ratios", os.path.abspath(path))
    if key in cache:
        return cache[key]
    ratios = {"percent": [], "ratio": []}
    small = []      # (值, 字段路径)：权重/占比，判别矩阵 a_ij = w_i/w_j 由此还原

    def siblings(node, prefix):
        """同一层里的小规模数值组：分项占比与两两比值（判别矩阵就是靠 w_i/w_j 还原的）。"""
        values = []
        if isinstance(node, dict):
            values = [(name, float(value)) for name, value in node.items()
                      if isinstance(value, (int, float)) and not isinstance(value, bool)]
        elif isinstance(node, list):
            values = [(str(index), float(value)) for index, value in enumerate(node)
                      if isinstance(value, (int, float)) and not isinstance(value, bool)]
        if len(values) < 2 or len(values) > 12:
            return
        total_key, total = "", None
        for name, value in values:
            if "total" in name.lower() or "合计" in name or "总额" in name:
                total_key, total = name, value
                break
        if total is None:
            total_key, total = max(values, key=lambda pair: abs(pair[1]))
        if total:
            for name, value in values:
                if name == total_key:
                    continue
                part_path = "{}.{}".format(prefix, name) if prefix else str(name)
                total_path = "{}.{}".format(prefix, total_key) if prefix else str(total_key)
                ratios["percent"].append((value / total * 100.0, value, total,
                                          "{}.{}".format(prefix, name) if prefix else name,
                                          part_path, total_path))
        if abs(total) <= 12:   # 判别矩阵量级（1/9–9），只在这类小矩阵里认两两比值
            for name, value in values:
                for other_name, other in values:
                    if name == other_name or not other:
                        continue
                    ratios["ratio"].append((value / other, value, other,
                                            "{}.{}÷{}".format(prefix, name, other_name)
                                            if prefix else "{}÷{}".format(name, other_name),
                                            "{}.{}".format(prefix, name) if prefix else str(name),
                                            "{}.{}".format(prefix, other_name) if prefix else str(other_name)))

    def walk(node, prefix):
        siblings(node, prefix)
        if isinstance(node, dict):
            for name, value in node.items():
                walk(value, "{}.{}".format(prefix, name) if prefix else str(name))
        elif isinstance(node, list):
            for index, value in enumerate(node):
                walk(value, "{}.{}".format(prefix, index) if prefix else str(index))
        elif isinstance(node, (int, float)) and not isinstance(node, bool) and 0 < node <= 1:
            if len(small) < 60:
                small.append((float(node), prefix))

    if path.lower().endswith(".json"):
        try:
            with open(path, encoding="utf-8") as handle:
                walk(json.load(handle), "")
        except (OSError, ValueError):
            pass
    # 权重间的两两比值：层级判别矩阵是由子权重按 a_ij = w_i/w_j 还原的，
    # 表里出现的 1.1482 这类数必须能由同一份权重表复算出来。
    for value, path in small:
        for other, other_path in small:
            if other and abs(value - other) > 1e-12:
                ratios["ratio"].append((value / other, value, other, "同源权重比", path, other_path))
    cache[key] = ratios
    return ratios


def _find_derivation(value, percent, pool, counts=(), ratios=None):
    """在来源数值池里找可复现的派生关系，返回 (公式, a, b, 上游a, 上游b) 或 None。

    * 占比：只认「分项 ÷ 合计 ×100」这类结构化关系，不认任意两数之比；
    * 比值：只认同层小矩阵内的两两相除（判别矩阵 a_ij = w_i/w_j 的还原）；
    * 计数：等于来源里某个清单的条数（如「共 16 项指标」）；
    * 一致性：CI = (λmax − n)/(n − 1)，λmax 已登记、n 为矩阵阶数。

    上游用 (文件相对路径, 字段路径) 表示；只有结构化关系带得出字段路径，
    计数/一致性只能给到文件，由调用方决定要不要再按值唯一匹配。
    """
    book = ratios if isinstance(ratios, dict) else {"percent": [], "ratio": []}
    if percent:
        for cand, part, total, label, part_path, total_path, rel in book["percent"]:
            if _close(value, cand, str(round(value, 2))):
                return ("{} ÷ {} ×100（{}）".format(_num(part), _num(total), label), part, total,
                        (rel, part_path), (rel, total_path))
        return None
    for cand, part, total, label, part_path, total_path, rel in book["ratio"]:
        if _close(value, cand, str(round(value, 4))):
            return ("{} ÷ {}（{}）".format(_num(part), _num(total), label), part, total,
                    (rel, part_path), (rel, total_path))
    if float(value).is_integer() and int(value) in counts:
        return ("原始清单条数 {} 项".format(int(value)), float(value), float(value), ("", ""), ("", ""))
    for lam in pool:
        for n in counts:
            if n and n >= 3 and abs((lam - n) / (n - 1) - value) <= max(0.0001, abs(value) * 1e-4):
                return ("({}−{})/({}−1)".format(_num(lam), n, n), lam, float(n), ("", ""), ("", ""))
    return None


def _resolve_leaf(pair, number, leaves):
    """把派生值的一个上游落到唯一字段，返回 (指针, 是否唯一)。

    先按登记的结构化路径精确取值（这是「真正被除/被乘的那个字段」）；
    路径取不到时只在同值叶子唯一的情况下退用，多个叶子同值 = 指不准，判不唯一。
    """
    rel, path = pair if isinstance(pair, (tuple, list)) and len(pair) == 2 else ("", "")
    if path:
        want = "{}#{}".format(rel, path) if rel else path
        for cand, ref in leaves:
            if ref == want and abs(cand - number) <= max(1e-9, abs(number) * 1e-9):
                return want, True
    hits = {ref for cand, ref in leaves
            if abs(cand - number) <= max(1e-9, abs(number) * 1e-9)}
    if len(hits) == 1:
        return hits.pop(), True
    # CA 的阶数 n 是结构计数：字段名就叫 n / order / 阶数，指到它而不是随便一个同值数
    named = {ref for ref in hits
             if ref.rsplit(".", 1)[-1] in ("n", "order", "count", "阶数", "矩阵阶数")}
    if len(named) == 1:
        return named.pop(), True
    # 计数/一致性这类没有字段路径的派生：同值叶子全落在同一份文件时指到文件本身，
    # 不硬凑某个"同值叶子"（那正是上一版把 0.0344 指到 0.0356 的原因）。
    files = {ref.split("#")[0] for ref in hits}
    if not path and len(files) == 1:
        return files.pop(), True
    return "", False


def _num(value):
    """权重/占比这类小数按 4 位显示：只留 2 位会把 0.0556÷0.0484 印成「0.06 ÷ 0.06」。"""
    digits = 4 if abs(value) < 1 else 2
    return ("{:,.{}f}".format(value, digits)).rstrip("0").rstrip(".")


# ---------------------------------------------------------------- 各项检查

def clean_body_source_lines(docx_path):
    """读成果 docx 正文，返回（段数, 来源行列表）；文件不存在返回 (None, [])。"""
    if not os.path.isfile(docx_path):
        return None, []
    try:
        from docx import Document
    except ImportError:
        return None, []
    doc = Document(docx_path)
    hits = []
    for para in doc.paragraphs:
        text = (para.text or "").strip()
        if text.startswith(SOURCE_PREFIX):
            hits.append(text[:60])
    return len(doc.paragraphs), hits


def _resolved_values(node):
    """把一个指针解出来的节点摊成数值集合（字符串里的数字也算，带符号）。"""
    values = []
    if isinstance(node, bool) or node is None:
        return values
    if isinstance(node, (int, float)):
        values.append(float(node))
    elif isinstance(node, str):
        for match in SIGNED_TOKEN.finditer(node):
            raw = match.group(0).strip()
            try:
                number = float(raw.rstrip("%").replace(",", "").strip())
            except ValueError:
                continue
            values.append(number / 100 if raw.endswith("%") else number)
    elif isinstance(node, (list, tuple)):
        for item in node:
            if isinstance(item, (list, tuple)) and len(item) == 2 \
                    and isinstance(item[1], (int, float)) and not isinstance(item[1], bool):
                values.append(float(item[1]))
            else:
                values.extend(_resolved_values(item))
    elif isinstance(node, dict):
        _walk_numbers(node, values)
    return values


def _value_matches(node, token):
    """解出的值是否等于正文数值：按正文显示位数容差、百分比 ×100、负号必须一致。"""
    match = SIGNED_TOKEN.search(token or "")
    if not match:
        return True
    raw = match.group(0).strip()
    percent = raw.endswith("%")
    try:
        value = float(raw.rstrip("%").replace(",", "").strip())
    except ValueError:
        return True
    scale = 100 if percent else 1
    if percent:
        # 比率在来源里可能存小数（0.788），也可能存百分数（78.8）；两种都认。
        # 正文用方向词写幅度时（「压减 57.7%」），来源存的是带负号的临界点 -57.69：
        # 正数 token 也按绝对值收一次；负数 token 仍要求同号，方向错误照样拦得住。
        hits = _resolved_values(node)
        if any(_close(value, number * scale, raw) or _close(value, number, raw)
               for number in hits):
            return True
        return value > 0 and any(_close(value, abs(number) * scale, raw)
                                 or _close(value, abs(number), raw) for number in hits)
    # 表格把单位提到表头（「FIRR（%）」）时单元格只剩 -2.55，来源存的是 -0.0255：
    # 不带 % 的百分数量级（|值| ≤ 100）也按小数点后两位折算一次，否则逐值回指必失败。
    if abs(value) <= 100:
        return any(_close(value, number * scale, raw) or _close(value / 100.0, number, raw)
                   for number in _resolved_values(node))
    return any(_close(value, number * scale, raw) for number in _resolved_values(node))


def _has_source(status):
    """块级状态有两种「有源」：全部数值绑上（有源）、部分绑上（部分有源）。

    05 不再把部分绑定伪装成「有源」；这里必须同步认这两种，否则 E-2/E-6 会把
    部分绑定块整块跳过，等于对真缺口视而不见。
    """
    text = str(status or "")
    return text == "有源" or text.startswith("部分有源") or text.startswith("人工核")


def register_links(root, rows, chapter_dir=""):
    """对照表 ↔ 正文块一一对应 + 每行落地文件真实存在 + 逐值指针真实可解。

    E-6 的核心不是「文件在不在」，而是「文件#键 在不在、解出来的值是不是正文那个数」：
    早先只校验文件存在，坏样例把指针改成不存在的键也照样 exit=0（上一轮复验失守）。
    """
    problems = []
    keys = [(r.get("chapter"), r.get("section"), r.get("seq")) for r in rows]
    dup = sorted({k for k in keys if keys.count(k) > 1})
    for item in dup:
        problems.append("重复行：{} {}第{}段".format(*item))
    cache = {}
    for row in rows:
        if not _has_source(row.get("status")):
            continue
        where = ""
        if chapter_dir:
            where = _locate_line(chapter_dir, row.get("chapter"), row.get("excerpt"))
        where = where or "{} {} 第{}段".format(
            row.get("chapter"), row.get("section"), row.get("seq"))
        files = []
        for src in row.get("sources") or []:
            if src.get("status") != "可解析":
                continue
            # 优先用对照表登记的结构化落地路径（path）；旧版对照表没有 path 时才从
            # note 里取。note 里可能带「（原路径无此文件，按文件名解析到 …）」这类
            # 说明文字，早期直接拿整串当路径，导致在位文件被判「落地文件不存在」。
            path = str(src.get("path") or "").strip()
            if not path:
                note = str(src.get("note") or "")
                path = note.split("#")[0]
                path = re.split(r"[（(]", path)[0]
                path = path.replace("（按 id 定位）", "").strip()
            if path:
                files.append(path)
        if not files:
            problems.append("{}：有源但没有任何可解析文件".format(where))
            continue
        missing = [p for p in files
                   if not (os.path.isfile(p) if os.path.isabs(p)
                           else os.path.isfile(os.path.join(root, p)))]
        if missing:
            problems.append("{}：落地文件不存在（{}）".format(where, "；".join(missing[:3])))
        # 逐值回指：valueRefs 里每个「文件#键」都要能真正取到那个数。
        for item in row.get("valueRefs") or []:
            ref = str(item.get("ref") or "")
            if not ref:
                continue                     # 未定位值由 E-7 继续核，不在这里重复报
            if ref.startswith("标准值:"):
                continue                     # AHP 查表标准值不对应工作区文件
            token = str(item.get("number") or "")
            path_part, _, pointer = ref.partition("#")
            real = path_part if os.path.isabs(path_part) else os.path.join(root, path_part)
            if not os.path.isfile(real):
                problems.append("{}：值 {} 的出处文件不存在（{}）".format(where, token, ref))
                continue
            if pointer and real.lower().endswith(".json"):
                try:
                    with open(real, encoding="utf-8") as handle:
                        node = json.load(handle)
                except (OSError, ValueError):
                    problems.append("{}：值 {} 的出处文件读不出来（{}）".format(where, token, ref))
                    continue
                node, ok = _dig(node, pointer)
                if not ok:
                    problems.append("{}：值 {} 指向的键不存在（{}）".format(where, token, ref))
                    continue
                if not _value_matches(node, token):
                    shown = item.get("value")
                    problems.append("{}：值 {} 与 {} 解出的 {} 不等".format(
                        where, token, ref, shown if shown is not None else "空值"))
                continue
            numbers = _file_numbers(real, "", cache)
            if numbers and not _value_matches(numbers, token):
                problems.append("{}：值 {} 在 {} 里找不到相同数值".format(where, token, ref))
    return problems


def _row_files(root, row):
    files = []
    for src in row.get("sources") or []:
        if src.get("status") != "可解析":
            continue
        path = str(src.get("path") or "").strip()
        if not path:
            continue
        real = path if os.path.isabs(path) else os.path.join(root, path)
        if not os.path.isfile(real):
            continue
        ref = str(src.get("ref") or "")
        pointer = ref.split("#", 1)[1] if "#" in ref else ""
        files.append((real, pointer, path))
    return files


def _file_leaves(path, cache):
    """来源文件里每个数值的定位（JSON 取键路径，xlsx 取 工作表!单元格），供派生值写双上游。"""
    key = ("leaves", os.path.abspath(path))
    if key in cache:
        return cache[key]
    leaves = []

    def walk(node, prefix):
        if isinstance(node, dict):
            for name, value in node.items():
                walk(value, "{}.{}".format(prefix, name) if prefix else str(name))
        elif isinstance(node, list):
            for index, value in enumerate(node):
                walk(value, "{}.{}".format(prefix, index) if prefix else str(index))
        elif isinstance(node, bool) or node is None:
            return
        elif isinstance(node, (int, float)):
            leaves.append((prefix, float(node)))
        elif isinstance(node, str):
            for match in NUM_TOKEN.finditer(node):
                raw = match.group(0).strip()
                try:
                    value = float(raw.rstrip("%").replace(",", "").strip())
                except ValueError:
                    continue
                leaves.append((prefix, value / 100 if raw.endswith("%") else value))

    lower = path.lower()
    try:
        if lower.endswith(".json"):
            with open(path, encoding="utf-8") as handle:
                walk(json.load(handle), "")
        elif lower.endswith((".xlsx", ".xlsm")):
            import openpyxl

            workbook = openpyxl.load_workbook(path, data_only=True, read_only=True)
            for sheet in workbook.worksheets:
                for row in sheet.iter_rows():
                    for cell in row:
                        if isinstance(cell.value, (int, float)) and not isinstance(cell.value, bool):
                            leaves.append(("{}!{}".format(sheet.title, cell.coordinate),
                                           float(cell.value)))
            workbook.close()
    except (OSError, ValueError, KeyError):
        leaves = []
    cache[key] = leaves
    return leaves


def _locate_line(chapter_dir, chapter, excerpt):
    """把对照表的一段正文定位回分章稿的 文件:行（校验未通过说明要能指出位置）。

    登记块有正文段，也有整张表（excerpt 是「a|b|c」的管表片段，md 里一行只是一行表格），
    单拿前 18 个字去匹配整行会落空（实测 1.1 那张表只报到「第7段」）。这里把表格
    拆成单元格，取最长、最能在某一行里出现的片段定位到具体那一行。
    """
    name = str(chapter or "")
    path = os.path.join(chapter_dir, name + ".md")
    if not os.path.isfile(path):
        return ""
    plain = re.sub(r"\s+", "", (excerpt or ""))
    if not plain:
        return ""
    cells = [re.sub(r"\s+", "", item) for item in re.split(r"[|｜\n]", excerpt or "")]
    cells = sorted({item for item in cells if len(item) >= 6}, key=len, reverse=True)
    needles = [plain[:18]] + [item[:24] for item in cells[:4]]
    is_table = "|" in (excerpt or "")
    with open(path, encoding="utf-8") as handle:
        lines = handle.readlines()
    for needle in needles:
        if len(needle) < 6:
            continue
        for index, line in enumerate(lines, start=1):
            plain_line = re.sub(r"\s+", "", line)
            # 表格块只在管表行里找，否则会落到正文里第一次出现项目名的那一句
            if is_table and not plain_line.startswith("|"):
                continue
            if needle in plain_line:
                return "{}:{}".format(path.replace(os.sep, "/"), index)
    return ""


def check_numbers(root, rows, chapter_dir, cache):
    """E-7 主体：逐块核对正文数值是否与登记来源同源。

    返回 (problems, derived, detail)：problems 每条都带「文件:行」。
    """
    problems, derived, detail = [], [], []
    for row in rows:
        numbers = row.get("numbers") or []
        if not numbers:
            continue
        files = _row_files(root, row)
        if not files:
            continue
        # 数值池按「被引用的整份文件」取：同一份投资估算里，正文写总额、出处标分项，
        # 值仍然查得到；指针只用来说明出处位置，不缩小可信数值域。
        pool, counts, leaves = [], set(), []
        ratios = {"percent": [], "ratio": []}
        for real, _pointer, rel in files:
            pool.extend(_file_numbers(real, "", cache))
            counts |= _file_counts(real, cache)
            book = _file_ratios(real, cache)
            for kind in ("percent", "ratio"):
                ratios[kind].extend(tuple(item) + (rel,) for item in book[kind])
            leaves.extend((value, "{}#{}".format(rel, pointer) if pointer else rel)
                          for pointer, value in _file_leaves(real, cache))
        allowed = _expand(pool)
        # AHP 查表标准值（RI 0.9 / CR 0.1）不对应本项目测算值，E-7 不按来源池核对。
        standard_tokens = {str(item.get("number")) for item in row.get("valueRefs") or []
                           if item.get("status") == "标准值"}
        misses = []
        where = _locate_line(chapter_dir, row.get("chapter"), row.get("excerpt")) or \
            "{} {} 第{}段".format(row.get("chapter"), row.get("section"), row.get("seq"))
        for raw in numbers:
            if raw in standard_tokens:
                detail.append({"where": where, "number": raw, "kind": "标准值"})
                continue
            parsed = numbers_in_text(raw)
            if not parsed:
                continue
            _, value, percent = parsed[0]
            if any(_close(value, cand, raw) for cand in allowed):
                detail.append({"where": where, "number": raw, "kind": "直接同源"})
                continue
            # 正文按百分数写、来源按小数存（290.88% ↔ 2.9088）：_expand 的 ×100 只覆盖 ≤2 的比率，
            # ROI 这类大于 2 的比率漏在外面，必须按 token 自身是否带 % 再折算一次。
            if percent and any(_close(value / 100.0, cand, raw) for cand in allowed):
                detail.append({"where": where, "number": raw, "kind": "直接同源"})
                continue
            hit = _find_derivation(value, percent, pool, counts, ratios)
            if hit:
                formula, a, b, up_a, up_b = hit
                upstream, unique = [], True
                for pair, number in ((up_a, a), (up_b, b)):
                    ref, ok = _resolve_leaf(pair, number, leaves)
                    if ref and ref not in upstream:
                        upstream.append(ref)
                    unique = unique and ok
                entry = {"chapter": row.get("chapter"), "section": row.get("section"),
                         "seq": row.get("seq"), "number": raw, "formula": formula,
                         "upstream": upstream, "unique": unique,
                         "a": a, "b": b}
                derived.append(entry)
                row.setdefault("derived", []).append(entry)
                detail.append({"where": where, "number": raw, "kind": "派生"})
                if not unique:
                    # 成员同值：指针只能指到"某个同值字段"，不是真正被除/被乘的那个
                    problems.append(
                        "{}（{} 第{}段）：派生值 {} = {} 的上游不唯一（候选 {}），"
                        "须在来源里登记唯一字段路径".format(
                            where, row.get("section"), row.get("seq"), raw, formula,
                            "；".join(upstream) or "无"))
                continue
            misses.append(raw)
            detail.append({"where": where, "number": raw, "kind": "无源"})
        if misses:
            problems.append("{}（{} 第{}段）：{} 在登记来源里查不到（{}）".format(
                where, row.get("section"), row.get("seq"),
                "、".join(misses[:6]), "；".join(rel for _r, _p, rel in files)[:120]))
    return problems, derived, detail


def _locate_text(chapter_dir, text):
    """把成品 docx 的一段文字定位回分章稿的 文件:行（docx 层也要报到行，不只到段）。"""
    plain = re.sub(r"\s+", "", (text or ""))
    if not plain or not chapter_dir or not os.path.isdir(chapter_dir):
        return ""
    # 改过的段落含被篡改的数字，长前缀会对不上；逐步退到短前缀仍能定位到该段
    for length in (18, 12, 8):
        needle = plain[:length]
        if len(needle) < 6:
            continue
        for name in sorted(os.listdir(chapter_dir)):
            if not (name.upper().startswith("CH-") and name.endswith(".md")):
                continue
            path = os.path.join(chapter_dir, name)
            try:
                with open(path, encoding="utf-8") as handle:
                    for index, line in enumerate(handle, start=1):
                        if needle in re.sub(r"\s+", "", line):
                            return "{}:{}".format(path.replace(os.sep, "/"), index)
            except OSError:
                continue
    return ""


def check_docx_numbers(docx_path, rows, chapter_dir=""):
    """E-7 附带：成品 docx 的数值必须与登记表一致（改了稿没改表 = 篡改）。

    docx 由分章稿渲染而来，数值集合应当同源；出现表里没有的数值，
    说明成品被直接改写或渲染过程掺了数，一律拦下。定位到分章稿 文件:行。
    """
    if not os.path.isfile(docx_path):
        return []
    try:
        from docx import Document
    except ImportError:
        return []
    known = set()
    for row in rows:
        for raw in row.get("numbers") or []:
            for _raw, value, _pct in numbers_in_text(raw):
                known.add(round(value, TOL_DIGITS))
    if not known:
        return []
    document = Document(docx_path)
    problems = []
    started = False
    for index, para in enumerate(document.paragraphs, start=1):
        text = (para.text or "").strip()
        style = (para.style.name or "") if para.style is not None else ""
        if style.startswith("Heading 1"):
            started = True
        if not started or style.startswith("Heading") or not text:
            continue
        for raw, value, _pct in numbers_in_text(text):
            if any(_close(value, cand, raw) for cand in known):
                continue
            where = _locate_text(chapter_dir, text) or os.path.basename(docx_path)
            problems.append("{}（docx 第{}段）：数值 {} 不在登记表中（表与正文不一致）：{}".format(
                where, index, raw, text[:36]))
            break
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                for raw, value, _pct in numbers_in_text(cell.text):
                    if any(_close(value, cand, raw) for cand in known):
                        continue
                    where = _locate_text(chapter_dir, cell.text) or os.path.basename(docx_path)
                    problems.append("{}（docx 表格单元格）：数值 {} 不在登记表中（{}…）".format(
                        where, raw, cell.text[:24]))
                    break
                else:
                    continue
                break
            else:
                continue
            break
    return problems


def write_failure_note(root, problems, derived, link_problems=None):
    """校验未通过时写《校验未通过说明》，把问题定位到 文件:行。

    E-6 与 E-7 一样要出说明：早先只有 E-7 落文件，E-6 失败时用户只看到
    依据核对.json，拿不到「哪一行、哪一段」。
    """
    path = os.path.join(root, "过程文件", "07-回读自检", "校验未通过说明.md")
    os.makedirs(os.path.dirname(path), exist_ok=True)
    lines = ["# 校验未通过说明", "",
             "> 07 依据核对未通过：对照表回指或正文数值与登记来源不一致。",
             "> 处置：把该数改回来源值，或在对照表补登记（派生值须写明两个上游）。", ""]
    if link_problems:
        lines += ["## 1. 对照表回指失败", ""]
        lines += ["- {}".format(item) for item in link_problems]
    lines += ["## 2. 无出处数值", ""]
    for problem in problems:
        lines.append("- {}".format(problem))
    if derived:
        lines += ["", "## 3. 已认的派生值（双上游）", ""]
        for item in derived:
            lines.append("- {} {}：{} = {}（上游：{}）".format(
                item.get("chapter"), item.get("section"), item.get("number"),
                item.get("formula"), "；".join(item.get("upstream") or [])))
    with open(path, "w", encoding="utf-8") as handle:
        handle.write("\n".join(lines) + "\n")
    return path


def load(path):
    with open(path, encoding="utf-8") as handle:
        return json.load(handle)


def main(argv=None):
    ap = argparse.ArgumentParser(description="Invest 依据核对")
    ap.add_argument("--project-dir", required=True)
    ap.add_argument("--register", default="")
    ap.add_argument("--format", choices=["text", "json"], default="text")
    ap.add_argument("--out", default="")
    ap.add_argument("--no-timestamp", action="store_true")
    args = ap.parse_args(argv)
    root = os.path.abspath(args.project_dir)
    register_path = args.register or os.path.join(
        root, "过程文件/05-章节成文/分章/依据对照表.json")
    chapter_dir = os.path.dirname(register_path)

    checks = []

    def add(cid, ok, detail):
        checks.append({"id": cid, "ok": bool(ok), "detail": detail})

    derived = []
    if not os.path.isfile(register_path):
        add("E-0", False, "缺依据对照表：{}（先跑 05 分章成文）".format(register_path))
        payload = {"ok": False, "checks": checks}
    else:
        data = load(register_path)
        stats = data.get("stats") or {}
        rows = data.get("rows") or []
        problems = data.get("problems") or []

        no_source = [r for r in rows if r.get("status") == "无源"]
        add("E-1", not no_source,
            "{} 块无来源{}".format(len(no_source),
                                "：" + "；".join("{} {}第{}段".format(r.get("chapter"), r.get("section"),
                                                                   r.get("seq")) for r in no_source[:5])
                                if no_source else ""))

        # E-2 每次重算来源行可解析性：直接读 05 缓存的 problems 等于只校验历史状态，
        # 事后把 sources.ref 的键改成不存在的，07 仍会 exit=0（与 E-6 双重标准）。
        source_bad = []
        for row in rows:
            if row.get("status") != "有源":
                continue
            for src in row.get("sources") or []:
                if src.get("status") != "可解析":
                    continue
                error = _source_error(src, root)
                if not error:
                    continue
                # E-2 必须落到「文件:行」：只有章/节/段时，读者翻不到那一行。
                line = row.get("line")
                if line:
                    where = "{}:{}({} 第{}段)".format(
                        str(row.get("chapter") or "") + ".md", line,
                        row.get("section"), row.get("seq"))
                else:
                    where = "{} {}第{}段".format(row.get("chapter"), row.get("section"),
                                              row.get("seq"))
                message = "{}：{}".format(where, error)
                if message not in source_bad:
                    source_bad.append(message)
        add("E-2", not source_bad and not problems,
            "{} 条来源行现场可解析".format(sum(len(r.get("sources") or []) for r in rows))
            if not source_bad and not problems
            else "{} 条引用查不到：{}".format(len(source_bad) or len(problems),
                                        "；".join((source_bad or problems)[:5])))

        naked = [r for r in rows if r.get("numbers")
                 and not (r.get("sources") or r.get("boundCount"))]
        add("E-3", not naked, "{} 个含数值的块无有效来源".format(len(naked)) if naked
            else "含数值的块全部有来源（逐值绑定情况见附件对照表「人工核」页）")

        add("E-4", bool(stats.get("blocks")), "登记 {} 章 / {} 块，有源 {}；逐值绑定 {}/{}".format(
            stats.get("chapters"), stats.get("blocks"), stats.get("withSource"),
            stats.get("valueBound", 0),
            stats.get("valueBound", 0) + stats.get("valueUnbound", 0)))

        # 成稿口径：来源行退出正文后，用「对照表回指 + 落地文件可查」等价替换。
        link = register_links(root, rows, chapter_dir)
        add("E-6", not link,
            "{} 行与正文块一一对应、逐值指针真实可解".format(len(rows)) if not link
            else "；".join(link[:5]))

        docx_path = os.path.join(root, "成果文件", "成品版", "项目投资决策建议书",
                                 "项目投资决策建议书.docx")
        if not os.path.isfile(docx_path):
            docx_path = os.path.join(root, "成果文件", "项目投资决策建议书.docx")
        if not os.path.isfile(docx_path):
            # 07 在 08 之前跑，成品还在暂存区：按暂存区找，别把「还没改名」当成缺件
            staged = os.path.join(root, "过程文件/00-总控/引擎项目/04_成果报告")
            if os.path.isdir(staged):
                hits = [name for name in sorted(os.listdir(staged))
                        if name.startswith("项目投资决策建议书") and name.endswith(".docx")]
                if hits:
                    docx_path = os.path.join(staged, hits[0])
        n_para, hits = clean_body_source_lines(docx_path)
        if n_para is None:
            add("E-5", False, "未找到成果 docx：成果文件/项目投资决策建议书.docx（先跑 06）")
        else:
            add("E-5", not hits,
                "正文 {} 段，来源行 0 处".format(n_para) if not hits
                else "正文仍有 {} 处来源行：{}".format(len(hits), "；".join(hits[:3])))

        # E-7 数值同源：正文每个数值都要在登记来源里找到；派生值必须可复算并留双上游
        number_problems, derived, _detail = check_numbers(root, rows, chapter_dir, {})
        docx_problems = check_docx_numbers(docx_path, rows, chapter_dir)
        all_number_problems = number_problems + docx_problems
        add("E-7", not all_number_problems,
            "{} 个数值块全部与来源同源（派生 {} 条已登记）".format(
                sum(1 for r in rows if r.get("numbers")), len(derived))
            if not all_number_problems
            else "{} 处数值无出处：{}".format(len(all_number_problems),
                                         "；".join(all_number_problems[:3])))
        if all_number_problems or link or source_bad:
            # E-2/E-6 失败同样落说明：只写依据核对.json 时，读者拿不到文件:行定位。
            write_failure_note(root, all_number_problems + source_bad, derived, link)
        else:
            stale = os.path.join(root, "过程文件", "07-回读自检", "校验未通过说明.md")
            if os.path.isfile(stale):
                os.remove(stale)     # 本轮已通过：上一轮的说明不能留在用户面前
        if derived:
            # 派生行回写对照表：登记行 + 双上游指针，供 xlsx/专家追问复核
            try:
                for row in rows:
                    row.pop("derived", None)
                for entry in derived:
                    for row in rows:
                        if (row.get("chapter"), row.get("section"), row.get("seq")) == (
                                entry["chapter"], entry["section"], entry["seq"]):
                            row.setdefault("derived", []).append(entry)
                            break
                data["rows"] = rows
                with open(register_path, "w", encoding="utf-8") as handle:
                    json.dump(data, handle, ensure_ascii=False, indent=1)
            except OSError:
                pass

        payload = {"ok": all(item["ok"] for item in checks),
                   "summary": {"passed": sum(1 for c in checks if c["ok"]), "total": len(checks)},
                   "register": register_path,
                   "problems": problems,
                   "numbers": {"problems": all_number_problems, "derived": derived},
                   "checks": checks}

    if args.format == "json":
        text = json.dumps(payload, ensure_ascii=False, indent=1)
        if args.out:
            os.makedirs(os.path.dirname(os.path.abspath(args.out)) or ".", exist_ok=True)
            with open(args.out, "w", encoding="utf-8") as handle:
                handle.write(text)
        print(text if not args.out else "[07] 依据核对 → {}".format(args.out))
    else:
        for item in payload["checks"]:
            print("  {} {}：{}".format("PASS" if item["ok"] else "FAIL", item["id"], item["detail"]))
        if payload.get("problems"):
            print("  未解析：")
            for problem in payload["problems"][:8]:
                print("   -", problem)
        print("[07] 依据核对 {}".format("通过" if payload["ok"] else "不通过"))
    return 0 if payload["ok"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
