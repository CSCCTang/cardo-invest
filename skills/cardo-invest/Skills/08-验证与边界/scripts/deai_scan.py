#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""中文汇报文稿去 AI 味：指标扫描 + 命中定位。

用法：
    python deai_scan.py <文件或目录> [--slides 2,3,5] [--genre deck|doc]
                        [--json 结果.json] [--md 报告.md] [--quiet]

默认 --genre doc（书面材料）；演示稿用 --genre deck（阈值放宽、按页统计）。
退出码：0 达标，1 未达标，2 输入错误。

口径见 ../references/02_四项指标.md，规则见 ../references/01_禁用与慎用清单.md。
"""
from __future__ import annotations

import argparse
import json
import math
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path

CJK = re.compile(r"[\u4e00-\u9fff]")
NUM = re.compile(r"\d+(?:[.,]\d+)?\s*%?")
CN_QTY = re.compile(r"(?<![下进同逐第])[一二三四五六七八九十百千两]+(页|表|图|项|条|人月|万元|秒|分钟|版|件|段|步)")
# 高频实义名词：2 字词没法靠 n-gram 抓，用受控表计数，可按项目增补
TERM_WATCH = "问题|工作|管理|项目|能力|水平|机制|体系|建设|发展|要求|标准|目标|任务|举措|支撑|保障|基础|作用|意义|质量|效率|服务|资源|数据|指标|流程|方法|模式"
# 实体名词表：可指认的人、产物、制度物。项目换了就增补，不在这里堆形容词。
ENTITY = re.compile(
    r"项目负责人|资料员|写手|专家|财务|甲方|业主|评委|评审|验收|部门|"
    r"报告|建议书|清单|对照表|表格|脚本|指标|标准|法规|文号|条款|图纸|"
    r"栋数|户数|密度|面积|单价|人日|人月|样例|模型|适配层|检查项|口径|"
    r"材料|电话|文件夹|会议|会上|合同|预算|工期|现场|工日")
PAGE_NO = re.compile(r"^\d{1,2}\s*[/／]\s*\d{1,2}$")
SENT_SPLIT = re.compile(r"[。！？；!?;\n]+")
LABEL = re.compile(r"^([\u4e00-\u9fff]{1,8})[：:]")
LEAD_NUM = re.compile(r"^(?:[（(]?[一二三四五六七八九十\d]+[）).、]\s*|[-·•*]\s*)+")
MD_NOISE = re.compile(r"^[|#>*_\-\s]+|[|*`]+")
# 专名与书名号：法规名、项目名里带“提升/推进”，不算套话
PROPER = re.compile(r"《[^》]{1,80}》|[\u4e00-\u9fff]{6,}?项目|[\u4e00-\u9fff]{4,}?(报告|建议书|作业标准)")
ABSTRACT_ITEM = re.compile(
    r"(性|化|力|度|面|型|位)$|赋能|抓手|闭环|生态|格局|价值|体系化|"
    r"管理|机制|体系|能力|水平|建设|工作|流程|标准|协同|融合|"
    r"思想|认识|方向|共识|精神|意识|理念|氛围|合力|实效|局面|贡献|"
    r"支撑|保障|基础|举措|任务|目标|要求|责任|资源|质量|效率")

# ---------------------------------------------------------------- 规则表
# 必删：(编号, 名称, 正则, 补充判定 doc|None, 改法)
# 补充判定 cjk_digits：句中无数字才命中（程度词代替结论类）
BIDAO = [
    ("B1", "假连接词/总结腔",
     r"值得注意的是|综上所述|总而言之|由此可见|不难看出|不难发现|众所周知|毋庸置疑|需要指出的是|换言之|一言以蔽之|总的来说|主要体现在以下|具体而言|从整体上看",
     None, "整句删；要过渡就直接说上一句的结论。"),
    ("B2", "空概念名词",
     r"赋能|抓手|闭环|生态位|护城河|底层逻辑|组合拳|新范式|颗粒度|全链路|业务链路|价值链路",
     None, "换成动作+对象，例如“每次改数后重算一遍”。"),
    ("B3", "无主语宏大句",
     r"^(随着|当前|在.{0,14}(背景|形势|环境|要求|趋势)下|当前形势|近年来|面对新|围绕|聚焦|以.{1,14}为(抓手|牵引|统领|引领))",
     "no_digit", "补主语、补时点、补数字；补不出就删。"),
    ("B4", "三项以上同构并列", r"(?!)", None, "留最硬的一项，其余改成分句或删。"),
    ("B5", "模板骨架外露",
     r"^(?:[（(]?[一二三四五六七八九十\d]+[）).、]\s*)(?:现状|问题|原因|对策|表现|证据|落点|成因|影响|对应|结论|建议|举措)",
     "repeat2", "提纲不进正文；改成“判断—依据—影响”自然段。"),
    ("B6", "同义堆叠",
     r"(提升|优化|加强|推进|完善|强化)[和与及、](提升|优化|加强|推进|完善|强化)",
     None, "删一个，另一个换成具体动作。"),
    ("B7", "程度副词代替结论",
     r"(显著|大幅|明显|有效|极大)(提升|提高|改善|增强|降低|缩短|减少)",
     "no_digit", "给出改前改后的数，或整句删。"),
    ("B8", "万能动词+抽象宾语",
     r"(加强|推进|完善|提升|强化|优化)[^，。；：！？、]{0,4}(管理|机制|水平|能力|建设|工作|体系)",
     None, "写清改哪一处、谁负责、从几到几。"),
    ("B9", "排比式空转",
     r"既要.{0,18}又要.{0,18}还要",
     None, "只留一条能验收的。"),
    ("B10", "口号当结论",
     r"[“\"](高质量发展|数字化转型|降本增效|提质增效|跨越式发展)[”\"]",
     None, "引号里的词不算结论，换成判断句。"),
    ("B11", "自证式空句",
     r"(具有|发挥).{0,6}(意义|作用)|提供.{0,8}(有力)?支撑|作出.{0,6}贡献",
     None, "删；用一条可核事实替代。"),
    ("B12", "空转过渡段",
     r"^(如前所述|上文提到|接下来我们|从这个角度看|回到刚才|前文已述)",
     None, "删，段落直接接上一条事实。"),
]

# 慎用：(编号, 名称, 正则, 换成什么)
SHENYONG = [
    ("C1", "进一步/持续/不断", r"进一步|持续|不断|一如既往", "写次数或时点，如“第二轮再跑一次”。"),
    ("C2", "全面/系统/深入", r"全面|系统性|深入|全方位|多维度", "写覆盖哪几项、几页、几条。"),
    ("C3", "有效/显著/明显", r"有效(?!期|性)|显著|明显|大幅", "换数字或删。"),
    ("C4", "切实/扎实/有力", r"切实|扎实|有力|稳步", "写验收口径：谁在哪天看什么。"),
    ("C5", "高度重视/大力支持", r"高度重视|大力支持|全力推动", "写谁在什么会上定了什么。"),
    ("C6", "优化/完善/提升", r"(优化|完善|提升|改进)[^，。；]{0,3}(水平|能力|机制|体系|工作|管理|质量|效率|流程|标准|建设|资源|配置)", "写从几到几、改哪一处。"),
    ("C7", "打造/构建/塑造", r"打造|构建|塑造|锻造", "写产物名：报告、清单、脚本。"),
    ("C8", "深度融合/协同联动", r"深度融合|协同联动|协同发力", "写接口、交接物、频次。"),
    ("C9", "以…为…", r"以.{1,12}为(抓手|中心|核心|重点|导向|契机)", "写执行主体+动作。"),
    ("C10", "模糊限定", r"一定程度上|相对而言|较为|有所|某种意义上", "换区间或删。"),
]

EMPTY_WORDS = re.compile(r"赋能|抓手|闭环|生态|护城河|底层逻辑|组合拳|新范式|颗粒度|链路|格局|价值|体系化")

THRESHOLDS = {
    "doc": {"mean": 32, "std": 8, "max": 60, "long_ratio": 0.10, "runs": 2,
            "rep": 0.10, "num": 3, "term": 4, "concrete": 0.80, "empty": 0.10},
    "deck": {"mean": 26, "std": 6, "max": 45, "long_ratio": 0.05, "runs": 4,
             "rep": 0.20, "num": 3, "term": 4, "concrete": 0.70, "empty": 0.10},
}


@dataclass
class Unit:
    name: str
    paras: list[str] = field(default_factory=list)
    title: str = ""


# ---------------------------------------------------------------- 读文件
def clean_para(text: str) -> str:
    text = text.replace("\u3000", " ").replace("\xa0", " ")
    text = MD_NOISE.sub("", text.strip())
    return re.sub(r"\s+", " ", text).strip()


def load_file(path: Path) -> list[Unit]:
    suffix = path.suffix.lower()
    if suffix == ".pptx":
        return load_pptx(path)
    if suffix == ".docx":
        return load_docx(path)
    if suffix in (".md", ".txt"):
        return load_text(path)
    raise ValueError(f"不支持的文件类型：{path.suffix}（只收 .md/.docx/.pptx）")


def load_pptx(path: Path) -> list[Unit]:
    from pptx import Presentation

    units: list[Unit] = []
    for idx, slide in enumerate(Presentation(str(path)).slides, 1):
        unit = Unit(name=f"第 {idx} 页")
        for shape in slide.shapes:
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    text = clean_para("".join(r.text for r in para.runs))
                    if not text or PAGE_NO.match(text):
                        continue
                    if not unit.title and CJK.search(text) and len(text) <= 40:
                        unit.title = text
                        continue
                    unit.paras.append(text)
            if getattr(shape, "has_table", False) and shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        for line in cell.text.split("\n"):
                            text = clean_para(line)
                            if text and not PAGE_NO.match(text):
                                unit.paras.append(text)
        units.append(unit)
    return units


def load_docx(path: Path) -> list[Unit]:
    from docx import Document

    units: list[Unit] = [Unit(name="正文")]
    for para in Document(str(path)).paragraphs:
        text = clean_para(para.text)
        if not text:
            continue
        style = (para.style.name or "").lower()
        if "heading" in style or "标题" in style or "title" in style:
            units.append(Unit(name=text[:30], title=text))
            continue
        units[-1].paras.append(text)
    return [u for u in units if u.paras]


def load_text(path: Path) -> list[Unit]:
    units: list[Unit] = [Unit(name="正文")]
    in_code = False
    for raw in path.read_text(encoding="utf-8", errors="ignore").splitlines():
        line = raw.strip()
        if line.startswith("```"):
            in_code = not in_code
            continue
        if in_code or not line:
            continue
        if line.startswith("## ") or (line.startswith("# ") and len(units) == 1 and not units[0].paras):
            name = line.lstrip("#").strip()
            units.append(Unit(name=name[:30], title=name))
            continue
        if re.match(r"^\|[-: |]+\|$", line):
            continue
        text = clean_para(line)
        if text:
            units[-1].paras.append(text)
    return [u for u in units if u.paras]


def chunk_units(units: list[Unit], max_chars: int = 1000) -> list[Unit]:
    """docx/md 没有页码，按约 1000 汉字切成“节·k”，让单页口径可算。"""
    out: list[Unit] = []
    for u in units:
        if sum(cjk_len(p) for p in u.paras) <= max_chars:
            out.append(u)
            continue
        buf: list[str] = []
        size = 0
        part = 1
        for p in u.paras:
            buf.append(p)
            size += cjk_len(p)
            if size >= max_chars:
                out.append(Unit(f"{u.name}·{part}", buf, u.title))
                buf, size, part = [], 0, part + 1
        if buf:
            out.append(Unit(f"{u.name}·{part}", buf, u.title))
    return out


# ---------------------------------------------------------------- 指标
def split_sents(text: str) -> list[str]:
    return [s.strip() for s in SENT_SPLIT.split(text) if CJK.search(s)]


def cjk_len(text: str) -> int:
    return len(CJK.findall(text))


def sentence_stats(texts: list[str]) -> dict:
    lens = [cjk_len(s) for t in texts for s in split_sents(t)]
    lens = [n for n in lens if n > 0]
    if not lens:
        return {"count": 0, "mean": 0.0, "std": 0.0, "max": 0, "long_ratio": 0.0, "runs": 0}
    mean = sum(lens) / len(lens)
    std = math.sqrt(sum((n - mean) ** 2 for n in lens) / len(lens))
    long_ratio = sum(1 for n in lens if n > 50) / len(lens)
    runs, i = 0, 0
    while i + 2 < len(lens):
        window = lens[i:i + 3]
        if min(window) >= 8 and max(window) - min(window) <= 3:
            runs += 1
            i += 3
        else:
            i += 1
    return {"count": len(lens), "mean": round(mean, 1), "std": round(std, 1),
            "max": max(lens), "long_ratio": round(long_ratio, 3), "runs": runs}


def para_key(para: str) -> str:
    label = LABEL.match(para)
    if label:
        return "标签:" + label.group(1)
    text = LEAD_NUM.sub("", para)
    chars = CJK.findall(text)
    return "".join(chars[:6])


def is_label_line(para: str) -> bool:
    """页眉、表头、卡片小标题：不算正文段落，不参与段均与具体度。"""
    return cjk_len(para) <= 6 and not re.search(r"[。！？；：]", para) and not NUM.search(para)


def body_paras(paras: list[str]) -> list[str]:
    return [p for p in paras if not is_label_line(p)]


def first_sentence_shape(paras: list[str], min_repeat: int = 2) -> dict:
    paras = body_paras(paras)
    keys = [para_key(p) for p in paras]
    freq: dict[str, int] = {}
    for k in keys:
        freq[k] = freq.get(k, 0) + 1
    repeat = {k: v for k, v in freq.items() if v >= min_repeat and len(k) >= min_repeat}
    repeated_paras = sum(repeat.values())
    # 首句开头 4 字，捕捉“判断—依据—”这类换词不换骨架
    heads = ["".join(CJK.findall(split_sents(p)[0])[:4]) for p in paras if split_sents(p)]
    head_freq: dict[str, int] = {}
    for h in heads:
        head_freq[h] = head_freq.get(h, 0) + 1
    head_repeat = {h: v for h, v in head_freq.items() if v >= min_repeat and len(h) >= 3}
    return {"total": len(paras), "repeat": repeat,
            "rate": round(repeated_paras / len(paras), 3) if paras else 0.0,
            "head_repeat": head_repeat}


def value_repeat(texts: list[str], title: str) -> tuple[dict, list[tuple[str, int]]]:
    nums: dict[str, int] = {}
    for t in texts:
        for m in re.finditer(r"\d+(?:[.,]\d+)?\s*%?", t):
            if t[m.end():m.end() + 1] in "。、）)．":  # 编号不是数值
                continue
            key = m.group(0).replace(" ", "").replace(",", "")
            if re.fullmatch(r"(19|20)\d{2}", key):  # 年份不算“同一个数值反复说”
                continue
            nums[key] = nums.get(key, 0) + 1
    # 用“，”隔开每段再取 n-gram，避免跨段拼出假词；标签前缀不计入
    corpus = "，".join(re.sub(r"^([\u4e00-\u9fff]{1,8})[：:]", "", p)
                       for p in body_paras(texts))
    terms: dict[str, int] = {}
    # 以“，”为界分段统计，不跨段
    segments = [s for s in re.split(r"[，、。；：！？]", corpus) if s]
    for seg in segments:
        chars = "".join(CJK.findall(seg))
        for size in (6, 5, 4, 3):
            for i in range(len(chars) - size + 1):
                term = chars[i:i + size]
                if term in title or term[0] in "的了是在和与或就都也很把被对从为以而及":
                    continue
                terms[term] = terms.get(term, 0) + 1
    picked: list[tuple[str, int]] = []
    for term, cnt in sorted(terms.items(), key=lambda kv: (-kv[1], -len(kv[0]))):
        if cnt < 4:
            break
        if any(term in p or p in term for p, _ in picked):
            continue
        picked.append((term, cnt))
        if len(picked) >= 3:
            break
    # 2 字高频实义名词：受控表计数，与 n-gram 结果合并去重
    for word in TERM_WATCH.split("|"):
        cnt = corpus.count(word)
        if cnt >= 4 and not any(word in p or p in word for p, _ in picked):
            picked.append((word, cnt))
    picked.sort(key=lambda kv: (-kv[1], -len(kv[0])))
    picked = picked[:3]
    max_num = max(nums.values()) if nums else 0
    return {"nums": nums, "max_num": max_num, "terms": picked,
            "max_term": max((c for _, c in picked), default=0)}, picked


def concreteness(paras: list[str]) -> dict:
    paras = body_paras(paras)
    concrete = empty = 0
    for p in paras:
        if NUM.search(p) or CN_QTY.search(p) or ENTITY.search(p):
            concrete += 1
        elif EMPTY_WORDS.search(p):
            empty += 1
    total = len(paras) or 1
    return {"total": len(paras), "concrete": concrete,
            "concrete_rate": round(concrete / total, 3),
            "empty": empty, "empty_rate": round(empty / total, 3)}


def b4_hits(para: str) -> list[str]:
    """三项以上、等长（3 或 4 字）且至少两项是抽象词的并列短语。"""
    found: list[str] = []
    for clause in re.split(r"[，。；：！？]", para):
        items = [re.sub(r"[^\u4e00-\u9fff]", "", x) for x in clause.split("、")]
        i = 0
        while i < len(items):
            size = len(items[i])
            if size not in (3, 4):
                i += 1
                continue
            j = i
            while j < len(items) and len(items[j]) == size:
                j += 1
            run = items[i:j]
            if len(run) >= 3 and sum(1 for it in run if ABSTRACT_ITEM.search(it)) >= 2:
                found.append("、".join(run))
            i = j
    return found


def rule_hits(unit: Unit) -> list[dict]:
    hits: list[dict] = []
    numbered = 0
    for no, para in enumerate(unit.paras, 1):
        masked = PROPER.sub("■", para)
        for rid, name, pattern, cond, fix in BIDAO:
            for m in re.finditer(pattern, masked):
                if cond == "no_digit" and NUM.search(para):
                    continue
                if rid == "B5":
                    continue
                hits.append({"unit": unit.name, "para": no, "rule": rid, "name": name,
                             "frag": m.group(0)[:40], "text": para[:80], "fix": fix})
        for frag in b4_hits(para)[:1]:
            hits.append({"unit": unit.name, "para": no, "rule": "B4", "name": "三项以上同构并列",
                         "frag": frag[:40], "text": para[:80],
                         "fix": "留最硬的一项，其余改成分句或删。"})
        if re.match(r"^[（(]?[一二三四五六七八九十\d]+[）).、]\s*(现状|问题|原因|对策|表现|证据|落点|成因|影响|对应|结论|建议|举措)", para):
            numbered += 1
        for rid, name, pattern, fix in SHENYONG:
            for m in re.finditer(pattern, masked):
                hits.append({"unit": unit.name, "para": no, "rule": rid, "name": name,
                             "frag": m.group(0)[:40], "text": para[:80], "fix": fix})
    if numbered >= 2:
        no, para = next((i, p) for i, p in enumerate(unit.paras, 1)
                        if re.match(r"^[（(]?[一二三四五六七八九十\d]+[）).、]", p))
        hits.append({"unit": unit.name, "para": no, "rule": "B5", "name": "模板骨架外露",
                     "frag": f"编号骨架 {numbered} 段", "text": para[:80],
                     "fix": "提纲不进正文；改成“判断—依据—影响”自然段。"})
    order = {r[0]: i for i, r in enumerate(BIDAO + SHENYONG)}
    return sorted(hits, key=lambda h: (h["para"], order.get(h["rule"], 99)))


# ---------------------------------------------------------------- 打分
def score_m1(st: dict, th: dict) -> int:
    s = 20.0
    if st["count"] < 3:
        return 20
    if st["mean"] > th["mean"]:
        s -= min(8, 3 * math.ceil((st["mean"] - th["mean"]) / 3))
    if st["std"] < th["std"]:
        s -= 5
    if st["max"] > th["max"]:
        s -= min(5, 2 * math.ceil((st["max"] - th["max"]) / 10))
    if st["long_ratio"] > th["long_ratio"]:
        s -= min(4, 2 * math.ceil((st["long_ratio"] - th["long_ratio"]) * 100 / 5))
    allowance = th["runs"]
    if st["runs"] > allowance:
        s -= min(3, math.ceil(st["runs"] - allowance))
    return int(max(0, round(s)))


def score_m2(shape: dict, th: dict) -> int:
    rate = shape["rate"]
    if rate <= th["rep"]:
        s = 20
    elif rate <= th["rep"] * 2:
        s = 14
    elif rate <= th["rep"] * 3:
        s = 8
    else:
        s = 3
    if shape["head_repeat"]:
        s -= min(6, 3 * len(shape["head_repeat"]))
    return max(0, s)


def score_m3(val: dict, th: dict) -> int:
    num, term = val["max_num"], val["max_term"]
    if num <= th["num"] and term <= th["term"]:
        return 20
    if num <= 5 and term <= 6:
        return 12
    if num <= 8 and term <= 9:
        return 6
    return 2


def score_m4(conc: dict, th: dict) -> int:
    rate = conc["concrete_rate"]
    if rate >= th["concrete"]:
        s = 20
    elif rate >= th["concrete"] - 0.10:
        s = 16
    elif rate >= th["concrete"] - 0.20:
        s = 10
    else:
        s = 4
    if conc["empty_rate"] > th["empty"]:
        s -= min(10, 3 * math.ceil((conc["empty_rate"] - th["empty"]) * 100 / 5))
    return max(0, s)


def grade(total: int, bidao: int) -> str:
    if bidao >= 3 or total < 70:
        return "C"
    if total >= 85 and bidao == 0:
        return "A"
    return "B"


# ---------------------------------------------------------------- 主流程
def scan(units: list[Unit], genre: str, slides: str | None) -> dict:
    th = THRESHOLDS[genre]
    if slides:
        want = {int(x) for x in re.split(r"[,，\s]+", slides) if x.strip().isdigit()}
        units = [u for i, u in enumerate(units, 1) if i in want] or units
    all_paras = body_paras([p for u in units for p in u.paras])
    stats = sentence_stats(all_paras)
    min_repeat = 5 if genre == "deck" else 2
    shape = first_sentence_shape(all_paras, min_repeat)
    val, terms = value_repeat(all_paras, " ".join(u.title for u in units))
    conc = concreteness(all_paras)
    hits = [h for u in units for h in rule_hits(u)]
    bidao = [h for h in hits if h["rule"].startswith("B")]
    shen = [h for h in hits if h["rule"].startswith("C")]

    gate = 20 if not bidao else 10 if len(bidao) == 1 else 5 if len(bidao) == 2 else 0

    per_unit = []
    for u in units:
        per_unit.append({
            "unit": u.name,
            "句长": sentence_stats(body_paras(u.paras)),
            "首句重复率": first_sentence_shape(u.paras, min_repeat)["rate"],
            "单值最大次数": value_repeat(u.paras, u.title)[0]["max_num"],
            "名词最大次数": value_repeat(u.paras, u.title)[0]["max_term"],
            "具体段率": concreteness(u.paras)["concrete_rate"],
            "必删": sum(1 for h in rule_hits(u) if h["rule"].startswith("B")),
            "慎用": sum(1 for h in rule_hits(u) if h["rule"].startswith("C")),
        })

    # M2/M3 按“单页”判：取最差一页，不拿全书平均把坏页洗掉
    worst_rep = max((u["首句重复率"] for u in per_unit), default=0.0)
    worst_num = max((u["单值最大次数"] for u in per_unit), default=0)
    worst_term = max((u["名词最大次数"] for u in per_unit), default=0)
    parts = {"必删闸门": gate, "M1 句长": score_m1(stats, th),
             "M2 首句重复": score_m2({"rate": worst_rep, "head_repeat": shape["head_repeat"]}, th),
             "M3 单值重复": score_m3({"max_num": worst_num, "max_term": worst_term}, th),
             "M4 具体度": score_m4(conc, th)}
    total = sum(parts.values())

    return {"file": "", "genre": genre, "thresholds": th, "units": [u.name for u in units],
            "stats": stats, "shape": shape,
            "max_num": val["max_num"], "terms": terms,
            "worst_rep": worst_rep, "worst_num": worst_num, "worst_term": worst_term,
            "concrete": conc, "parts": parts, "total": total,
            "grade": grade(total, len(bidao)),
            "bidao_count": len(bidao), "shen_count": len(shen),
            "hits": hits, "per_unit": per_unit}


def render_md(res: dict) -> str:
    th = res["thresholds"]
    lines = [f"# 去 AI 味扫描：{res['file']}", "",
             f"口径：{res['genre']}｜抽样：{'、'.join(res['units'])}", "",
             "## 一、评分", "",
             "| 项 | 得分 | 满分 |", "|---|---:|---:|"]
    for k, v in res["parts"].items():
        lines.append(f"| {k} | {v} | 20 |")
    lines += [f"| **合计** | **{res['total']}** | **100** |", "",
              f"等级 **{res['grade']}**；必删 {res['bidao_count']} 处，慎用 {res['shen_count']} 处。", "",
              "## 二、四项指标", "",
              "| 指标 | 实测 | 阈值 |", "|---|---:|---|",
              f"| 平均句长 | {res['stats']['mean']} 字 | ≤{th['mean']} |",
              f"| 句长标准差 | {res['stats']['std']} | ≥{th['std']} |",
              f"| 最长句 | {res['stats']['max']} 字 | ≤{th['max']} |",
              f"| 长句占比（>50 字） | {res['stats']['long_ratio']:.1%} | ≤{th['long_ratio']:.0%} |",
              f"| 连续同长（3 句） | {res['stats']['runs']} 次 | ≤{th['runs']} |",
              f"| 段首形态重复率 | 全篇 {res['shape']['rate']:.1%}／最差页 {res['worst_rep']:.1%} | ≤{th['rep']:.0%} |",
              f"| 同一数值最多出现（单页） | {res['worst_num']} 次 | ≤{th['num']} |",
              f"| 同一名词最多出现（单页） | {res['worst_term']} 次 | ≤{th['term']} |",
              f"| 具体段占比 | {res['concrete']['concrete_rate']:.1%} | ≥{th['concrete']:.0%} |",
              f"| 空段占比 | {res['concrete']['empty_rate']:.1%} | ≤{th['empty']:.0%} |", "",
              "## 三、逐页", "",
              "| 页/节 | 平均句长 | 最长句 | 首句重复率 | 数值峰值 | 名词峰值 | 具体段率 | 必删 | 慎用 |",
              "|---|---:|---:|---:|---:|---:|---:|---:|---:|"]
    for u in res["per_unit"]:
        s = u["句长"]
        lines.append(f"| {u['unit']} | {s['mean']} | {s['max']} | {u['首句重复率']:.0%} | "
                     f"{u['单值最大次数']} | {u['名词最大次数']} | {u['具体段率']:.0%} | "
                     f"{u['必删']} | {u['慎用']} |")
    low = [u for u in res["per_unit"] if u["具体段率"] < th["concrete"]]
    if low:
        lines += ["", "具体段低于阈值的页：",
                  "、".join(f"{u['unit']}（{u['具体段率']:.0%}）" for u in low)]
    lines += ["", "## 四、命中明细", ""]
    if not res["hits"]:
        lines.append("无命中。")
    else:
        lines += ["| 页/节 | 段 | 规则 | 命中 | 原文（前 80 字） | 改法 |",
                  "|---|---:|---|---|---|---|"]
        for h in res["hits"]:
            text = h["text"].replace("|", "／")
            lines.append(f"| {h['unit']} | {h['para']} | {h['rule']} {h['name']} | "
                         f"{h['frag'].replace('|', '／')} | {text} | {h['fix']} |")
    return "\n".join(lines) + "\n"


def main() -> int:
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8")

    ap = argparse.ArgumentParser(description="中文汇报文稿去 AI 味扫描")
    ap.add_argument("path", help="文件（.md/.docx/.pptx）或目录")
    ap.add_argument("--slides", help="只查指定页/节，如 2,3,5")
    ap.add_argument("--genre", choices=["doc", "deck"], default="doc")
    ap.add_argument("--json", dest="json_out")
    ap.add_argument("--md", dest="md_out")
    ap.add_argument("--quiet", action="store_true")
    args = ap.parse_args()

    target = Path(args.path)
    if not target.exists():
        print(f"找不到：{target}", file=sys.stderr)
        return 2
    files = [target] if target.is_file() else sorted(
        p for p in target.rglob("*") if p.suffix.lower() in {".md", ".docx", ".pptx"})
    if not files:
        print("目录内没有 .md/.docx/.pptx。", file=sys.stderr)
        return 2

    worst = 0
    for path in files:
        try:
            units = load_file(path)
            if path.suffix.lower() != ".pptx":
                units = chunk_units(units)
        except Exception as exc:  # 读不了的件跳过，不猜内容
            print(f"跳过 {path.name}：{exc}", file=sys.stderr)
            continue
        res = scan(units, args.genre, args.slides)
        res["file"] = str(path)
        worst = max(worst, 1 if res["grade"] == "C" else 0)
        if args.json_out:
            out = Path(args.json_out) if len(files) == 1 else Path(args.json_out).with_name(
                Path(args.json_out).stem + "_" + path.stem + ".json")
            out.parent.mkdir(parents=True, exist_ok=True)
            out.write_text(json.dumps(res, ensure_ascii=False, indent=2), encoding="utf-8")
        if args.md_out:
            out = Path(args.md_out) if len(files) == 1 else Path(args.md_out).with_name(
                Path(args.md_out).stem + "_" + path.stem + ".md")
            out.parent.mkdir(parents=True, exist_ok=True)
            out.write_text(render_md(res), encoding="utf-8")
        if not args.quiet:
            print(render_md(res))
        else:
            print(f"{path.name}：{res['total']} 分 {res['grade']}｜必删 {res['bidao_count']}｜"
                  f"慎用 {res['shen_count']}｜具体段 {res['concrete']['concrete_rate']:.0%}")
    return worst


if __name__ == "__main__":
    raise SystemExit(main())
