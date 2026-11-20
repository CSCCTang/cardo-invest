#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""上游正文抽取（T-UPSTREAM）：把上游成果（现状分析报告）抽成结构化的
「上游摘要.json」——按 H1/H2/H3 分节保留原文段落与表格，并按关键词登记关键事实。

为什么要有这一步：
  05-章节成文的第 1—6 章要写项目自身的范围、现状、问题、范围界定、改造内容、组织运营。
  这些内容只能来自**上游成果**与本工作区填写的槽位。抽不出来就写「资料缺失」，
  绝不能用别的项目、也不能用本项目后续才产出的文件（如可研）来凑。

输入（只读，按优先级）：
  04_成果报告/*现状分析报告*.docx|.md      CardoStatus 成果
  输入文件/任意层级/*现状分析*.docx|.md    用户投递的同名成果
  05_过程/投资决策/槽位.json              功能定位/业态/改造思路/运营模式/组织架构/时序

输出（默认 05_过程/投资决策/上游摘要.json）：
  {schema, sources[[...]], sections[{id,title,level,path,paragraphs[],tables[]}],
   facts{key:{value,unit,name,evidence,source}}, slots{...}, gaps[{code,what,why,owner,where}]}

纪律：
  * 只读原件，不改、不移动、不重排；
  * 抽不到就进 gaps，不猜、不按经验补；
  * 每条事实带 evidence（命中原句）与 source（小节路径），可回溯。

用法：
  python extract_upstream.py --project-dir <项目目录> [--file <报告>] [--out <目录>]
退出码：0 抽取完成（可有 gaps）；1 原件不可读；2 项目目录无法识别。
"""
from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPT_DIR))

ENGINE_VERSION = "extract_upstream/1.0.0"
INVEST_REL = Path("05_过程") / "投资决策"
OUT_NAME = "上游摘要.json"
REPORT_PATTERNS = ("*现状分析报告*.docx", "*现状分析报告*.md", "*现状分析*.docx", "*现状分析*.md")

HEADING_RE = re.compile(r"^(?:Heading|标题)\s*([1-9])$", re.IGNORECASE)
MD_HEADING_RE = re.compile(r"^(#{1,6})\s+(.+?)\s*$")
NUM_PREFIX_RE = re.compile(r"^\s*(?:第?[一二三四五六七八九十百]+[章节]|[0-9]+(?:\.[0-9]+)*)[\s、.．]*")

# ---------------------------------------------------------------- 事实规则
# 只登记「原文能直接读到」的事实；推断值一律不做。
#
# 单位的取法：**跟着原文走**（正则里捕获单位，捕获不到再退回默认单位）。
# 早先单位写死在规则里、正则只匹配数字，于是「62680.5平方米」被标成
# 「62680.5 万平方米」——量纲放大一万倍，且报告里查不出来。
# 面积类尤其危险：62680.5 ㎡ 是合理面积，62680.5 万㎡ 相当于 668 平方公里。
#
# 数字正则一律要求**整数或小数点后至少一位**，且十进制点前不允许紧邻另一个数字，
# 免得从「744.0户」里截出「0 户」。
_NUM = r"(\d+(?:\.\d+)?)"
FACT_RULES = [
    ("scopeArea", "改造范围面积", "万平方米",
     ["改造范围", "更新范围", "改造总面积", "整治范围", "实施范围"],
     _NUM + r"\s*(万)?\s*(平方米|平米|㎡|m2|m²)"),
    ("publicSpaceArea", "公共空间面积", "万平方米",
     ["公共空间", "公共活动空间", "开敞空间"],
     _NUM + r"\s*(万)?\s*(平方米|平米|㎡|m2|m²)"),
    ("facadeArea", "外立面整治面积", "万平方米",
     ["外立面", "立面整治", "立面改造", "立面排危"],
     _NUM + r"\s*(万)?\s*(平方米|平米|㎡|m2|m²)"),
    ("roadLength", "配套道路长度", "公里",
     ["道路", "连接道路", "街巷长度"],
     _NUM + r"\s*(公里|km|千米)"),
    ("parkingSpots", "停车位数量", "个", ["停车位", "车位"], _NUM + r"\s*个"),
    ("chargers", "充电桩数量", "个", ["充电桩", "充电设施"], _NUM + r"\s*个"),
    ("adBoards", "广告位数量", "个", ["广告位", "店招", "广告牌"], _NUM + r"\s*个"),
    ("stalls", "摊位或餐车数量", "个", ["流动餐车", "摊位", "临时摊点"], _NUM + r"\s*个"),
    ("vendors", "沿街游商数量", "户", ["游商", "游贩", "占道经营"], _NUM + r"\s*户"),
    # 报告里唯一的人口推算口径：「11 个住宅小区合计 1832 户，按户均 3.2 人推算约 5862.4 人」。
    # 必须先于「万人」规则命中，且不能抓「户均 3.2 人」里的 3.2。
    ("population", "辐射人口", "人",
     ["人口规模", "辐射人口", "服务人口", "常住人口", "周边人口"],
     r"推算约\s*([0-9,]+(?:\.\d+)?)\s*人"),
    ("population", "辐射人口", "万人",
     ["辐射人口", "服务人口", "常住人口", "周边人口", "周边住户", "辐射住户",
      "辐射学校", "服务范围人口", "人口规模"],
     _NUM + r"\s*万人"),
    ("buildings", "建筑栋数", "栋", ["建筑", "房屋", "栋数"], _NUM + r"\s*栋"),
    ("density", "建筑密度", "%", ["建筑密度"], _NUM + r"\s*%"),
    # 现状图 3-1 的三项来源：绿地率、道路占比与建筑密度同段出现（实测样区口径），
    # 只认带 % 的相邻表述，避免抓到其他段落的无关比例。
    ("greeneryRate", "绿地率", "%", ["绿地率"], _NUM + r"\s*%"),
    ("roadRate", "道路占比", "%", ["道路占比"], _NUM + r"\s*%"),
    # 容积率必须带单位限定（容积率 2.5 / FAR 2.5），
    # 否则规则会从「道路网密度为6.3」这类句子里抓一个无关数字当容积率。
    ("floorAreaRatio", "容积率", "", ["容积率", "FAR"],
     r"(?:容积率|FAR)\s*(?:为|是|约|=|：|:)?\s*([0-9]+(?:\.[0-9]+)?)"),
    # 台账合计优先：「11 个住宅小区合计 1832 户」是片区总数，
    # 「其余 10 个小区累计 1088 户」只是局部，不能让后者覆盖前者。
    ("households", "住户数", "户", ["住宅小区", "合计"], r"合计\s*([0-9,]+)\s*户"),
    ("households", "住户数", "户", ["住户", "户数", "居民户"], r"([0-9,]+)\s*户"),
]

TEXT_RULES = [
    ("projectName", "项目名称", ["项目名称", "工程名称"]),
    ("siteBoundary", "四至范围", ["西至", "东至", "北至", "南至"]),
    ("projectNature", "项目性质", ["项目性质", "建设性质"]),
    ("owner", "建设单位", ["建设单位", "项目单位", "业主单位"]),
    ("direction", "项目方向", ["老旧街区", "老旧小区", "老旧厂区"]),
    ("landUse", "用地情况", ["新增建设用地", "用地性质", "土地征收"]),
]

GAP_DEFS = {
    "scopeArea": ("改造范围面积", "报告中未给出可识别的改造范围面积", "报告「现状详细分析/更新思路」章节"),
    "facadeArea": ("外立面整治面积", "报告中未给出可识别的立面面积", "报告「现状详细分析」章节"),
    "siteBoundary": ("四至范围", "报告中未写明四至或范围边界", "报告「片区基础概况」章节"),
    "owner": ("建设单位", "报告中未写明建设单位", "报告「项目概述」章节"),
}


def check_project_consistency(payload, project_dir):
    """校验上游现状报告与本项目是否**同一个项目**。

    背景（实测踩过）：工作区里放了别的项目的现状分析报告时，
    其 facts（项目名/方向/面积等）会被原样采信并写进正文，
    导致成品串项目（实测：老旧厂区的某企业饮品项目，
    却引用了老旧街区的某街某片区片区数据）。

    判据：把上游报告抽到的项目名与财务参数表里的项目名做**包含/相似**比对；
    方向不同也算不匹配（老旧街区 vs 老旧厂区 是硬冲突）。
    不匹配 → 登记缺项，**不删除事实**（事实本身对那个项目是对的），
    但在缺项表里明确告知，成文时不得引用。

    返回新增的 gaps 列表。
    """
    gaps = []
    facts = payload.get("facts") or {}
    upstream_name = (facts.get("projectName") or {}).get("value") or ""
    upstream_direction = (facts.get("direction") or {}).get("value") or ""
    if not upstream_name and not upstream_direction:
        return gaps

    params_path = project_dir / "05_过程" / "投资决策" / "财务输入.json"
    target_name = ""
    target_direction = ""
    try:
        import json
        params = json.loads(params_path.read_text(encoding="utf-8"))
        target = params.get("project") or {}
        target_name = str(target.get("name") or "")
        target_direction = str(target.get("direction") or "")
    except Exception:
        return gaps

    def normalize(text):
        return "".join(ch for ch in str(text) if ch not in " \t\n（）()·—-_")

    mismatch = []
    if upstream_direction and target_direction and upstream_direction != target_direction:
        mismatch.append("项目方向（上游 {} ≠ 本项目 {}）".format(upstream_direction, target_direction))
    if upstream_name and target_name:
        a, b = normalize(upstream_name), normalize(target_name)
        # 任一方包含另一方才算同项目；否则视为不同项目
        if not (a in b or b in a):
            mismatch.append("项目名称（上游「{}」≠ 本项目「{}」）".format(upstream_name, target_name))
    if mismatch:
        gaps.append({
            "code": "UPSTREAM-PROJECT-MISMATCH",
            "what": "上游现状分析报告与本项目不是同一个项目",
            "why": "；".join(mismatch) +
                   "。上游报告的事实只对该报告所属项目成立，不得用于本项目正文",
            "need": "提供与本项目对应的现状分析报告（或改名后重跑）",
            "owner": "user",
        })
    return gaps


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with open(path, "rb") as handle:
        for chunk in iter(lambda: handle.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def find_report(project: Path, explicit: str = "") -> list[Path]:
    if explicit:
        path = Path(explicit)
        if not path.is_absolute():
            path = project / path
        return [path] if path.is_file() else []
    hits: list[Path] = []
    seen: set[str] = set()
    bases = [project / "04_成果报告", project / "输入文件", project]
    for base in bases:
        if not base.is_dir():
            continue
        for pattern in REPORT_PATTERNS:
            for path in sorted(base.rglob(pattern)):
                if not path.is_file() or path.name.startswith("~$"):
                    continue
                key = str(path.resolve()).lower()
                if key in seen:
                    continue
                seen.add(key)
                hits.append(path)
    return hits


# ---------------------------------------------------------------- 解析

def _clean(text: str) -> str:
    return re.sub(r"[ \t\u3000]+", " ", (text or "").replace("\xa0", " ")).strip()


def _title(text: str) -> str:
    return NUM_PREFIX_RE.sub("", text).strip() or text


def parse_docx(path: Path) -> dict:
    import docx
    from docx.table import Table
    from docx.text.paragraph import Paragraph

    document = docx.Document(str(path))
    body = document.element.body
    sections: list[dict] = []
    stack: list[dict] = []
    current = {"id": "S0", "title": "（封面与前置部分）", "level": 0, "path": [],
               "paragraphs": [], "tables": []}
    sections.append(current)
    table_index = 0

    for child in body.iterchildren():
        tag = child.tag.split("}")[-1]
        if tag == "p":
            paragraph = Paragraph(child, document)
            text = _clean(paragraph.text)
            if not text:
                continue
            level = 0
            try:
                match = HEADING_RE.match((paragraph.style.name or "").strip())
                level = int(match.group(1)) if match else 0
            except AttributeError:
                level = 0
            if level:
                title = _title(text)
                while stack and stack[-1]["level"] >= level:
                    stack.pop()
                stack.append({"title": title, "level": level})
                current = {"id": "S{}".format(len(sections)), "title": title, "level": level,
                           "path": [item["title"] for item in stack], "paragraphs": [], "tables": []}
                sections.append(current)
            else:
                current["paragraphs"].append(text)
        elif tag == "tbl":
            table = Table(child, document)
            rows = []
            for row in table.rows:
                cells = [_clean(cell.text) for cell in row.cells]
                if any(cells):
                    rows.append(cells)
            if rows:
                table_index += 1
                current["tables"].append({"id": "TB{}".format(table_index),
                                          "caption": current["title"],
                                          "path": list(current["path"]) or [current["title"]],
                                          "rows": rows})
    return {"sections": sections, "tableCount": table_index}


def parse_markdown(path: Path) -> dict:
    text = path.read_text(encoding="utf-8", errors="replace")
    sections: list[dict] = []
    stack: list[dict] = []
    current = {"id": "S0", "title": "（封面与前置部分）", "level": 0, "path": [],
               "paragraphs": [], "tables": []}
    sections.append(current)
    table_index = 0
    table_buffer: list[list[str]] = []

    def flush_table():
        nonlocal table_index, table_buffer
        if len(table_buffer) >= 2:
            table_index += 1
            current["tables"].append({"id": "TB{}".format(table_index), "caption": current["title"],
                                      "path": list(current["path"]) or [current["title"]],
                                      "rows": [list(row) for row in table_buffer]})
        table_buffer = []

    for raw in text.splitlines():
        line = raw.strip()
        if line.startswith("|") and line.endswith("|"):
            cells = [_clean(cell) for cell in line.strip("|").split("|")]
            if all(re.fullmatch(r":?-{2,}:?", cell or "-") for cell in cells if cell):
                continue
            table_buffer.append(cells)
            continue
        if table_buffer:
            flush_table()
        if not line:
            continue
        match = MD_HEADING_RE.match(line)
        if match:
            level = len(match.group(1))
            title = _title(match.group(2))
            while stack and stack[-1]["level"] >= level:
                stack.pop()
            stack.append({"title": title, "level": level})
            current = {"id": "S{}".format(len(sections)), "title": title, "level": level,
                       "path": [item["title"] for item in stack], "paragraphs": [], "tables": []}
            sections.append(current)
            continue
        current["paragraphs"].append(_clean(line))
    if table_buffer:
        flush_table()
    return {"sections": sections, "tableCount": table_index}


def _sentences(text: str) -> list[str]:
    return [item.strip() for item in re.split(r"[。；\n]", text) if item.strip()]


def extract_facts(sections: list[dict]) -> tuple[dict, list[dict]]:
    pool: list[tuple[str, str]] = []
    for section in sections:
        label = " / ".join(section["path"]) or section["title"]
        for paragraph in section["paragraphs"]:
            for sentence in _sentences(paragraph):
                pool.append((sentence, label))
        for table in section["tables"]:
            for row in table["rows"]:
                for cell in row:
                    for sentence in _sentences(cell):
                        if len(sentence) >= 4:
                            pool.append((sentence, label))

    sentences = [item[0] for item in pool]
    label_of: dict[str, str] = {}
    for sentence, label in pool:
        label_of.setdefault(sentence, label)

    facts: dict[str, dict] = {}
    for key, name, default_unit, keywords, pattern in FACT_RULES:
        # 同一事实允许多条规则：先命中的优先（如「合计 1832 户」优先于局部「1088 户」）。
        if key in facts:
            continue
        hit = None
        for sentence in sentences:
            for keyword in keywords:
                pos = sentence.find(keyword)
                if pos < 0:
                    continue
                tail = sentence[pos:pos + 80]
                # 只在**关键词之后**的窗口里找数字。
                # 早先还兜底到整句搜（`or re.search(pattern, sentence)`），
                # 关键词定位失效时会把句中别处的数字算成该指标的事实——
                # 例如「占道经营」这句里别处的「30个」被当成游商数。
                match = re.search(pattern, tail)
                if not match:
                    continue
                raw = match.group(1).replace(",", "")
                try:
                    value = float(raw)
                except ValueError:
                    continue
                unit = default_unit
                # 单位**跟着原文走**：正则里若捕获了单位标记，就按它定单位与量纲。
                # 早先把单位写死（面积一律"万平方米"），而正则允许"万"字可选，
                # 于是「62680.5平方米」被标成 62680.5 万平方米——放大一万倍。
                groups = match.groups()
                if len(groups) >= 3:
                    wan_flag = groups[1]
                    unit_text = (groups[2] or "").strip()
                    if wan_flag:
                        # 原文带「万」→ 万单位；不带 → 原单位，数值原样保留
                        unit = "万" + (unit_text if unit_text in ("平方米", "平米", "㎡")
                                       else "平方米")
                    else:
                        if unit_text in ("㎡", "m2", "m²", "平米"):
                            unit = "平方米"
                        elif unit_text == "米":
                            unit = "米"
                        elif unit_text:
                            unit = unit_text
                        else:
                            unit = default_unit.replace("万", "") if default_unit.startswith("万") \
                                else default_unit
                hit = (value, sentence, keyword)
                break
            if hit:
                break
        if hit:
            facts[key] = {"value": hit[0], "unit": unit, "name": name,
                          "evidence": hit[1], "source": label_of.get(hit[1], "")}

    for key, name, keywords in TEXT_RULES:
        for sentence in sentences:
            matched = next((keyword for keyword in keywords if keyword in sentence), None)
            if not matched:
                continue
            value = sentence
            if key == "direction":
                value = matched
            elif key == "projectName":
                # 表格里的「项目名称」是栏目名，不是项目名。只认「项目名称：XXX」式取值；
                # 取不到就跳过，避免把栏目标签当成项目名，进而误判「上游报告不是本项目」。
                found = re.search(r"(?:项目名称|工程名称)(?:为|是|：|:)\s*([^。；\n]{4,60})", sentence)
                if not found:
                    continue
                value = found.group(1).strip()
            elif key == "projectNature":
                # 「项目性质为改建」→「改建」：只留性质本身，避免整句复述
                found = re.search(r"(?:项目性质|建设性质)(?:为|是|：|:)\s*([^，。；、]{1,8})", sentence)
                if found:
                    value = found.group(1).strip()
            elif key == "owner":
                found = re.search(r"(?:建设单位|项目单位|业主单位)(?:为|是|：|:)\s*([^，。；]{2,30})", sentence)
                if found:
                    value = found.group(1).strip()
            elif key == "siteBoundary":
                # 「四至范围：东起…」→「东起…」：剥掉引导前缀，
                # 否则正文会拼成「建设地点与范围：四至范围：东起…」（前缀重复）
                found = re.search(r"(?:四至范围|范围边界|用地范围)(?:为|是|：|:)\s*([^。；]{4,200})", sentence)
                if found:
                    value = found.group(1).strip()
            facts[key] = {"value": value, "unit": "", "name": name,
                          "evidence": sentence, "source": label_of.get(sentence, "")}
            break

    gaps = []
    for key, definition in GAP_DEFS.items():
        if not definition:
            continue
        if key in facts:
            continue
        name, why, where = definition
        gaps.append({"code": "UP-{}".format(key.upper()), "what": name,
                     "why": why, "owner": "user", "where": where})
    return facts, gaps


def read_slots(project: Path) -> tuple[dict, list[dict]]:
    path = project / INVEST_REL / "槽位.json"
    if not path.is_file():
        return {}, [{"code": "UP-SLOTS", "what": "槽位表",
                     "why": "05_过程/投资决策/槽位.json 不存在",
                     "owner": "user", "where": "输入文件/槽位.xlsx（功能定位/业态/改造思路/运营模式/组织架构/时序）"}]
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}, [{"code": "UP-SLOTS", "what": "槽位表",
                     "why": "槽位.json 无法解析", "owner": "user", "where": "输入文件/槽位.xlsx"}]
    inner = data.get("slots") if isinstance(data.get("slots"), dict) else data
    keep = ("functionalPosition", "businessTypes", "renovationIdea",
            "operationModel", "organization", "schedule")
    return {key: inner[key] for key in keep if inner.get(key)}, []


# ---------------------------------------------------------------- 主流程

def build(project: Path, explicit: str = "") -> dict:
    hits = find_report(project, explicit)
    slots, slot_gaps = read_slots(project)
    if not hits:
        gaps = [{"code": "UP-REPORT", "what": "现状分析报告",
                 "why": "04_成果报告 与 输入文件/ 下均无 *现状分析*.docx|.md",
                 "owner": "user", "where": "输入文件/现状分析报告.docx（CardoStatus 成果）"}] + slot_gaps
        return {"schema": "cardo/upstream-digest/v1", "engine": ENGINE_VERSION,
                "sources": [], "sections": [], "facts": {}, "slots": slots, "gaps": gaps}
    if len(hits) > 1:
        names = "、".join(path.name for path in hits)
        raise SystemExit("现状分析报告有多份（{}）：用 --file 指定本次要用的那一份".format(names))

    path = hits[0]
    parsed = parse_markdown(path) if path.suffix.lower() == ".md" else parse_docx(path)
    facts, gaps = extract_facts(parsed["sections"])
    rel = path.name
    try:
        rel = path.resolve().relative_to(project.resolve()).as_posix()
    except ValueError:
        pass
    return {
        "schema": "cardo/upstream-digest/v1",
        "engine": ENGINE_VERSION,
        "sources": [{"file": rel, "sha256": _sha256(path),
                     "sections": len(parsed["sections"]), "tables": parsed["tableCount"]}],
        "sections": parsed["sections"],
        "facts": facts,
        "slots": slots,
        "gaps": gaps + slot_gaps,
    }


def write(payload: dict, out_dir: Path) -> Path:
    out_dir.mkdir(parents=True, exist_ok=True)
    target = out_dir / OUT_NAME
    with open(target, "w", encoding="utf-8") as handle:
        json.dump(payload, handle, ensure_ascii=False, indent=2)
        handle.write("\n")
    return target


def rel_path(path: Path, base: Path) -> str:
    """路径 → 相对 base 的 posix 串；越界时回退为绝对路径。

    stdout 会被总控原样写进 运行日志.json，08 步按「过程产物不得含盘符绝对路径」
    扫描该日志；这里统一输出相对路径，避免流水线把自己的日志判成违规。
    """
    try:
        return Path(path).resolve().relative_to(Path(base).resolve()).as_posix()
    except ValueError:
        return str(path)


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="上游成果抽取 → 上游摘要.json")
    parser.add_argument("--project-dir", default=".", help="项目根目录")
    parser.add_argument("--file", default="", help="指定现状分析报告（相对项目根或绝对路径）")
    parser.add_argument("--out", default="", help="输出目录（缺省 05_过程/投资决策）")
    args = parser.parse_args(argv)

    project = Path(args.project_dir).resolve()
    if not project.is_dir():
        print("[错误] 项目目录不存在：{}".format(project))
        return 2
    out_dir = Path(args.out) if args.out else project / INVEST_REL
    payload = build(project, args.file)
    # 项目一致性校验：上游现状报告是否就是本项目（防止串项目污染正文）
    mismatch_gaps = check_project_consistency(payload, project)
    if mismatch_gaps:
        payload["gaps"] = list(payload.get("gaps") or []) + mismatch_gaps
        payload["projectMismatch"] = True
    target = write(payload, out_dir)

    if not payload["sources"]:
        print("[上游] 未找到现状分析报告，已登记缺口 UP-REPORT")
        print("[上游] 落位：输入文件/现状分析报告.docx（CardoStatus 成果）")
    else:
        source = payload["sources"][0]
        print("[上游] 原件：{}（{} 节 / {} 张表）".format(
            source["file"], source["sections"], source["tables"]))
    print("[上游] 事实 {} 条；槽位 {} 项；缺口 {} 条".format(
        len(payload["facts"]), len(payload["slots"]), len(payload["gaps"])))
    for gap in payload["gaps"]:
        print("  ! {} {}".format(gap["code"], gap["what"]))
    print("[上游] 产物：{}".format(rel_path(target, project)))
    return 0


if __name__ == "__main__":
    sys.exit(main())
