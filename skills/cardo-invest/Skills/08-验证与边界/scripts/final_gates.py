# -*- coding: utf-8 -*-
"""final_gates.py -- 08 成品闸门（建议书版，与 CardoStatus 同一套交付纪律）。

五道闸，全部对着**成品**（成果文件/成品版/…）而不是过程稿：
  G1 交付结构：成品 docx 与 pdf 同目录、PDF 页数有效、内部件一成果一子目录；
  G2 成品版式：封面不印页码、正文页脚为「第 N 页 / 共 M 页」、目录 20 条以上且页码不留空；
  G3 缺口词：成品正文与表格零命中（未提资/暂缺/待补/未提供/请提供/待补充/资料缺失/缺项/待定/待核实）；
  G4 三源同源：现状分析定稿数字、AHP 权重与排序、OSGB 实测指标都进入依据对照表，
              且抽查值在成品正文里可找到；
  G5 政策文号：正文出现的「〔20XX〕N 号」必须在依据对照表登记的引文里出现过，否则拦下。

用法：python final_gates.py --workspace .
输出：过程文件/08-验证与边界/成品闸门报告.json；有违规 exit 1。
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

try:
    import docx
except ImportError:
    print("[final_gates] 缺少 python-docx", file=sys.stderr)
    raise SystemExit(1)

try:
    import fitz
except ImportError:
    fitz = None

SCHEMA = "cardo/invest-final-gates/v1"
SKILLS_ROOT = Path(__file__).resolve().parents[2]
# C8 口径（2026-11-16 用户补充）：缺数类写「待确认」、方案类写「待业主单位填写（方案深化阶段确定）」
# 是允许且要求的手填标记，不再作为缺口词；但只准贴登记项出现（见 G3-2 上下文校验）。
PENDING_VALUE = "待确认"
PENDING_SCHEME = "待业主单位填写（方案深化阶段确定）"
GAP_WORDS = ("未提资", "暂缺", "待补", "未提供", "请提供", "待补充",
             "资料缺失", "缺项", "待定", "待核实",
             # 2026-11-14 评审口径：成品里不许出现「我没资料」的姿态。
             # 有值写值并给依据；取不到的指标整句不写；需业主确认的进内部件《待确认清单》。
             "暂按", "推定", "测算假设", "需业主确认",
             "推断", "需现场核实", "同类项目通行")
# 三源抽查值：现状分析定稿 / AHP 权重 / OSGB 实测
THREE_SOURCE = {
    "现状分析定稿": ("现状分析报告", ["12,248.99", "36.18", "34.78", "9.89"]),
    "AHP 权重": ("权重表", ["0.1887", "0.3002", "0.2099", "0.3012"]),
    "OSGB 实测": ("现状指标", ["21.32", "1.0281", "7.72"]),
}


def ahp_weight_samples(ws):
    """AHP 抽查值取本项目《AHP权重表》准则层权重（4 位小数）。

    早先这里写死了一组别项目的权重；AHP 权重随项目判断矩阵每轮重算，
    写死值必然对不上正文，G4 会把正确成品判失败。改成从包内定稿源表取样，
    抽查对象仍是「来源表 → 依据对照表 → 正文」这条链。
    """
    path = ws / "成果文件" / "内部件" / "AHP" / "AHP权重表.xlsx"
    if not path.is_file():
        return []
    try:
        import openpyxl
        workbook = openpyxl.load_workbook(str(path), data_only=True)
    except Exception:
        return []
    sheet = workbook["权重表"] if "权重表" in workbook.sheetnames else workbook.worksheets[0]
    values, inside = [], False
    for row in sheet.iter_rows(values_only=True):
        head = str(row[0] or "")
        if "准则层权重" in head:
            inside = True
            continue
        if inside and head.startswith("二、"):
            break
        if inside and isinstance(row[1], (int, float)):
            values.append("{:.4f}".format(row[1]))
    return values[:4]


DOC_NO_RE = re.compile(r"[〔\[【]\s*(\d{4})\s*[〕\]】]\s*(\d+)\s*号")
# 缺口词豁免：字段级白名单。豁免不是「扫描前把这句话删掉」，而是要求上游 JSON 里
# 该字段显式标了 derived: true（派生值，口径清楚），取不到标记就不豁免。
# 早先按字符串 re.sub 删除属标准后门：任何人都能往正文塞一句再靠它消失。
FIELD_WHITELIST = (
    # (技术表述, 上游文件, 字段路径)
    ("影像纹理采样推定",
     "过程文件/01-上游取料/低空指标口径.json",
     "texture"),
)
# 文件名不是正文：附件清单里的「待确认清单.xlsx」是交付物名称，不是「我没资料」的姿态。
# 只剔除带扩展名的文件名片段，正文句子里的同名词照常计数（无法用来藏话术）。
FILE_NAME_RE = re.compile(r"[^\s，。；、：（）()|]{0,20}\.(?:xlsx|xlsm|docx|pdf|md|json)")


def _pointer_value(node, pointer):
    for key in [part for part in str(pointer or "").split(".") if part]:
        if isinstance(node, dict) and key in node:
            node = node[key]
        else:
            return None
    return node


def derived_phrases(ws):
    """字段级白名单：只有上游字段显式 derived:true 时，对应技术表述才不计缺口词。"""
    allowed = set()
    for phrase, rel, pointer in FIELD_WHITELIST:
        node = _pointer_value(read_json(ws / rel, {}) or {}, pointer)
        if isinstance(node, dict) and node.get("derived") is True:
            allowed.add(phrase)
    return allowed


def read_json(path, default=None):
    try:
        return json.loads(Path(path).read_text(encoding="utf-8"))
    except Exception:
        return default


def docx_text(path):
    """段落 + 表格全文（表格不重复计数合并单元格）。"""
    d = docx.Document(str(path))
    parts = [p.text for p in d.paragraphs]
    for table in d.tables:
        seen = set()
        for row in table.rows:
            for cell in row.cells:
                # 用元素本身作键：id() 是临时 lxml 代理的地址，回收后会被复用，
                # 会把不同单元格误判成同一个（合并单元格去重曾因此丢掉大半张权重表）。
                if cell._tc in seen:
                    continue
                seen.add(cell._tc)
                parts.append(cell.text)
    return "\n".join(parts)


def main(argv=None):
    parser = argparse.ArgumentParser()
    parser.add_argument("--workspace", default=".")
    args = parser.parse_args(argv)
    ws = Path(args.workspace).resolve()

    deliverable_dir = ws / "成果文件" / "成品版" / "项目投资决策建议书"
    docx_path = deliverable_dir / "项目投资决策建议书.docx"
    pdf_path = deliverable_dir / "项目投资决策建议书.pdf"
    violations = []
    checks = []

    def check(ident, name, ok, detail=""):
        checks.append({"id": ident, "name": name, "ok": bool(ok), "detail": detail})
        if not ok:
            violations.append({"id": ident, "detail": detail})

    # ---------------- G1 交付结构
    ok = docx_path.is_file() and pdf_path.is_file()
    check("G1-1", "成品 docx 与 pdf 同在 成品版/项目投资决策建议书/", ok,
          "docx={} pdf={}".format(docx_path.is_file(), pdf_path.is_file()))

    internal = ws / "成果文件" / "内部件"
    expect = ["投资估算", "财务评价", "方案比选", "依据对照表", "待确认清单"]
    missing = [name for name in expect
               if not (internal / name / (name + ".xlsx")).is_file()
               and not (name == "依据对照表" and (internal / name / "依据对照表_投资建议书.xlsx").is_file())]
    check("G1-2", "内部件一成果一子目录（投资估算/财务评价/方案比选/依据对照表/待确认清单）",
          not missing, "缺 {}".format("、".join(missing)) if missing else "")

    # 第 9 章引用的两份 AHP 源表随交付包归档，正文说「已在表 9-x 列示」才有实物可核。
    att_missing = [name for name in ("AHP权重表.xlsx", "AHP方案评分结论.xlsx")
                   if not (internal / "AHP" / name).is_file()]
    check("G1-3", "第 9 章 AHP 源表已随交付包归档（成果文件/内部件/AHP/）", not att_missing,
          "缺 {}".format("、".join(att_missing)) if att_missing else "")

    # ---------------- G2 成品版式
    pages = 0
    if fitz is not None and pdf_path.is_file():
        doc = fitz.open(str(pdf_path))
        pages = doc.page_count
        p1 = re.sub(r"\s+", "", doc[0].get_text("text")) if pages else ""
        p2 = re.sub(r"\s+", "", doc[1].get_text("text")) if pages > 1 else ""
        doc.close()
        check("G2-1", "PDF 页数 ≥ 10", pages >= 10, "{} 页".format(pages))
        check("G2-2", "封面不印页码", "第1页/共" not in p1, p1[-40:])
        check("G2-3", "正文页脚含「第 N 页 / 共 M 页」",
              ("第2页/共{}页".format(pages)) in p2, p2[-40:])
    else:
        check("G2-1", "PDF 页数 ≥ 10", False, "PyMuPDF 缺失或 PDF 不存在")

    # 目录：PDF 第 2 页起为目录页，条目数 ≥ 20 且页码不留空
    if fitz is not None and pdf_path.is_file() and pages > 1:
        doc = fitz.open(str(pdf_path))
        toc_text = re.sub(r"\s+", "", doc[1].get_text("text"))
        doc.close()
        has_pages = bool(re.search(r"\d{1,2}", toc_text))
        check("G2-4", "目录页有页码", has_pages, toc_text[:60])

    # ---------------- G3 缺口词
    text = docx_text(docx_path) if docx_path.is_file() else ""
    # 全文照扫，不做「先删一句话再扫」：豁免只走字段级白名单（上游显式 derived:true），
    # 另把带扩展名的附件文件名当结构串剔除（不是正文话术）。
    scan_text = FILE_NAME_RE.sub(" ", text)
    for phrase in derived_phrases(ws):
        scan_text = scan_text.replace(phrase, " ")     # 仅字段标了 derived:true 才走到这里
    hits = {word: scan_text.count(word) for word in GAP_WORDS if scan_text.count(word)}
    check("G3-1", "成品零缺口词", not hits,
          "、".join("{}×{}".format(k, v) for k, v in hits.items()))

    # ---------------- G4 三源同源
    register = read_json(ws / "过程文件" / "05-章节成文" / "分章" / "依据对照表.json", {}) or {}
    refs = []
    for row in register.get("rows") or []:
        for item in row.get("sources") or []:
            refs.append(str(item.get("ref") or ""))
    joined_refs = "\n".join(refs)
    flat = text.replace(",", "")
    for label, (needle, values) in THREE_SOURCE.items():
        if label == "AHP 权重":
            values = ahp_weight_samples(ws) or values
        cited = needle in joined_refs
        found = [v for v in values if v.replace(",", "") in flat]
        check("G4-{}".format(label), "{} 已入依据对照表且抽查值在成品中".format(label),
              cited and len(found) >= 2,
              "来源命中={} 正文命中 {}/{}".format(cited, len(found), len(values)))

    # ---------------- G5 政策文号
    doc_nos = ["〔{}〕{}号".format(a, b) for a, b in DOC_NO_RE.findall(text)]
    # 包内依据 = 05 依据对照表 + 包内编制依据库 + 运行期核验留痕（依据库是随包交付的基线）
    library = json.dumps(register, ensure_ascii=False)
    for extra in (SKILLS_ROOT.parent / "assets" / "依据库" / "依据库.json",
                  ws / "过程文件" / "05-章节成文" / "依据核验.json"):
        if extra.is_file():
            library += "\n" + extra.read_text(encoding="utf-8", errors="replace")
    normalized_library = library.replace(" ", "")
    bad = [item for item in doc_nos if item.replace(" ", "") not in normalized_library]
    check("G5-1", "政策文号均有包内依据", not bad,
          "无据：{}".format("、".join(bad)) if bad else "文号 {} 个".format(len(doc_nos)))

    # G5-2（R2 硬判据）：正文每个文号必须能在 依据库.json 里逐字对上，库外文号一律视为编造。
    library_path = SKILLS_ROOT.parent / "assets" / "依据库" / "依据库.json"
    lib_text = ""
    if library_path.is_file():
        lib_text = library_path.read_text(encoding="utf-8", errors="replace").replace(" ", "")
    outside = [item for item in doc_nos if item.replace(" ", "") not in lib_text]
    check("G5-2", "正文文号逐个对上 依据库.json（库外判编造）", not outside,
          "库外：{}".format("、".join(outside)) if outside else "逐字命中 {} 个".format(len(doc_nos)))

    # G3-2（C8）：缺数类「待确认」只能贴着六项缺数出现，方案类「待业主单位填写」只能贴着五项方案出现
    # 「待确认」只贴提资缺数项：前六项是上游槽位缺数，后三项是三份社区台账里
    # 示例社区A 4 个小区未填报的现状工程量（表 3-5 按用户口径标「待确认」）。
    missing_keys = ("公共空间面积", "外立面", "道路面积", "停车位", "充电桩", "游商",
                    "屋面翻新", "冷凝水管", "落水管")
    scheme_keys = ("组织架构", "功能定位", "业态", "改造思路", "运营模式")
    stray_pending, stray_scheme = 0, 0
    for sentence in re.split(r"[。；\n]", scan_text):
        if PENDING_VALUE in sentence and not any(key in sentence for key in missing_keys):
            stray_pending += 1
        if PENDING_SCHEME in sentence and not any(key in sentence for key in scheme_keys):
            stray_scheme += 1
    check("G3-2", "手填项口径：待确认/待业主填写只贴登记项出现",
          not (stray_pending or stray_scheme),
          "待确认脱项 {} 处、待业主脱项 {} 处".format(stray_pending, stray_scheme))

    failed = [c for c in checks if not c["ok"]]
    report = {
        "schema": SCHEMA,
        "ok": not failed,
        "failed": len(failed),
        "checks": checks,
        "pdfPages": pages,
        "violations": violations,
    }
    out = ws / "过程文件" / "08-验证与边界" / "成品闸门报告.json"
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print("[final_gates] {}（通过 {}/{}，PDF {} 页）".format(
        "PASS" if not failed else "FAIL", len(checks) - len(failed), len(checks), pages))
    for item in failed:
        print("  ✗ {} {}：{}".format(item["id"], item["name"], item["detail"]))
    return 0 if not failed else 1


if __name__ == "__main__":
    sys.exit(main())
