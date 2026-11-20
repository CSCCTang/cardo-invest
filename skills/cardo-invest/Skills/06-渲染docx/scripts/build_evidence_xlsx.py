#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""build_evidence_xlsx.py -- 把 05 的《依据对照表.json》渲染成人能直接翻的 xlsx。

三个主页签：
  按段对照   一章一页一行：章 / 节 / 段序 / 类型 / 正文摘要 / 数值 / 已绑定 / 待人工核 / 出处 / 状态
  逐数明细   每个数值一行：数值 → 出自哪个文件# 哪个键（专家追问时按这个翻）
  人工核     绑不上字段的数值（逐条列出，不与「有源」混排）
  来源文件清单  文件全路径 + 真实的绑定率/未解析数
  未解析     引用不实的清单（查不到出处的，单列出来，不藏在正文里）

用法：
    python build_evidence_xlsx.py --project-dir <工作区> [--out 成果文件/依据对照表_投资建议书.xlsx]
退出码：0 生成成功；2 缺对照表（先跑 05）。
"""

import argparse
import json
import os
import re

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.properties import PageSetupProperties

HEAD_FILL = PatternFill("solid", fgColor="1F3864")
HEAD_FONT = Font(color="FFFFFF", bold=True, size=10)
BAD_FILL = PatternFill("solid", fgColor="FCE4E4")
WRAP = Alignment(wrap_text=True, vertical="top")

# 行首禁则：这些全角标点不得出现在折行后的行首（Excel 本身不做禁则，必须预折行）。
# 只收全角标点：半角「.」是「投资估算.json」的一部分，当成标点会把整段路径并成超宽行。
NO_HEAD = "，。、；：？！）》」』】”’"


# 断点优先落在这些字符之后（路径按「/」、定位按「#」断，不把 05_过程 拆成两半）
BREAK_AFTER = "/#；;，,、"


# 文件名整体不拆：「现状分析报告.docx」连扩展名一起算一个单位。
# 之前按字符切，拉丁串会以「.docx」开头落到下一行（实测 256 处行首「.」）。
FILE_RE = re.compile(r"[^\s/\\#；;，,、]*\.(?:docx|xlsx|xlsm|json|md|ya?ml|csv|txt|pdf|doc|xls)")
NO_TAIL_START = "."     # 行首不允许以点号开头（扩展名被拆掉的典型症状）


def _plain_tokens(text):
    out, buf = [], ""
    for ch in text:
        if ch in BREAK_AFTER:
            out.append(buf + ch)
            buf = ""
        elif ord(ch) > 127:
            if buf:
                out.append(buf)
                buf = ""
            out.append(ch)
        else:
            buf += ch
    if buf:
        out.append(buf)
    return out


def _tokens(text):
    """切成不拆开的最小单位：中文逐字，文件名整体保留，其余拉丁串整体保留。"""
    out, pos = [], 0
    for match in FILE_RE.finditer(text):
        out.extend(_plain_tokens(text[pos:match.start()]))
        out.append(match.group(0))
        pos = match.end()
    out.extend(_plain_tokens(text[pos:]))
    return out


def _width(text):
    return sum(2 if ord(ch) > 127 else 1 for ch in text)


def _fix_line_heads(lines):
    """兜底：任何以标点/点号开头的行，都把上一行末字挪下来陪它。

    贪心折行在「本行恰好占满」时会把标点单独开一行；合并到上一行又会让该行超宽，
    Excel 再折一次还是把标点顶到行首（实测 p4 两处行首「，」）。挪末字下来既
    让行首有正文，又不让上一行超宽。
    """
    index = 1
    while index < len(lines):
        line = lines[index]
        if not line or line[0] not in NO_HEAD + NO_TAIL_START:
            index += 1
            continue
        prev = lines[index - 1]
        # 上一行末字是中文才挪（拉丁串/文件名是一个单位，不能拆）
        if len(prev) > 1 and ord(prev[-1]) > 127 and prev[-1] not in NO_HEAD:
            lines[index - 1] = prev[:-1]
            lines[index] = prev[-1] + line
        else:
            lines[index - 1] = prev + line[0]
            lines[index] = line[1:]
            if not lines[index]:
                lines.pop(index)
                continue
        index += 1
    return [line for line in lines if line]


def wrap_cell(value, width):
    """按列宽把长文本预先折成多行，并保证禁则标点不落行首。

    为什么要预折：Excel 的自动换行按字符硬切，长路径会在「/」「_」处断，
    标点会被顶到行首（实测 46 处），行高也会失控（200 行撑到 465 页）。
    这里按显示宽度算好断点、把行首标点并回上一行，行高按行数写死。
    """
    text = " ".join(str(value if value is not None else "").split())
    if not text:
        return "—"
    limit = max(6, int(width) - 3)
    lines, cur, used = [], "", 0
    for token in _tokens(text):
        weight = _width(token)
        if used + weight > limit:
            if (token[0] in NO_HEAD or token[0] in NO_TAIL_START) and cur:
                # 行首禁则：标点/扩展名点号并回本行末，不在下一行开头出现
                lines.append(cur + token)
                cur, used = "", 0
                continue
            if cur:
                lines.append(cur)
                cur, used = "", 0
        if weight > limit:                 # 单个超长串（无断点的长路径）硬切
            for ch in token:
                w = 2 if ord(ch) > 127 else 1
                if used + w > limit:
                    if ch in NO_HEAD and cur:
                        cur += ch
                    lines.append(cur)
                    cur, used = "", 0
                    if ch in NO_HEAD:
                        continue
                cur += ch
                used += w
            continue
        cur += token
        used += weight
    if cur:
        lines.append(cur)
    return "\n".join(_fix_line_heads(lines))



def _position(row):
    """「CH-04 4.2 第3段」：只留节号，不带节标题——带上标题每行都多折一行，全表多出十几页。"""
    section = str(row.get("section") or "")
    number = section.split(" ")[0] if section else ""
    return "{} {} 第{}段".format(row.get("chapter", ""), number, row.get("seq", "")).replace("  ", " ")


def write_sheet(ws, headers, rows, widths, bad_flags=None, wrap_cols=None):
    # 表头也走预折行：Excel 自己折会把「…「来源文件清单」）」折成以「」）」开头的行，
    # 每页重复表头就每页一处行首标点（实测 25 处）。表头加粗、字宽与正文不同，
    # 折行宽度要多留余量，否则 Excel 会把预折好的行再折一次。
    headers = [wrap_cell(value, max(6, widths[index] - 4)) for index, value in enumerate(headers)]
    ws.append(headers)
    for cell in ws[1]:
        cell.fill, cell.font = HEAD_FILL, HEAD_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 13.5 * (max(str(h).count("\n") + 1 for h in headers)) + 4
    wrap_cols = wrap_cols or list(range(1, len(widths) + 1))
    body_font = Font(size=10)
    for idx, row in enumerate(rows):
        cells = []
        for col, value in enumerate(row, start=1):
            text = wrap_cell(value, widths[col - 1]) if col in wrap_cols else value
            cells.append(text)
        ws.append(cells)
        if bad_flags and bad_flags[idx]:
            for cell in ws[ws.max_row]:
                cell.fill = BAD_FILL
        tallest = max((str(c).count("\n") + 1 for c in cells), default=1)
        # 行高按预折行数写死：不写死 Excel 会按自身换行重算，高度失控；
        # 也不能截断（高度压到 8 行会把折到第 9 行的路径藏掉）。
        ws.row_dimensions[ws.max_row].height = 13.5 * tallest + 4
        for cell in ws[ws.max_row]:
            cell.font = body_font
    for col, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(col)].width = width
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = WRAP
    ws.freeze_panes = "A2"
    # 打印：A4 横向、宽度压到 1 页、表头每页重复（早先未设，横向被切成 3 段、每页只剩 4 列）
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_title_rows = "1:1"
    ws.print_options.horizontalCentered = True
    ws.page_margins.left = ws.page_margins.right = 0.3
    ws.page_margins.top = ws.page_margins.bottom = 0.5



def main(argv=None):
    ap = argparse.ArgumentParser(description="Invest 依据对照表 → xlsx")
    ap.add_argument("--project-dir", required=True)
    ap.add_argument("--register", default="")
    ap.add_argument("--out", default="")
    args = ap.parse_args(argv)
    ws_root = os.path.abspath(args.project_dir)
    register_path = args.register or os.path.join(
        ws_root, "过程文件/05-章节成文/分章/依据对照表.json")
    if not os.path.isfile(register_path):
        print("[依据对照表] 缺 {}：请先跑 05 分章成文（会同目录产出对照表）。".format(register_path))
        return 2
    with open(register_path, encoding="utf-8") as f:
        data = json.load(f)

    rows = data.get("rows") or []
    out = args.out or os.path.join(ws_root, "成果文件/依据对照表_投资建议书.xlsx")
    os.makedirs(os.path.dirname(out) or ".", exist_ok=True)

    wb = Workbook()
    summary = wb.active
    summary.title = "按段对照"
    by_block, bad_flags = [], []
    detail_rows = []
    file_map = {}

    def brief_ref(item):
        """「文件名#字段」：全路径每次重复会把行高和页数放大好几倍，完整路径进来源文件清单。"""
        path = str(item.get("path") or "").split("#")[0]
        ref = str(item.get("ref") or "")
        pointer = "#" + ref.split("#", 1)[1] if "#" in ref else ""
        name = os.path.basename(path) or ref.split("#")[0]
        if path:
            file_map.setdefault(name, path)
        return name + pointer

    for row in rows:
        refs = row.get("sources") or []
        value_refs = row.get("valueRefs") or []
        # 按值绑定优先：一个值一个「文件名#字段」，同一字段只列一次。
        # 只有块里一个值都绑不上时，才退回节级来源（否则就是又把整节 5–9 个来源并排抄一遍）。
        bound = []
        seen_ref = set()
        for item in value_refs:
            ref = str(item.get("ref") or "")
            if not ref or ref in seen_ref:
                continue
            seen_ref.add(ref)
            bound.append(item)
        if bound:
            # 一行绑定三个以上字段时不再并排罗列（表 8-3/8-5/9-1 会挤成 5–9 条）；
            # 逐数明细里每个值都有自己的一行，这里指过去即可。
            where = ("见逐数明细" if len(bound) > 2
                     else "；".join(brief_ref(item) for item in bound) or "—")
        else:
            where = ("见逐数明细" if len(refs) > 2
                     else "；".join(
                         brief_ref(item) if item.get("path") else (item.get("note") or item.get("ref") or "")
                         for item in refs) or "—")
        brief = where
        kinds = []
        for item in refs:
            status = item.get("status")
            if status == "可解析" and "json" in (item.get("ref") or ""):
                kinds.append("数据文件")
            elif status == "可解析":
                kinds.append("材料")
            elif status == "说明":
                kinds.append("说明")
            elif status:
                kinds.append(status)
        by_block.append([
            _position(row),
            row.get("kind", ""), row.get("excerpt", ""),
            "、".join(dict.fromkeys(row.get("numbers") or [])) or "—",
            row.get("boundCount", 0), row.get("unboundCount", 0),
            "+".join(dict.fromkeys(kinds)) or "—",
            brief, row.get("status", ""),
        ])
        bad_flags.append(row.get("status") == "引用不实")
        # 逐数明细只登记「值 → 唯一字段」：每个数值用它自己的 valueRefs 那一行，
        # 不再整段共用第一个来源；裸整数序号（1、9、7、0828）与文件名里的数字不算数值。
        seen = set()
        for item in value_refs:
            token = str(item.get("number") or "")
            if not token or not item.get("ref") or token in seen:
                continue
            if re.fullmatch(r"\d{1,6}", token):
                continue
            seen.add(token)
            detail_rows.append([_position(row), token, brief_ref(item)])
    write_sheet(summary, ["位置", "类型", "正文摘要", "数值", "已绑定", "待人工核", "依据类型",
                          "出处（文件名#字段，全路径见「来源文件清单」）", "状态"],
                by_block, [24, 7, 44, 20, 6, 7, 9, 44, 20], bad_flags)

    detail = wb.create_sheet("逐数明细")
    write_sheet(detail, ["位置", "数值", "出处（文件名#字段，全路径见「来源文件清单」）"],
                detail_rows, [24, 18, 96])

    # 绑不上字段的数值单独一页：它们不能混在「有源」里冒充逐数有源，
    # 也不能只给一个总数——评审要按位置逐条去看、逐条判「补字段还是改为定性表述」。
    manual_items = data.get("manual") or []
    manual_rows = []
    for item in manual_items:
        section = str(item.get("section") or "")
        number = section.split(" ")[0] if section else ""
        manual_rows.append([
            "{} {} 第{}段".format(item.get("chapter", ""), number, item.get("seq", "")).replace("  ", " "),
            item.get("kind", ""), item.get("number", ""), item.get("excerpt", ""),
        ])
    manual_sheet = wb.create_sheet("人工核")
    write_sheet(manual_sheet, ["位置", "类型", "数值", "正文摘要"], manual_rows,
                [24, 7, 16, 60])

    files = wb.create_sheet("来源文件清单")
    write_sheet(files, ["文件名", "工作区内完整路径"],
                [[name, path] for name, path in sorted(file_map.items())], [34, 88])

    problems_list = data.get("problems") or []
    stats = data.get("stats") or {}
    bound = stats.get("valueBound", 0)
    unbound = stats.get("valueUnbound", 0)
    fact = bound + unbound
    percent = (lambda n: "{:.1f}%".format(100.0 * n / fact) if fact else "0.0%")
    if problems_list:
        problems = wb.create_sheet("未解析")
        write_sheet(problems, ["问题"], [[item] for item in problems_list], [96],
                    [True] * len(problems_list))
    # 实报：早先这里只写「未解析问题：无」，读起来像全册逐数有源；
    # 实际有近三成数值绑不上字段。总数、占比、去向都写出来，读者才能自己判断可信度。
    files.append([])
    files.append(["事实数值合计", "{}（已绑定 {}，待人工核 {}）".format(fact, bound, unbound)])
    files.append(["数值绑定率", "{}（已绑定 / 事实数值）".format(percent(bound))])
    files.append(["待人工核（未绑定）", "{}（{}），逐条见「人工核」页".format(unbound, percent(unbound))])
    files.append(["结构值（序号/编号）", "{}，非数据事实，不计入绑定率".format(
        stats.get("valueStructural", 0))])
    files.append(["查表标准值（RI/CR 阈值）", "{}，来源为规范查表，不绑项目字段".format(
        stats.get("standardValue", 0))])
    files.append(["未解析引用（查不到文件/键）", "{}".format(len(problems_list))])
    files.append(["块状态", "全绑定 {} / 部分绑定 {} / 待人工核 {} / 引用不实 {}".format(
        stats.get("blocksFullBound", 0), stats.get("blocksPartial", 0),
        stats.get("blocksManual", 0), len(problems_list))])

    wb.save(out)
    print("[依据对照表] {} → {}（{} 块 / 已绑定 {} / 待人工核 {} / 引用不实 {}）".format(
        register_path.replace(ws_root + os.sep, ""), os.path.relpath(out, ws_root),
        len(by_block), bound, unbound, len(problems_list)))
    if stats:
        print("  块有源 {}/{}".format(stats.get("withSource"), stats.get("blocks")))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
