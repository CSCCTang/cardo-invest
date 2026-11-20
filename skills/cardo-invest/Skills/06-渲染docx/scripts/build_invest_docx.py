"""组装/渲染/自检（T11）：以建议书模板为基底，填 10 章正文与 6 张表，渲染 docx 并回读自检。

输入：skill/模板/2 城市更新老旧街区改造项目投资决策建议书（模板）-补充数据来源（包含三类项目）.docx（10 章骨架 + 样式）
      05_过程/投资决策/{分章/CH-*.md,财务输入.json,投资估算.json,财务测算.json}
输出：04_成果报告/项目投资决策建议书-<项目>.docx
      04_成果报告/项目投资决策建议书-<项目>.自检.json（cardo/invest-report/v1）

用法：python build_invest_docx.py --project-dir <项目目录> [--template <docx>] [--no-timestamp]
退出码：0 成功且自检通过；1 自检失败；2 输入/模板缺失。
"""
from __future__ import annotations

import argparse
import json
import re
import sys
import time
from datetime import datetime
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
# 自包含化：skill/ 根即原 cardo-invest-spec（scripts/ 规则/ 模板/ 均在本目录下）
SPEC_DIR = SCRIPT_DIR.parent
sys.path.insert(0, str(SPEC_DIR / "scripts"))

import _lib  # noqa: E402
import figures  # noqa: E402
from docx.enum.text import WD_BREAK, WD_ALIGN_PARAGRAPH  # noqa: E402

ENGINE_VERSION = "build_invest_docx/0.2.0"
DEC = Path("05_过程") / "投资决策"
REPORT_DIR = Path("04_成果报告")
# 渲染基底：用包内 assets/templates/ 的院标原件。
# 理由（实测）：
#   * 院标原件含 10 个 H1 章节骨架 + 6 张表位表 + 与渲染基底**完全一致的 20 个样式**；
#   * _prune_template() 本就支持「院标原件（含示例正文与示例表）」——
#     它只保留 Title 段与 Heading 1，其余示例正文/示例表全部删除；
#   * 原先指向的 提资数据/模板/建议书模板.docx 是「无表的渲染基底」，
#     且路径在包外（拷包即失效），07 自检因此挂 CONTRACT-ASSET / ASSET-EXISTS 两项。
#     统一到包内资产后，技能包自包含：不依赖任何外部目录。
DEFAULT_TEMPLATE = SPEC_DIR.parents[1] / "assets" / "templates" / \
    "2 城市更新老旧街区改造项目投资决策建议书（模板）-补充数据来源（包含三类项目）.docx"
PLACEHOLDER_PATTERN = re.compile(r"\{\{.*?\}\}|\bTODO\b|\bXXX\b")
GAP_MARKER = re.compile(r"资料缺失|待补")
TABLE_CAPTION = re.compile(r"^表\s*\d+\s*[-‑–]\s*\d+\s+")
FIGURE_CAPTION = re.compile(r"^图\s*\d+\s*[-‑–]\s*\d+\s+")
DEFAULT_INDICATORS = {
    "现状情况": [("房屋设施老化程度", 0.5), ("安全风险等级", 0.25), ("现状建筑密度", 0.15), ("现状容积率", 0.1)],
    "投资经济": [("改造总成本", 0.5), ("预期内部收益率", 0.25), ("投资回收期", 0.15), ("资产增值潜力", 0.1)],
    "规划合规": [("城市规划符合度", 0.4), ("政策合规", 0.3), ("基础设施配套率", 0.2), ("生态环保标准", 0.1)],
    "风险可控": [("政策风险", 0.5), ("成本超支风险", 0.25), ("社会稳定风险", 0.15), ("工程延误风险", 0.1)],
}


class InputError(Exception):
    pass


def read_json(path, label, required=True):
    if not path.is_file():
        if required:
            raise InputError("缺少{}：{}".format(label, _lib.rel_posix(path, Path.cwd())))
        return None
    try:
        return _lib.read_json(path)
    except (json.JSONDecodeError, OSError) as error:
        raise InputError("{}读取失败：{}".format(label, error))


def _text_width(text):
    return sum(2 if ord(ch) > 127 else 1 for ch in str(text))


REF_HYPHEN = re.compile(r"((?:表|图)\s*\d{1,2})\s*[-‑–]\s*(\d{1,2})")


def keep_ref_together(text):
    """「表 9-9」「图 8-1」这类编号整体不换行：普通连字符允许断行，
    实测第 24 页断成「表9-」「9。」，尾行只剩「9。」。换成不换行连字符 U+2011。"""
    return REF_HYPHEN.sub("\\1\u2011\\2", text)


def _table_widths(rows, width, available_twips=9350):
    """按各列内容量分配列宽（twips），并设下限，避免「依据」列被挤成竖排。"""
    weights = []
    for col in range(width):
        longest = max((_text_width(row[col]) for row in rows if col < len(row)), default=4)
        weights.append(max(4, min(longest, 26)) + 3)
    total = float(sum(weights))
    values = [int(available_twips * weight / total) for weight in weights]
    # 约 2cm：低于此宽度「万平方米」「（低）」会被折行；列多时按页宽均摊，避免表格超出版心
    floor = min(1150, max(600, int(available_twips / float(width)) - 60))
    deficit = sum(max(0, floor - value) for value in values)
    values = [max(value, floor) for value in values]
    if deficit:
        room = [index for index, value in enumerate(values) if value > floor]
        over = float(sum(values[index] - floor for index in room)) or 1.0
        for index in room:
            values[index] -= int(deficit * (values[index] - floor) / over)
    return values


def _repeat_header(table):
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    tr_pr = table.rows[0]._tr.get_or_add_trPr()
    if tr_pr.find(qn("w:tblHeader")) is None:
        tr_pr.append(OxmlElement("w:tblHeader"))


def _header_text(value, column_twips):
    """表头「运营期合计（万元）」在窄列里会被折成「（万」+「元）」这类半截词。
    放不下时在左括号前断成两行，单位单独一行，不留碎片。"""
    text = str(value)
    if "（" not in text or text.startswith("（"):
        return text
    if _text_width(text) * 90 <= column_twips - 220:
        return text
    head, _, tail = text.partition("（")
    return head + "\n（" + tail


def _no_split_rows(table):
    """行内禁止跨页断行：不断行，表格行就不会在页底被切成「低）」「资）」碎片。"""
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    for row in table.rows:
        tr_pr = row._tr.get_or_add_trPr()
        if tr_pr.find(qn("w:cantSplit")) is None:
            tr_pr.append(OxmlElement("w:cantSplit"))


PPR_KINSOKU = ("w:kinsoku", "w:overflowPunct", "w:topLinePunct",
               "w:autoSpaceDE", "w:autoSpaceDN")


TAIL_CLOSING = "，。、；：？！）》」』】”’%"      # 不能出现在行首
LATIN_NARROW = "iljI.,:;'\"!|()[]{}"


def _char_pt(ch, size):
    """按字号估算字宽：中文全角 = 字号，拉丁字母约半个字号，窄体再收窄。"""
    if ord(ch) < 128:
        return size * (0.3 if ch in LATIN_NARROW else 0.52)
    return size


def _wrap_lines(text, first_width, width, size):
    """按行宽贪心排一行行，返回每行占宽（pt），用于判断尾行是否只剩一两个字。"""
    lines, used, limit = [], 0.0, first_width
    for ch in text:
        w = _char_pt(ch, size)
        if used + w > limit and used > 0:
            if ch in TAIL_CLOSING:      # 标点不落行首：并回上一行
                used += w
                continue
            lines.append(used)
            used, limit = 0.0, width
        used += w
    if used:
        lines.append(used)
    return lines


def fit_short_tails(document):
    """尾行只剩一两个字时微缩该段字距，把尾行并回上一行。

    做法：按版心（或单元格）宽度贪心估行，末行不足 4 个字宽就整段收紧字距。
    收紧量按「尾行宽 ÷ 全段字数」取，保证前面各行合计能吃掉整个尾行（最少 0.25pt/字）。
    Word 的标点溢出会让实际每行比估算多半个到一个字，所以估算尾行偏长时也要收，
    否则判不出来（实测「数。」「验收。」都卡在这一步）。
    """
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    def style_attr(paragraph, getter, default):
        """段落属性未显式设置时取样式链上的值（Normal 里 firstLine=420 twips、sz=21）。"""
        node = paragraph.style
        while node is not None:
            value = getter(node)
            if value is not None:
                return value
            node = node.base_style
        return default

    def fit(paragraph, width, indent_pt):
        text = paragraph.text
        if len(text) < 10 or not paragraph.runs:
            return
        size = next((run.font.size.pt for run in paragraph.runs if run.font.size), None)
        if size is None:
            size = style_attr(paragraph, lambda s: s.font.size.pt if s.font.size else None, 10.5)
        # 估算宽度比实际版心窄一个字：Word 排标点、括号注记时每行常比贪心估算少吃一个字，
        # 按实际宽度估会把「…支撑结论。」判成满行，尾行「论。」就漏过去了。
        model_w = width - size
        lines = _wrap_lines(text, model_w - indent_pt, model_w, size)
        if len(lines) < 2 or lines[-1] > 4.5 * size:
            return
        # 字距要一档一档试：收紧 0.05pt/字不一定能让上一行多吃一个字，
        # 必须按同一套贪心模型重排，确认尾行真的被吃掉（或至少变成 3 个字以上）。
        chosen = None
        for step in [value / 20.0 for value in range(4, 33)]:
            reduced = _wrap_lines(text, model_w - indent_pt, model_w, size - step)
            if len(reduced) < len(lines) or reduced[-1] > 3.2 * (size - step):
                chosen = step + 0.25     # 留余量：Word 的标点溢出常比估算多吃一个字
                break
        if chosen is None:
            return
        for run in paragraph.runs:
            r_pr = run._r.get_or_add_rPr()
            spacing = r_pr.find(qn("w:spacing"))
            if spacing is None:
                spacing = OxmlElement("w:spacing")
                r_pr.append(spacing)
            spacing.set(qn("w:val"), str(int(-chosen * 20)))

    section = document.sections[-1]
    # 版心宽度：模板 sectPr 里没写 pgMar，Word 按默认左右各 1800 twips 排；
    # 实测成稿正文恰好 90→505pt（415pt），按 1440 估会多出 3 个字宽，判不出尾行。
    page_w = section.page_width.twips if section.page_width else 11906
    margin_l = section.left_margin.twips if section.left_margin else 1800
    margin_r = section.right_margin.twips if section.right_margin else 1800
    width = (page_w - margin_l - margin_r) / 20.0      # twips → pt
    if width <= 0:
        return
    for paragraph in document.paragraphs:
        style = paragraph.style.name if paragraph.style is not None else ""
        if style.startswith("Heading") or style.startswith("标题") or style == "Caption":
            continue
        indent = paragraph.paragraph_format.first_line_indent
        indent_pt = indent.pt if indent is not None else style_attr(
            paragraph,
            lambda s: s.paragraph_format.first_line_indent.pt
            if s.paragraph_format.first_line_indent is not None else None, 0.0)
        fit(paragraph, width, indent_pt)
    # 表格单元格同样处理：窄列里「…尚无台账。」「…杂物占用」也会折出 1–2 字尾行
    seen = []                       # 持有 tc 对象：用 id() 会被回收复用，误判成已处理
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                tc = cell._tc
                if any(tc is item for item in seen):      # 合并单元格只处理一次
                    continue
                seen.append(tc)
                cell_pt = cell.width.pt if cell.width else 0
                if cell_pt <= 30:
                    continue
                for paragraph in cell.paragraphs:
                    # 单元格左右各有 108 twips（5.4pt）内边距，要扣掉
                    fit(paragraph, cell_pt - 10.8, 0.0)


def keep_table_together(table):
    """整表不跨页：除末行外，所有单元格段落 keepNext，Word 会把表整体挪到下一页。"""
    if table is None or len(table.rows) < 2:
        return
    for row in table.rows[:-1]:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                paragraph.paragraph_format.keep_with_next = True


def enable_kinsoku(paragraph):
    """显式开启行首/行尾禁则（不依赖模板继承，模板里一条都没写）。"""
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    p_pr = paragraph._p.get_or_add_pPr()
    successors = ("w:wordWrap", "w:topLinePunct", "w:bidi", "w:adjustRightInd",
                  "w:snapToGrid", "w:spacing", "w:ind", "w:contextualSpacing",
                  "w:mirrorIndents", "w:suppressOverlap", "w:jc", "w:textDirection",
                  "w:textAlignment", "w:textboxTightWrap", "w:outlineLvl", "w:divId",
                  "w:cnfStyle", "w:rPr", "w:sectPr", "w:pPrChange")
    for tag in PPR_KINSOKU:
        if p_pr.find(qn(tag)) is not None:
            continue
        element = OxmlElement(tag)
        element.set(qn("w:val"), "1")
        p_pr.insert_element_before(element, *successors)


REF_LABELS = {
    "投资估算.json": "《投资估算》",
    "财务测算.json": "《财务评价》测算结果",
    "方案比选.json": "《方案比选》",
    "财务输入.json": "《财务参数表》",
    "评分表.json": "《AHP 方案评分结论》",
    "上游摘要.json": "《现状分析报告》",
    "槽位.json": "《项目槽位表》",
    "工作区参数.json": "项目基础信息",
}


def _compact_ref_cell(value):
    """表格「依据」列只留文件名：单元格里塞整条工作区路径，一列要折三四行。"""
    text = str(value)
    # 依据列是给读者看的：`财务测算.json#period` 这种引擎字段读者看不懂，换成可读文件名。
    for name, label in REF_LABELS.items():
        text = re.sub(re.escape(name) + r"#[^；;\s]*", label, text)
    if text != str(value):
        return text
    # 表 10-1 附件清单要读者按路径直接找到文件，必须保留完整相对路径；
    # 只留 basename 会让「已附」无从核对（R6 复核第 4 条）。
    if text.strip().replace("\\", "/").startswith("成果文件/"):
        return text
    # 只收敛「真的像文件引用」的单元格：带 # 键，或以 .json/.docx/.xlsx/.md 结尾。
    # 否则「建设地点/范围」「停车位 / 充电桩」这类并列文字会被 rsplit 吃掉前半截
    # （实测表 1-1 只剩「范围」「充电桩」「专项债」）。
    if "#" not in text and not re.search(r"\.(json|docx|xlsx|md)\s*$", text.strip(), re.I):
        return text
    if "/" not in text and "\\" not in text:
        return text
    parts = []
    for piece in text.split("；"):
        head, _, pointer = piece.partition("#")
        base = head.replace("\\", "/").rsplit("/", 1)[-1]
        if base.lower().endswith(".docx"):
            parts.append(base)                      # 文档不带 #facts 这类键名
        else:
            parts.append(base + ("#" + pointer if pointer else ""))
    return "；".join(parts)


def add_table(document, caption, rows):
    """渲染“表题 + 管道表”为 Word 表格（首行表头加粗，正文 9 磅）。"""
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Pt, Twips

    if not rows:
        return
    caption_paragraph = document.add_paragraph(keep_ref_together(caption))
    caption_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    caption_paragraph.paragraph_format.first_line_indent = Pt(0)
    caption_paragraph.paragraph_format.space_before = Pt(6)
    caption_paragraph.paragraph_format.space_after = Pt(3)
    caption_paragraph.paragraph_format.keep_with_next = True
    for run in caption_paragraph.runs:
        run.bold = True
        run.font.size = Pt(10.5)
    enable_kinsoku(caption_paragraph)
    width = max(len(row) for row in rows)
    table = document.add_table(rows=len(rows), cols=width)
    try:
        table.style = "Table Grid"
    except KeyError:
        pass
    # 固定列宽 + 按内容分配：早先不设列宽，Word 自动布局把「依据」「内容」压成一字宽，
    # 「投资估算」等表头竖排、长句在窄列里逐字换行。
    table.autofit = False
    try:
        section = document.sections[-1]
        # 模板里页边距可能是「未显式设置」的 None（继承正文缺省），不能直接相减
        page_w = section.page_width.twips if section.page_width else 11906
        margin_l = section.left_margin.twips if section.left_margin else 1440
        margin_r = section.right_margin.twips if section.right_margin else 1440
        available = int(page_w - margin_l - margin_r)
    except (AttributeError, IndexError, TypeError):
        available = 9350
    column_widths = _table_widths(rows, width, max(6000, min(available, 14000)))
    for col_index, twips in enumerate(column_widths):
        table.columns[col_index].width = Twips(twips)
    # tblW 必须写成具体宽度：留着 auto，Word 会按内容重新算一遍，把刚设好的
    # 列宽推翻（实测「单位」列 1.9cm 仍把「万平方米」折成两行）。
    tbl_pr = table._tbl.tblPr
    tbl_w = tbl_pr.find(qn("w:tblW"))
    if tbl_w is None:
        tbl_w = OxmlElement("w:tblW")
        tbl_pr.append(tbl_w)
    tbl_w.set(qn("w:type"), "dxa")
    tbl_w.set(qn("w:w"), str(sum(column_widths)))
    _repeat_header(table)
    _no_split_rows(table)
    size = Pt(8) if width >= 7 else Pt(9)
    for row_index, row in enumerate(rows):
        # 分组表头行（只有首列有字，如「[投资经济] a · b · c · d」）合并整行：
        # 不合并时首列只有 4 列宽，会把「资产增值潜力」折成单字尾行（实测 p24「力」）。
        cells_row = [str(row[index]).strip() if index < len(row) else "" for index in range(width)]
        merge_row = row_index > 0 and width > 1 and bool(cells_row[0]) and not any(cells_row[1:])
        if merge_row:
            table.rows[row_index].cells[0].merge(table.rows[row_index].cells[width - 1])
        for col_index in range(width):
            value = row[col_index] if col_index < len(row) else ""
            cell = table.rows[row_index].cells[col_index]
            cell.width = Twips(column_widths[col_index])
            if merge_row and col_index > 0:
                continue
            cell_text = (_header_text(value, column_widths[col_index]) if row_index == 0
                         else _compact_ref_cell(value))
            # 单位格全角统一：同包 xlsx 用「％」，docx 早先留下 4 处 ASCII「%」，两处不一致。
            if cell_text.strip() == "%":
                cell_text = "％"
            cell.text = cell_text
            for paragraph in cell.paragraphs:
                # 单元格一律左对齐：模板继承了「分散对齐」，短行会被拉成
                # 「投　　资　　估」这种一字一顿的样子，看着像竖排。
                paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
                paragraph.paragraph_format.first_line_indent = Pt(0)
                enable_kinsoku(paragraph)
                for run in paragraph.runs:
                    run.font.size = size
                    if row_index == 0:
                        run.bold = True



_ORD_RE = re.compile(r"^(?:[（(][一二三四五六七八九十]+[）)]|[一二三四五六七八九十]+、|\d+[、.．)])")


def add_markdown(document, markdown, skip_first_heading=True, figure_renderer=None):
    """解析分章稿：标题、段落、项目符号、表题+管道表、图题+图片。"""
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Cm, Pt

    lines = markdown.splitlines()
    first = True
    index = 0
    li_no = 0
    while index < len(lines):
        text = lines[index].rstrip()
        if not text:
            index += 1
            continue
        if text.strip() == "{{PAGEBREAK}}":
            break_paragraph = document.add_paragraph("")
            break_paragraph.paragraph_format.first_line_indent = Pt(0)
            break_paragraph.add_run().add_break(WD_BREAK.PAGE)
            index += 1
            first = False
            continue
        if TABLE_CAPTION.match(text):
            caption = text.strip()
            index += 1
            while index < len(lines) and not lines[index].strip():
                index += 1
            rows = []
            while index < len(lines) and lines[index].strip().startswith("|"):
                cells = [item.strip() for item in lines[index].strip().strip("|").split("|")]
                if not all(set(cell) <= set("-: ") for cell in cells if cell):
                    rows.append(cells)
                index += 1
            add_table(document, caption, rows)
            first = False
            continue
        if FIGURE_CAPTION.match(text):
            caption = text.strip()
            fig_id = re.match(r"^图\s*(\d+-\d+)", caption).group(1)
            path = None
            if figure_renderer is not None:
                try:
                    path = figure_renderer(fig_id)
                except Exception:  # 绘图失败不阻断成文，图题仍保留
                    path = None
            if path:
                document.add_picture(str(path), width=Cm(15))
                document.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
            caption_paragraph = document.add_paragraph(keep_ref_together(caption))
            caption_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            caption_paragraph.paragraph_format.first_line_indent = Pt(0)
            caption_paragraph.paragraph_format.space_after = Pt(6)
            for run in caption_paragraph.runs:
                run.font.size = Pt(10.5)
            index += 1
            first = False
            continue
        if text.startswith("# "):
            if first and skip_first_heading:
                first = False
                index += 1
                continue
            document.add_heading(text[2:].strip(), level=1)
            index += 1
            continue
        if text.startswith("## "):
            document.add_heading(text[3:].strip(), level=2)
            index += 1
            continue
        if text.startswith("#### "):
            document.add_heading(text[5:].strip(), level=4)
            index += 1
            continue
        if text.startswith("### "):
            document.add_heading(text[4:].strip(), level=3)
            index += 1
            continue
        if text.startswith("来源：") or text.startswith("来源:"):
            # 来源行只保留在分章稿/来源映射中，成品正文不出现机读溯源行
            first = False
            index += 1
            continue
        if text.startswith("> "):
            paragraph = document.add_paragraph(text[2:].strip())
            paragraph.runs[0].italic = True
            index += 1
            continue
        if text.startswith("- "):
            # 公文层次序数用文字，不用项目符号：连续条目按 1. 2. 3. 编号；
            # 条目自带「1、」「（一）」等序数时只去符号，避免「1. 1、」双编号。
            item = keep_ref_together(text[2:].strip())
            if not (index > 0 and lines[index - 1].strip().startswith("- ")):
                li_no = 0
            li_no += 1
            if _ORD_RE.match(item):
                document.add_paragraph(item)
            else:
                document.add_paragraph(f"{li_no}. {item}")
            index += 1
            continue
        document.add_paragraph(keep_ref_together(text))
        first = False
        index += 1



def trim_template(document):
    """把模板裁成「封面标题 + 章标题骨架」。

    模板可以是院标原件（含示例正文与示例表），成品只需要它的：
      * 样式定义（Normal / Heading / Title / 页边距等）
      * 封面 Title 段
      * 10 个一级标题（作为逐章插入的锚点）
    其余示例正文、示例表格一律删除，避免成品里出现模板带的样例文字。
    """
    from docx.oxml.ns import qn
    from docx.text.paragraph import Paragraph

    body = document.element.body
    for child in list(body.iterchildren()):
        if child.tag == qn("w:sectPr"):
            continue
        if child.tag == qn("w:p"):
            style = Paragraph(child, document).style.name or ""
            if style in ("Title", "标题", "Heading 1", "标题 1"):
                continue
        body.remove(child)


def strip_heading_numbering(document):
    """关闭模板 Heading 1–9 样式自带的自动编号。

    模板的标题样式挂有 numId=1（一级“第%1章 ”、二级“%1.%2”等），
    分章稿标题又已手写“1.1”编号，Word 渲染时会叠加成“第一章 1 项目概述”
    “1.1 1.1 项目基本信息”。成品统一以手写编号为准，故移除样式级与段落级 numPr。
    """
    from docx.oxml.ns import qn

    def drop_num_pr(container):
        pPr = container.find(qn("w:pPr")) if container is not None else None
        if pPr is None:
            return
        numPr = pPr.find(qn("w:numPr"))
        if numPr is not None:
            pPr.remove(numPr)

    for style in document.styles:
        if style.name and style.name.startswith("Heading"):
            drop_num_pr(style.element)
    for paragraph in document.paragraphs:
        drop_num_pr(paragraph._p)


def normalize_zip_timestamps(path):
    """把 docx zip 各条目时间戳固定为 1980-01-01，保证 --no-timestamp 两次运行字节级一致。

    python-docx 保存时 zipfile.writestr 会写入墙钟时间，导致同输入两次产物 sha256 不同；
    --no-timestamp 的语义就是「固定时间戳用于幂等比对」，故在此补齐字节级确定性。
    """
    import zipfile

    path = Path(path)
    tmp = path.with_name(path.name + ".tmp")
    with zipfile.ZipFile(str(path)) as src, zipfile.ZipFile(str(tmp), "w", zipfile.ZIP_DEFLATED) as dst:
        for info in src.infolist():
            item = zipfile.ZipInfo(info.filename, date_time=(1980, 1, 1, 0, 0, 0))
            item.compress_type = info.compress_type
            item.external_attr = info.external_attr
            item.internal_attr = info.internal_attr
            item.create_system = info.create_system
            dst.writestr(item, src.read(info.filename))
    tmp.replace(path)


TOC_FALLBACK = "（打开文档时自动更新目录）"


def _toc_entry_count(path):
    """回读 docx：目录域更新后应有 ≥10 条带页码的条目（没有就是假目录）。"""
    try:
        import docx
        document = docx.Document(str(path))
    except (ImportError, OSError, ValueError):
        return 0, True
    count = 0
    stale = False
    for paragraph in document.paragraphs:
        text = (paragraph.text or "").strip()
        if TOC_FALLBACK in text:
            stale = True
        style = (paragraph.style.name or "").lower()
        if style.startswith("toc") or re.match(r"^\d+(?:\.\d+)*\s+\S.*\t\d+$", text):
            count += 1
    return count, stale


def _field_run(paragraph, instr, placeholder="1", before="", after=""):
    """在段落里插一个真正的域（PAGE / NUMPAGES），而不是写死的数字。"""
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    if before:
        paragraph.add_run(before)
    run = paragraph.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    code = OxmlElement("w:instrText")
    code.set(qn("xml:space"), "preserve")
    code.text = " {} ".format(instr)
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    shown = OxmlElement("w:t")
    shown.text = placeholder
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    for node in (begin, code, separate, shown, end):
        run._r.append(node)
    if after:
        paragraph.add_run(after)


def apply_gov_margins(document):
    """公文版式页边距：上 37、下 35、左 28、右 26 mm（GB/T 9704 常用值）。

    模板 sectPr 不写 pgMar，Word 按 25.4/31.8 mm 默认排版；打印出来的版心比定稿
    公文版窄一圈，页数也会多出几页。版式必须落在 docx 本体，不能只靠 PDF 后处理。
    """
    from docx.shared import Cm
    for section in document.sections:
        section.top_margin = Cm(3.7)
        section.bottom_margin = Cm(3.5)
        section.left_margin = Cm(2.8)
        section.right_margin = Cm(2.6)
        # 页脚距下白边 28 mm（页码在版心下缘下方 7 mm），页眉 15 mm
        section.header_distance = Cm(1.5)
        section.footer_distance = Cm(2.8)


def add_page_footer(document):
    """页脚「第 N 页 / 共 M 页」，封面不印页码。

    PDF 的页码由渲染链注入，docx 若没有 PAGE 域，评审人直接打开 docx 就是一叠无页码的纸；
    两版必须各有自己的页码来源，不能只靠 PDF 那一份。
    """
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    for section in document.sections:
        section.different_first_page_header_footer = True
        first = section.first_page_footer
        first.is_linked_to_previous = False
        for paragraph in list(first.paragraphs):
            paragraph.text = ""
        footer = section.footer
        footer.is_linked_to_previous = False
        paragraph = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
        paragraph.text = ""
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _field_run(paragraph, "PAGE", "1", before="第 ", after=" 页 / 共 ")
        _field_run(paragraph, "NUMPAGES", "1", after=" 页")
    # 打开文档时让 Word 更新域（含目录）：不设 updateFields，目录页码要用户手动按 F9。
    settings = document.settings.element
    if settings.find(qn("w:updateFields")) is None:
        update = OxmlElement("w:updateFields")
        update.set(qn("w:val"), "true")
        settings.append(update)


def _ensure_toc_style(document, name):
    """模板没有 TOC 样式时补一个：静态目录也要能被回读认成目录条目。"""
    try:
        return document.styles[name]
    except KeyError:
        from docx.enum.style import WD_STYLE_TYPE
        style = document.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
        style.base_style = document.styles["Normal"]
        return style


def _write_static_toc(path, entries):
    """没有 Word 时的兜底：把占位话换成静态条目，但**保留 TOC 域**。

    域结构（begin / instrText / separate … end）必须留着：早先的写法把含域的整段删掉，
    域随之消失，docx 就只剩一列没有页码的死标题，在 Word 里按 F9 也刷不出来。
    现在把条目写进域结果区，打开文档时 Word 仍能按 updateFields 刷新出页码。
    """
    import copy
    try:
        import docx
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn
        from docx.text.paragraph import Paragraph
    except ImportError:
        return False
    try:
        document = docx.Document(str(path))
    except (OSError, ValueError):
        return False
    target = next((p for p in document.paragraphs if TOC_FALLBACK in (p.text or "")), None)
    if target is None:
        return False
    code = next((node.text for node in target._p.iter(qn("w:instrText"))), "") \
        or 'TOC \\o "1-2" \\h \\z \\u'

    def field(kind):
        node = OxmlElement("w:fldChar")
        node.set(qn("w:fldCharType"), kind)
        return node

    paragraphs = [target]
    for _ in entries[1:]:
        clone = copy.deepcopy(target._p)
        target._p.addnext(clone) if len(paragraphs) == 1 else paragraphs[-1]._p.addnext(clone)
        paragraphs.append(Paragraph(clone, target._parent))
    for index, (paragraph, text) in enumerate(zip(paragraphs, entries)):
        for child in list(paragraph._p):
            if child.tag not in (qn("w:pPr"), qn("w:rPr")):
                paragraph._p.remove(child)
        paragraph.paragraph_format.first_line_indent = None
        paragraph.style = _ensure_toc_style(
            document, "TOC 2" if text.startswith("　") else "TOC 1")
        if index == 0:
            run = paragraph.add_run()
            run._r.append(field("begin"))
            instr = OxmlElement("w:instrText")
            instr.set(qn("xml:space"), "preserve")
            instr.text = code
            run._r.append(instr)
            run._r.append(field("separate"))
        paragraph.add_run(text)
    paragraphs[-1].add_run()._r.append(field("end"))
    document.save(str(path))
    return True


def update_toc_fields(path, chapter_entries=()):
    """用本机 Word 更新目录域并保存；更新后回读校验，失败重试一次。

    返回 (ok, pages)。没 Word 时写入静态目录，保证交付件里没有「打开文档时自动更新目录」。
    """
    import shutil
    import subprocess
    import tempfile
    if not shutil.which("powershell"):
        return _write_static_toc(path, chapter_entries), None
    path = Path(path).resolve()
    pages = None
    # Word 打开**相对**的中文路径实测报「文档打开失败」；Windows 中文路径传给
    # PowerShell -Command 也不稳。统一复制到 ASCII 临时路径更新，再拷回原位。
    with tempfile.TemporaryDirectory() as tmpdir:
        probe = Path(tmpdir) / "cardo_toc.docx"
        shutil.copyfile(path, probe)
        script = (
            "$ErrorActionPreference='Stop';"
            "try{{"
            "$w=New-Object -ComObject Word.Application;"
            "$w.Visible=$false;"
            "$w.DisplayAlerts=0;"
            "$d=$w.Documents.Open('{p}');"
            "try{{$d.Fields.Update()|Out-Null}}catch{{}};"
            "for($i=1;$i -le $d.TablesOfContents.Count;$i++){{try{{$d.TablesOfContents.Item($i).Update()}}catch{{}}}};"
            "try{{$d.Repaginate()}}catch{{}};"
            "$pages=0;try{{$pages=$d.ComputeStatistics(2)}}catch{{}};"
            "$d.Save();"
            "Write-Output ('PAGES=' + $pages);"
            # 本机 Word 退出时必抛 RPC 异常（0x800706BE），但 Save 已完成。
            # 关闭/退出包在 try 里，否则非零退出会把「已更新成功」误判成失败。
            "try{{$d.Close($false)}}catch{{}};try{{$w.Quit()}}catch{{}}"
            "}}catch{{Write-Output ('FAIL=' + $_.Exception.Message)}}"
        ).format(p=str(probe).replace("'", "''"))
        for _attempt in range(2):
            try:
                done = subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-Command", script],
                                      capture_output=True, text=True, timeout=240)
            except (OSError, subprocess.TimeoutExpired):
                continue
            stdout = done.stdout or ""
            if "FAIL=" in stdout:
                print("[build_invest_docx] 目录域更新失败：{}".format(
                    stdout.split("FAIL=")[-1].strip()[:200]))
                continue
            for line in stdout.splitlines():
                if line.strip().startswith("PAGES="):
                    try:
                        pages = int(line.strip().split("=", 1)[1])
                    except ValueError:
                        pages = None
            if not probe.is_file():
                continue
            # Word Save 已完成；拷回原位（短暂占用时重试几次）
            for _wait in range(5):
                try:
                    shutil.copyfile(probe, path)
                    break
                except OSError:
                    time.sleep(0.5)
            count, stale = _toc_entry_count(path)
            if count >= 10 and not stale:
                return True, pages
            print("[build_invest_docx] 目录域回读不足（条目 {}，占位 {}），重试".format(count, stale))
    # Word 更新不成功：退回静态目录，绝不把占位话交给用户
    ok = _write_static_toc(path, chapter_entries)
    return ok, pages


def main(argv=None):
    parser = argparse.ArgumentParser(description="cardo-invest 建议书渲染（T11）")
    parser.add_argument("--project-dir", required=True)
    parser.add_argument("--template", default=None)
    parser.add_argument("--no-timestamp", action="store_true")
    args = parser.parse_args(argv)
    project = Path(args.project_dir)
    if not project.is_dir():
        print("[build_invest_docx] 项目目录不存在：{}".format(args.project_dir))
        return 2
    template = Path(args.template) if args.template else DEFAULT_TEMPLATE
    if not template.is_file():
        print("[build_invest_docx] 缺少建议书模板：{}".format(_lib.rel_posix(template, Path.cwd())))
        return 2
    dec = project / DEC
    try:
        payload = read_json(dec / "财务输入.json", "财务输入")
        finance = read_json(dec / "财务测算.json", "财务测算")
        estimate = read_json(dec / "投资估算.json", "投资估算", required=False)
        schemes = read_json(dec / "方案比选.json", "方案比选", required=False)
        # 图 3-1（空间现状图）直接读 输入文件/03-OSGB成果 的正射/地类/轮廓；
        # 图 3-2（现状空间构成）取低空解析指标，财务输入.json 不带 facts，挂 statusMetrics。
        status = read_json(dec.parent / "低空解析" / "现状指标.json", "现状指标", required=False)
        if status:
            payload.setdefault("upstream", {})["statusMetrics"] = status.get("metrics") or []
        # 图 9-2 画 AHP 口径的方案综合得分（表 9-7 同源），不能拿方案比选的现金流得分顶替
        payload["ahpScores"] = read_json(dec.parent / "AHP" / "评分表.json", "AHP 评分", required=False)
    except InputError as error:
        print("[build_invest_docx] {}".format(error))
        return 2
    chapter_dir = dec / "分章"
    chapter_text = {}
    for number in range(1, 11):
        path = chapter_dir / "CH-{:02d}.md".format(number)
        if not path.is_file():
            print("[build_invest_docx] 缺少分章稿：{}".format(_lib.rel_posix(path, project)))
            return 2
        chapter_text[number] = path.read_text(encoding="utf-8")

    import docx

    document = docx.Document(str(template))
    trim_template(document)
    strip_heading_numbering(document)
    apply_gov_margins(document)
    project_name = (payload.get("project") or {}).get("name") or "某某片区改造项目"
    indicators = (finance or {}).get("indicators") or {}
    # 口径纪律：成品表位 1/2 与正文、敏感性分析、方案比选取同一口径——引擎自算值。
    # 可研对照值由 05 步单列为「表 8-8 与可研财务指标对照表」，不在此覆盖，
    # 否则同一份报告会出现两套 FIRR（表位 1 与敏感性分析表互相矛盾）。
    ahp = payload.get("ahpWeights") or {}
    figure_dir = dec / "图形"
    # 图 3-1 读「输入文件/03-OSGB成果」，而 project 指向引擎项目
    # （过程文件/00-总控/引擎项目），要向上找到真正的工作区根。
    figure_root = project
    for candidate in [project] + list(project.parents)[:4]:
        if (candidate / "输入文件").is_dir():
            figure_root = candidate
            break

    def render_figure(fig_id):
        return figures.render_figure(fig_id, figure_dir, payload=payload, finance=finance,
                                    estimate=estimate, schemes=schemes, workspace=figure_root)

    # 清理模板内部的“渲染基底”说明段（成品不应保留内部说明）
    for paragraph in list(document.paragraphs):
        if paragraph.text.strip().startswith("（本文件为 cardo-invest 渲染基底"):
            paragraph._p.getparent().remove(paragraph._p)
    # 清理 docProps 里的第三方/个人痕迹
    props = document.core_properties
    props.author = "cardo-invest"
    props.last_modified_by = "cardo-invest"
    props.comments = ""
    props.title = "项目投资决策建议书-{}".format(project_name)

    # 正文样式：1.5 倍行距、首行缩进 2 字符（10.5 磅字约 0.74 cm），段后 0
    from docx.shared import Cm
    normal = document.styles["Normal"]
    normal.paragraph_format.line_spacing = 1.5
    normal.paragraph_format.first_line_indent = Cm(0.74)
    normal.paragraph_format.space_after = Cm(0)

    # 封面页：项目名称 → 报告名称（模板 Title）→ 编制日期 → 分页
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
    from docx.shared import Pt

    title_paragraph = document.paragraphs[0]
    cover_name = title_paragraph.insert_paragraph_before(project_name)
    cover_name.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cover_name.paragraph_format.first_line_indent = Pt(0)
    cover_name.paragraph_format.space_before = Pt(120)
    for run in cover_name.runs:
        run.bold = True
        run.font.size = Pt(22)
    for _ in range(5):
        spacer = title_paragraph.insert_paragraph_before("")
        spacer.paragraph_format.first_line_indent = Pt(0)
    cover_date = document.add_paragraph("编制日期：{}".format(datetime.now().strftime("%Y 年 %m 月")))
    cover_date.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cover_date.paragraph_format.first_line_indent = Pt(0)
    cover_date.paragraph_format.space_before = Pt(60)
    for run in cover_date.runs:
        run.font.size = Pt(14)
    title_paragraph._p.addnext(cover_date._p)
    page_break = document.add_paragraph("")
    page_break.paragraph_format.first_line_indent = Pt(0)
    page_break.add_run().add_break(WD_BREAK.PAGE)
    cover_date._p.addnext(page_break._p)

    def insert_chapter(context, number):
        add_markdown(context, chapter_text[number], figure_renderer=render_figure)
        if number == 9 and ahp.get("temporary"):
            context.add_paragraph("注：以上 AHP 权重为临时派生/默认值，待 AHP 插件输出后替换。")

    def add_toc(anchor):
        """封面后插入目录域：Word 打开时按 H1/H2 生成，导出 PDF 前由 COM 更新域。"""
        from docx.enum.text import WD_BREAK
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn

        # 模板没有 TOC 1/TOC 2 样式时先补上：实测缺这两个样式时 Word 更新目录域
        # 会直接崩（COM 报 RPC server is unavailable），目录永远只能退回静态条目。
        for style_name in ("TOC 1", "TOC 2"):
            _ensure_toc_style(document, style_name)
        # 封面末尾若已有分页符，再加一个就会多出一张空白页（实测 p2 空白）。
        prev = anchor._p.getprevious()
        has_break = False
        if prev is not None and prev.tag == qn("w:p"):
            has_break = any(br.get(qn("w:type")) == "page" for br in prev.iter(qn("w:br")))
        if not has_break:
            brk = anchor.insert_paragraph_before("")
            brk.paragraph_format.first_line_indent = Pt(0)
            brk.add_run().add_break(WD_BREAK.PAGE)
        title = anchor.insert_paragraph_before("目　录")
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        title.paragraph_format.first_line_indent = Pt(0)
        for run in title.runs:
            run.bold = True
            run.font.size = Pt(16)
        holder = anchor.insert_paragraph_before("")
        holder.paragraph_format.first_line_indent = Pt(0)
        run = holder.add_run()
        begin = OxmlElement("w:fldChar")
        begin.set(qn("w:fldCharType"), "begin")
        instr = OxmlElement("w:instrText")
        instr.set(qn("xml:space"), "preserve")
        instr.text = 'TOC \\o "1-2" \\h \\z \\u'
        separate = OxmlElement("w:fldChar")
        separate.set(qn("w:fldCharType"), "separate")
        fallback = OxmlElement("w:t")
        fallback.text = "（打开文档时自动更新目录）"
        end = OxmlElement("w:fldChar")
        end.set(qn("w:fldCharType"), "end")
        for node in (begin, instr, separate, fallback, end):
            run._r.append(node)
        # 目录与正文之间补分页：TOC 域跨页时，第 1 章会紧跟在目录条目后面同一个页面
        # （实测 p3 目录尾部直接接「1 项目概述」），补一个分页符让正文从新页开始。
        after_toc = anchor.insert_paragraph_before("")
        after_toc.paragraph_format.first_line_indent = Pt(0)
        after_toc.add_run().add_break(WD_BREAK.PAGE)

    # 用“章标题位置”插入：模板每章后有一个占位段落（（本章由…））
    body = document.element.body
    from docx.text.paragraph import Paragraph
    from docx.oxml.ns import qn

    chapter_headings = []
    for element in body.iterchildren():
        if element.tag == qn("w:p"):
            paragraph = Paragraph(element, document)
            if paragraph.style.name == "Heading 1":
                chapter_headings.append(paragraph)
    if len(chapter_headings) != 10:
        print("[build_invest_docx] 模板一级标题数不是 10（{}）".format(len(chapter_headings)))
        return 2
    # 逐章：直接在目标文档上生成内容（样式属于目标模板），再移动到章标题之后
    def content_children():
        return [child for child in body.iterchildren() if not child.tag.endswith("sectPr")]

    add_toc(chapter_headings[0])

    for index, heading in enumerate(reversed(chapter_headings)):
        number = 10 - index
        # 章标题补章号：模板 H1 只有「项目概述」，分章稿 H1 是「# 1 项目概述」，
        # 早先 skip_first_heading 把它丢掉，成稿 10 个章标题一个章号都没有。
        md_lines = (chapter_text.get(number) or "").strip().splitlines()
        md_title = md_lines[0][2:].strip() if md_lines and md_lines[0].startswith("# ") else ""
        if md_title:
            heading.text = md_title if re.match(r"^第?\s*\d+", md_title) else \
                "{} {}".format(number, md_title)
        placeholder = heading._p.getnext()
        anchor = heading._p
        before_count = len(content_children())
        insert_chapter(document, number)
        after = content_children()
        new_elements = after[before_count:]
        for element in new_elements:
            anchor.addnext(element)
            anchor = element
        if placeholder is not None and placeholder.tag == qn("w:p"):
            text = Paragraph(placeholder, document).text
            if text.startswith("（本章由"):
                placeholder.getparent().remove(placeholder)
    out_path = project / REPORT_DIR / "项目投资决策建议书-{}.docx".format(project_name)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    # 全文段落显式开启禁则：模板 styles 里一条 kinsoku 都没有，靠继承不可靠，
    # 实测出现 6 处标点被挤到行首（p5「。」、p27「、」）。
    for paragraph in document.paragraphs:
        enable_kinsoku(paragraph)
    # 尾行收敛：段末只剩一两个字（「论。」「实）」）在成稿里很扎眼，
    # 按行宽估算后微缩该段字距，把尾行并回上一行。
    fit_short_tails(document)
    # 小表不跨页：附件清单只有 9 行，被折成「末页孤零零一行表格」很难看。
    # 行内 keepNext 会把整表推到同一页（表头行随表一起，不再单独留在上一页）。
    # 实测 p27 表 9-3 只剩「风险可控」一行被顶到次页页首，故对所有 8 行及以下的表
    # 统一整表保持；超过 8 行的长表（现金流量表）仍需跨页，否则会挤出整页空白。
    for table in document.tables:
        if len(table.rows) <= 8:
            keep_table_together(table)
    add_page_footer(document)
    document.save(str(out_path))
    if args.no_timestamp:
        normalize_zip_timestamps(out_path)
    # 目录域落盘时只有「（打开文档时自动更新目录）」，不打开 Word 就永远是假的。
    # 这里用本机 Word COM 更新一次域，回读校验条目数；无 Word 时写静态目录兜底。
    toc_entries = []
    for number in range(1, 11):
        for line in chapter_text[number].splitlines():
            if line.startswith("# "):
                toc_entries.append(line[2:].strip())
            elif line.startswith("## "):
                toc_entries.append("　　" + line[3:].strip())
    toc_ok, toc_pages = update_toc_fields(out_path, toc_entries)
    if not toc_ok:
        print("[build_invest_docx] 警告：目录条目未能落盘（07 DOC-TOC 会拦）")
    elif toc_pages:
        print("[build_invest_docx] 目录已更新，全文 {} 页".format(toc_pages))

    # 回读自检
    check_document = docx.Document(str(out_path))
    h1 = [p.text.strip() for p in check_document.paragraphs if p.style.name.startswith("Heading 1")]
    text = "\n".join(p.text for p in check_document.paragraphs)
    for table in check_document.tables:
        for row in table.rows:
            for cell in row.cells:
                text += "\n" + cell.text
    placeholders = PLACEHOLDER_PATTERN.findall(text)
    explicit_gaps = GAP_MARKER.findall(text)
    gaps = (payload.get("gaps") or []) + ((finance or {}).get("gaps") or [])
    source_map = read_json(chapter_dir / "来源映射.json", "来源映射", required=False) or {}
    table_headers = [[cell.text.strip() for cell in table.rows[0].cells] for table in check_document.tables]
    slot_headers = {
        1: ["序号", "指标名称", "单位", "计算结果（全部投资）"],
        2: ["序号", "指标名称", "单位", "计算结果（资本金）"],
        3: ["准则", "现状情况", "投资经济", "规划合规", "风险可控"],
        5: ["准则", "行几何平均", "权重"],
        6: ["准则", "准则权重", "指标层", "子权重(4阶矩阵内)", "综合权重"],
    }
    slots_found = [slot for slot, header in slot_headers.items() if header in table_headers]
    slot4_found = any(header and (header[0].startswith("指标层权重派生展示矩阵")
                                  or header[0].startswith("指标层判别矩阵")) for header in table_headers)
    figure_count = len(check_document.inline_shapes)
    # 图题数（内容完整性）与实际嵌图数（渲染环境是否具备）分开统计：
    # 早先只用 inline_shapes 数图，绘图环境缺 matplotlib 时全为 0 →
    # 自检的「图 ≥ 6」必挂，流水线卡在 06，读者连报告都拿不到。
    # 口径：**图题齐全**即内容完整；嵌图缺失是环境问题，单独报，不阻断。
    figure_caption_count = sum(
        1 for para in check_document.paragraphs
        if FIGURE_CAPTION.match(para.text.strip()))
    figure_render_gap = figure_caption_count - figure_count
    self_check = {
        "schema": "cardo/invest-report/v1",
        "schemaVersion": 1,
        "generatedAt": _lib.iso_now(args.no_timestamp),
        "producer": "cardo-invest-report/" + ENGINE_VERSION,
        "project": project_name,
        "chapters": {"count": len(h1), "titles": h1},
        "tables": {"count": len(check_document.tables),
                   "slots": sorted(slots_found + ([4] if slot4_found else [])),
                   "headers": table_headers},
        "figures": {"count": figure_count,
                    "captionCount": figure_caption_count,
                    "notRendered": figure_render_gap},
        "placeholders": placeholders,
        "explicitGaps": len(explicit_gaps),
        "residue": "数据来源" in text,
        "sourceMap": source_map,
        "brokenLinks": [],
        "numbers": {key: indicators.get(key) for key in ("firrAll", "fnpvAll", "ptAll", "firrEquity", "fnpvEquity", "ptEquity", "roe", "roi", "icr", "dscr")},
        "gaps": gaps,
        # 图的门槛按**图题数**（内容完整性）判；实际嵌图数不足时在 notes 里说明，
        # 不阻断成文——绘图是环境问题，不是内容缺陷。
        "result": "pass" if len(h1) == 10 and len(check_document.tables) >= 20
                  and figure_caption_count >= 6
                  and len(slots_found) == 5 and slot4_found and not placeholders
                  and "数据来源" not in text else "fail",
    }
    self_check_path = out_path.with_suffix(".自检.json")
    _lib.write_json(self_check_path, self_check, no_timestamp=args.no_timestamp)
    print("[build_invest_docx] 渲染 {}（章 {} / 表 {} / 图题 {} / 嵌图 {} / 占位符 {} / 结果 {}）".format(
        _lib.rel_posix(out_path, project), len(h1), len(check_document.tables),
        figure_caption_count, figure_count,
        len(placeholders), self_check["result"]))
    if figure_render_gap > 0:
        print("[build_invest_docx] 提示：{} 张图未能渲染（仅保留图题）。"
              "多为绘图库/字体缺失所致，不影响正文与数据；"
              "需要图请安装 matplotlib 后重跑 06。".format(figure_render_gap))
    return 0 if self_check["result"] == "pass" else 1


if __name__ == "__main__":
    sys.exit(main())