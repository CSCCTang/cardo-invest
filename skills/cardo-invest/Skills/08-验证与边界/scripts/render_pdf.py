# -*- coding: utf-8 -*-
"""render_pdf.py -- 成品 PDF：docx → A4 HTML → 无头 Edge 打印 → 目录页码 → 页脚。

为什么不用 Word 转 PDF：交付环境没有 Office，且目录页码/页脚必须在离线条件下可复现。
链路与 CardoStatus 的 pdf_render 同源（已在现状分析 R3 验证）：python-docx 读成稿
→ A4 HTML → Edge headless --print-to-pdf → PyMuPDF 定位标题页码 → 回填目录
→ 二次打印 → 烘页脚（封面不印页码）。

与 Status 版的差异（建议书模板特有的两处）：
  * docx 自带的「目录」是 TOC 样式的一组静态条目、没有页码 —— 全部跳过，由本脚本重排；
  * 封面的标题/项目名是 Normal 段落，按「首个标题之前」整体收进封面容器。

用法：
  python render_pdf.py --docx 建议书.docx --out 建议书.pdf [--title 标题]
"""
from __future__ import annotations

import argparse
import base64
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

try:
    from docx import Document
    from docx.table import Table
    from docx.text.paragraph import Paragraph
except ImportError:
    print("[pdf] 缺少 python-docx", file=sys.stderr)
    raise SystemExit(1)

try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None


def iter_block_items(doc):
    """按文档顺序迭代段落与表格。"""
    from docx.oxml.ns import qn
    body = doc.element.body
    for child in body.iterchildren():
        if child.tag == qn("w:p"):
            yield Paragraph(child, doc)
        elif child.tag == qn("w:tbl"):
            yield Table(child, doc)


def heading_level(p: Paragraph) -> int:
    name = (p.style.name or "").lower()
    if name.startswith("heading"):
        m = re.search(r"(\d+)", name)
        if m:
            return int(m.group(1))
    m = re.search(r"标题\s*(\d+)", p.style.name or "")
    return int(m.group(1)) if m else 0


def esc(s: str) -> str:
    return (s or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def table_html(t: Table) -> str:
    rows = []
    for row in t.rows:
        cells = []
        for cell in row.cells:
            cells.append("<td>" + esc(cell.text.strip()).replace("\n", "<br>") + "</td>")
        rows.append("<tr>" + "".join(cells) + "</tr>")
    if not rows:
        return ""
    return ('<table class="data"><thead>' + rows[0] + "</thead><tbody>"
            + "".join(rows[1:]) + "</tbody></table>")


def _is_bold_caption(p: Paragraph) -> bool:
    """短句全加粗 → 表题/图题的候选（避免表题落在页底）。"""
    text = p.text.strip()
    if not text or len(text) > 40:
        return False
    runs = [r for r in p.runs if r.text.strip()]
    return bool(runs) and all(r.bold for r in runs)


def is_toc_para(p: Paragraph) -> bool:
    """模板自带的目录条目（TOC 1/2/3 样式）与「目录」标题：全部交给本脚本重排。"""
    style = (p.style.name or "")
    return style.upper().startswith("TOC")


def docx_to_html(doc):
    parts, headings = [], []
    cover, in_cover = [], True
    blocks = list(iter_block_items(doc))
    skip = set()
    for i, block in enumerate(blocks):
        if i in skip:
            continue
        if isinstance(block, Table):
            parts.append(table_html(block))
            in_cover = False
            continue
        p = block
        # 图片段落（w:drawing→a:blip）
        try:
            blips = p._p.xpath(".//a:blip/@r:embed")
        except Exception:  # noqa: BLE001
            blips = []
        if blips:
            for rid in blips:
                try:
                    part = doc.part.related_parts[rid]
                except KeyError:
                    continue
                b64 = base64.b64encode(part.blob).decode("ascii")
                parts.append(f'<p class="fig"><img src="data:{part.content_type};base64,{b64}" alt=""/></p>')
            in_cover = False
            continue
        text = p.text.strip()
        if not text:
            continue
        if text.replace("　", "").replace(" ", "") == "目录" or is_toc_para(p):
            continue
        if _is_bold_caption(p) and i + 1 < len(blocks) and isinstance(blocks[i + 1], Table):
            parts.append('<div class="keep"><p class="tcap">' + esc(text) + "</p>"
                         + table_html(blocks[i + 1]) + "</div>")
            skip.add(i + 1)
            in_cover = False
            continue
        lvl = heading_level(p)
        if 1 <= lvl <= 3:
            headings.append((lvl, text))
            parts.append(f'<h{lvl} data-h="{esc(text)}">{esc(text)}</h{lvl}>')
            in_cover = False
        else:
            parts.append("<p>" + esc(text) + "</p>")
            if in_cover:
                cover.append(parts.pop())
    return "\n".join(parts), headings, "\n".join(cover)


CSS = """
@page { size: A4; margin: 22mm 20mm 20mm 20mm; }
* { box-sizing: border-box; }
body { font-family: "Noto Serif CJK SC", "Source Han Serif SC", "SimSun", serif;
       font-size: 11.5pt; line-height: 1.75; color: #1a1a1a; margin: 0; }
h1 { font-family: "Noto Sans CJK SC", "Microsoft YaHei", sans-serif;
     font-size: 18pt; margin: 0 0 14pt; padding-bottom: 6pt;
     border-bottom: 1.2pt solid #333; page-break-before: always; }
h1:first-of-type { page-break-before: auto; }
h2 { font-family: "Noto Sans CJK SC", "Microsoft YaHei", sans-serif;
     font-size: 14pt; margin: 16pt 0 8pt; }
h3 { font-size: 12pt; margin: 12pt 0 6pt; }
p { margin: 0 0 8pt; text-align: justify; text-indent: 2em; }
p.fig { text-align: center; text-indent: 0; margin: 8pt 0 4pt; page-break-inside: avoid; }
p.fig img { width: 100%; max-width: 170mm; }
p.toc-line { text-indent: 0; margin: 3pt 0; display: flex;
             justify-content: space-between; border-bottom: 0.5pt dotted #bbb; }
p.toc-line.lv1 { font-weight: 700; margin-top: 8pt; }
p.toc-line .pg { color: #444; }
.keep { page-break-inside: auto; }
p.tcap { text-indent: 0; font-weight: 700; margin: 10pt 0 2pt; page-break-after: avoid; }
table.data { width: 100%; border-collapse: collapse; margin: 6pt 0 10pt;
             font-size: 9.5pt; page-break-inside: auto; }
table.data thead { display: table-header-group; }
table.data tr { page-break-inside: avoid; }
table.data td { border: 0.5pt solid #999; padding: 3pt 4pt; vertical-align: top;
                word-break: break-word; overflow-wrap: anywhere; }
table.data tr:first-child td { background: #f0f0f0; font-weight: 700; }
.toc-page { page-break-after: always; }
.cover { page-break-after: always; text-align: center; padding-top: 55mm; }
.cover p { text-align: center; text-indent: 0; margin: 0 0 14pt; }
.cover p:first-child { font-family: "Noto Sans CJK SC", "Microsoft YaHei", sans-serif;
                       font-size: 26pt; font-weight: 700; line-height: 1.5; margin-bottom: 22pt; }
.cover p:nth-child(2) { font-size: 13pt; color: #444; }
.toc-title { font-size: 18pt; font-weight: 700; text-align: center; margin: 10mm 0 12pt; }
"""


def build_html(body: str, toc_html: str, cover_html: str = "", title: str = "") -> str:
    cover = f'<div class="cover">{cover_html}</div>' if cover_html else ""
    return ("<!DOCTYPE html><html lang='zh-CN'><head><meta charset='utf-8'>"
            f"<title>{esc(title) or '投资决策建议书'}</title><style>{CSS}</style></head><body>"
            f"{cover}{toc_html}{body}</body></html>")


def toc_block(headings, pages: dict[int, int], sentinel: bool = True) -> str:
    if not headings:
        return ""
    lines = ['<div class="toc-page"><div class="toc-title">目录</div>']
    for i, (lvl, text) in enumerate(headings):
        if lvl > 2:
            continue
        pg = pages.get(i, "")
        lines.append(f'<p class="toc-line lv{lvl}"><span>{esc(text)}</span>'
                     f'<span class="pg">{pg}</span></p>')
    if sentinel:
        lines.append('<div style="font-size:1pt;color:#ffffff">CARDOTOCEND</div>')
    else:
        lines.append("<!--CARDOTOCEND-->")
    lines.append("</div>")
    return "".join(lines)


def find_browser():
    candidates = [
        os.path.join(os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)"),
                     "Microsoft", "Edge", "Application", "msedge.exe"),
        os.path.join(os.environ.get("ProgramFiles", r"C:\Program Files"),
                     "Microsoft", "Edge", "Application", "msedge.exe"),
        shutil.which("msedge.exe"),
        shutil.which("chrome.exe"),
        os.path.join(os.environ.get("ProgramFiles", r"C:\Program Files"),
                     "Google", "Chrome", "Application", "chrome.exe"),
    ]
    for c in candidates:
        if c and os.path.isfile(c):
            return c
    return None


def render(html: str, pdf: Path, work: Path) -> bool:
    browser = find_browser()
    if not browser:
        print("[pdf] 找不到 Edge/Chrome，无法打印 PDF", file=sys.stderr)
        return False
    pdf = pdf.resolve()          # 相对路径会被 Edge 解析到自己的目录，必须绝对
    pdf.parent.mkdir(parents=True, exist_ok=True)
    html_path = (work / "report.html").resolve()
    html_path.write_text(html, encoding="utf-8")
    cmd = [browser, "--headless=new", "--disable-gpu", "--no-pdf-header-footer",
           f"--print-to-pdf={pdf}", html_path.as_uri()]
    try:
        subprocess.run(cmd, check=True, timeout=180,
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except Exception as exc:
        print(f"[pdf] 打印失败：{exc}", file=sys.stderr)
        return False
    if not pdf.exists():
        print(f"[pdf] 打印后未生成 PDF：{pdf}", file=sys.stderr)
        return False
    return True


def heading_pages(pdf: Path, headings) -> dict[int, int]:
    """在 PDF 里逐页找标题文字，返回 {标题序号: 页码}（页码从 1 起）。

    按「行首」匹配而不是整页子串：正文里引用章名（如「必要性与投资匹配……
    见 5.3」或「该口径在「现状详细分析」一章……」）会让整页子串先命中引用句
    所在页，目录页码就跳到引用页而不是标题页。标题在版面上独占一行，
    行首比对既能命中标题，也不会被正文引用误伤（R11 跨包修正）。
    """
    if fitz is None or not pdf.exists():
        return {}
    out = {}
    doc = fitz.open(str(pdf))
    norm = lambda s: re.sub(r"\s+", "", s or "")
    texts = [norm(p.get_text("text")) for p in doc]
    pages = [[norm(line) for line in (p.get_text("text") or "").splitlines()] for p in doc]
    start = 0
    for idx, t in enumerate(texts):
        if "CARDOTOCEND" in t:
            start = idx + 1
    for i, (_lvl, text) in enumerate(headings):
        key = norm(text)
        probe = key[:14]
        if not probe:
            continue
        for idx in range(start, len(pages)):
            if any(line.startswith(probe) for line in pages[idx]):
                out[i] = idx + 1
                break
    doc.close()
    return out


def add_bookmarks(pdf: Path, headings, pages) -> None:
    """按正文标题给 PDF 加书签（R3 P2-17：PDF 只有页脚、没有大纲）。"""
    if fitz is None or not pages:
        return
    doc = fitz.open(str(pdf))
    toc = []
    for index, (level, text) in enumerate(headings):
        page = pages.get(index)
        if page and str(text).strip():
            toc.append([max(1, int(level)), str(text).strip(), int(page)])
    if toc:
        doc.set_toc(toc)
    doc.saveIncr()
    doc.close()


def bake_footer(pdf: Path) -> None:
    """页脚：第 N 页 / 共 M 页（烘进 PDF，任何阅读器一致）。封面不印页码。"""
    if fitz is None:
        return
    doc = fitz.open(str(pdf))
    total = doc.page_count
    for idx, page in enumerate(doc):
        if idx == 0:
            continue
        w, h = page.rect.width, page.rect.height
        label = f"第 {idx + 1} 页 / 共 {total} 页"
        page.insert_text((w / 2 - 42, h - 26), label,
                         fontsize=9, fontname="china-s", color=(0.35, 0.35, 0.35))
    doc.saveIncr()
    doc.close()


def patch_app_pages(docx: Path, total: int) -> None:
    """把 docProps/app.xml 的 Pages 改成实际页数。"""
    if total <= 0:
        return
    try:
        with zipfile.ZipFile(str(docx)) as z:
            names = z.namelist()
            if "docProps/app.xml" not in names:
                return
            parts = {n: z.read(n) for n in names}
    except Exception:
        return
    xml = parts["docProps/app.xml"].decode("utf-8")
    new, n = re.subn(r"<Pages>\d+</Pages>", f"<Pages>{total}</Pages>", xml, count=1)
    if n == 0 or new == xml:
        return
    parts["docProps/app.xml"] = new.encode("utf-8")
    tmp = str(docx) + f".tmp-{os.getpid()}"
    try:
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
            for name, data in parts.items():
                z.writestr(name, data)
        os.replace(tmp, str(docx))
    except Exception:
        try:
            if os.path.exists(tmp):
                os.remove(tmp)
        except OSError:
            pass


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--docx", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--title", default="")
    args = ap.parse_args()

    docx, out = Path(args.docx), Path(args.out)
    if not docx.exists():
        print(f"[pdf] docx 不存在：{docx}", file=sys.stderr)
        return 2
    out.parent.mkdir(parents=True, exist_ok=True)

    doc = Document(str(docx))
    body, headings, cover = docx_to_html(doc)

    with tempfile.TemporaryDirectory(prefix="cardo_invest_pdf_") as tmp:
        work = Path(tmp)
        probe = work / "probe.pdf"
        if not render(build_html(body, "", cover, args.title), probe, work):
            return 1
        pages = heading_pages(probe, headings)
        toc_pdf = work / "toc.pdf"
        toc = toc_block(headings, pages)
        if not render(build_html(body, toc, cover, args.title), toc_pdf, work):
            return 1
        pages2 = heading_pages(toc_pdf, headings)
        toc = toc_block(headings, pages2, sentinel=False)
        if not render(build_html(body, toc, cover, args.title), out, work):
            return 1

    bake_footer(out)
    add_bookmarks(out, headings, pages2)
    if fitz is not None:
        d = fitz.open(str(out))
        total = d.page_count
        d.close()
        patch_app_pages(docx, total)
        print(f"[pdf] ok：{out.name}（{total} 页）")
    else:
        print(f"[pdf] ok：{out.name}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
