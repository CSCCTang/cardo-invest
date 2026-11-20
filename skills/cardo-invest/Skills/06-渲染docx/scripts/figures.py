"""PIL 绘图：为建议书生成 6 张必备图（不引入 matplotlib/numpy 依赖）。

输入：05_过程/投资决策/{财务输入.json,投资估算.json,财务测算.json,方案比选.json}
输出：05_过程/投资决策/图形/图-<章>-<序>.png（确定性渲染，无时间戳）
"""
from __future__ import annotations

import json
import math
import os
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

FONT_CANDIDATES = [
    r"C:\Windows\Fonts\simhei.ttf",
    r"C:\Windows\Fonts\msyh.ttc",
    r"C:\Windows\Fonts\msyh.ttf",
]
PALETTE = ["#2F6B9A", "#D98B4A", "#5C8A6A", "#8A6BA8", "#B45B5B", "#6E7B8B"]
BG = "#FFFFFF"
INK = "#23303B"
GRID = "#D8DEE4"


def _font(size):
    for candidate in FONT_CANDIDATES:
        if os.path.isfile(candidate):
            try:
                return ImageFont.truetype(candidate, size)
            except OSError:
                continue
    return ImageFont.load_default()


def _canvas(width=1600, height=900):
    image = Image.new("RGB", (width, height), BG)
    return image, ImageDraw.Draw(image)


def _text(draw, xy, content, size=24, fill=INK, anchor="la"):
    draw.text(xy, str(content), font=_font(size), fill=fill, anchor=anchor)


def _axes(draw, box, y_min, y_max, labels, title, unit=""):
    left, top, right, bottom = box
    draw.rectangle([left, top, right, bottom], outline=GRID, width=2)
    span = float(y_max - y_min) or 1.0
    for step in range(5):
        value = y_min + span * step / 4.0
        y = bottom - (bottom - top) * step / 4.0
        draw.line([left, y, right, y], fill=GRID, width=1)
        _text(draw, (left - 12, y), "{:.2f}".format(value) if abs(y_max) < 10 else "{:,.0f}".format(value),
              16, "#6E7B8B", "rm")
    # 图题只由 docx 的「图 X-Y …」题注承担：图里再画一遍会与题注重复（实测 3 张重题）。
    if unit:
        _text(draw, (left - 70, top - 34), unit, 18, "#6E7B8B", "la")
    return left, top, right, bottom


def _line_chart(path, series, title, x_labels, unit="万元"):
    width, height = 1600, 900
    image, draw = _canvas(width, height)
    values = [value for _, points in series for value in points]
    y_min, y_max = min(values), max(values)
    pad = (y_max - y_min) * 0.08 or 1
    left, top, right, bottom = 150, 130, width - 60, height - 120
    _axes(draw, (left, top, right, bottom), y_min - pad, y_max + pad, x_labels, title, unit)
    count = max(len(x_labels) - 1, 1)
    for index, (name, points) in enumerate(series):
        color = PALETTE[index % len(PALETTE)]
        coords = []
        for position, value in enumerate(points):
            x = left + (right - left) * position / count
            y = bottom - (bottom - top) * (value - (y_min - pad)) / ((y_max + pad) - (y_min - pad))
            coords.append((x, y))
        draw.line(coords, fill=color, width=4)
        for x, y in coords[::max(len(coords) // 12, 1)]:
            draw.ellipse([x - 4, y - 4, x + 4, y + 4], fill=color)
        _text(draw, (left + 20, top + 16 + index * 32), "— " + name, 20, color)
    for index in range(0, len(x_labels), max(len(x_labels) // 8, 1)):
        x = left + (right - left) * index / count
        _text(draw, (x, bottom + 14), x_labels[index], 16, "#6E7B8B", "ma")
    image.save(path, "PNG")


def _bar_chart(path, labels, values, title, unit="", horizontal=False, digits=2):
    # 横向条形图有 16 个指标，1600px 宽渲染后缩到 15cm 版心，字只剩 4pt 左右；
    # 改窄画布 + 提高行高，保证打印出来指标名可读（实测约 8pt）。
    if horizontal:
        width, height = 1000, max(760, 54 * len(labels) + 170)
    else:
        width, height = 1600, 900
    image, draw = _canvas(width, height)
    v_max = max(values + [0]) or 1
    v_min = min(values + [0])
    if horizontal:
        left, right = 330, width - 60
        top, bottom = 110, height - 70
        draw.rectangle([left, top, left, bottom], outline=GRID, width=2)
        count = len(values)
        for index, (label, value) in enumerate(zip(labels, values)):
            y = top + (bottom - top) * (index + 0.5) / count
            span = right - left
            zero = left + span * (0 - v_min) / ((v_max - v_min) or 1)
            x = left + span * (value - v_min) / ((v_max - v_min) or 1)
            draw.rectangle([min(zero, x), y - 18, max(zero, x), y + 18], fill=PALETTE[index % len(PALETTE)])
            _text(draw, (left - 12, y), label, 20, INK, "rm")
            _text(draw, (max(zero, x) + 10, y), "{:.{d}f}{u}".format(value, d=digits, u=unit), 17, "#6E7B8B", "lm")
        image.save(path, "PNG")
        return
    left, top, right, bottom = 130, 110, width - 60, height - 130
    _axes(draw, (left, top, right, bottom), 0, v_max * 1.15, [], title, unit)
    count = len(values)
    slot = (right - left) / max(count, 1)
    for index, (label, value) in enumerate(zip(labels, values)):
        x0 = left + slot * index + slot * 0.18
        x1 = left + slot * (index + 1) - slot * 0.18
        y = bottom - (bottom - top) * max(value, 0) / (v_max * 1.15)
        draw.rectangle([x0, y, x1, bottom], fill=PALETTE[index % len(PALETTE)])
        _text(draw, ((x0 + x1) / 2, y - 12), "{:.{d}f}{u}".format(value, d=digits, u=unit), 18, INK, "mb")
        _text(draw, ((x0 + x1) / 2, bottom + 14), label, 16, "#6E7B8B", "ma")
    image.save(path, "PNG")


def _org(path, layers, title):
    width, height = 1600, 900
    image, draw = _canvas(width, height)
    y = 140
    for index, items in enumerate(layers):
        count = len(items)
        box_w = min(360, int((width - 200) / max(count, 1)) - 30)
        box_h = 100
        total = count * box_w + (count - 1) * 30
        x = (width - total) / 2
        for text in items:
            box = [x, y, x + box_w, y + box_h]
            draw.rectangle(box, fill="#EAF1F7", outline=PALETTE[index % len(PALETTE)], width=3)
            _text(draw, ((box[0] + box[2]) / 2, (box[1] + box[3]) / 2), text, 20, INK, "mm")
            x += box_w + 30
        if index < len(layers) - 1:
            draw.line([width / 2, y + box_h, width / 2, y + box_h + 50], fill="#6E7B8B", width=4)
        y += box_h + 50
    image.save(path, "PNG")


def _gantt(path, tasks, title):
    width, height = 1600, 900
    image, draw = _canvas(width, height)
    left, right = 340, width - 80
    top, bottom = 150, height - 120
    # 横轴按实际建设期取整到年：早先固定 36 个月，2 年项目只占 1/3 图宽，右侧大段空白
    last_month = max([start + span for _name, start, span in tasks] or [12])
    months = max(12, int(math.ceil(last_month / 12.0)) * 12)
    for index, (name, start, span) in enumerate(tasks):
        y = top + index * 130
        draw.rectangle([left, y, right, y + 70], outline=GRID, width=2)
        x0 = left + (right - left) * start / months
        x1 = left + (right - left) * (start + span) / months
        draw.rectangle([x0, y + 8, x1, y + 62], fill=PALETTE[index % len(PALETTE)])
        _text(draw, (left - 16, y + 35), name, 20, INK, "rm")
        _text(draw, ((x0 + x1) / 2, y + 35), "{} 个月".format(span), 18, "#FFFFFF", "mm")
    for index in range(0, months + 1, 6):
        x = left + (right - left) * index / months
        draw.line([x, top, x, bottom], fill=GRID, width=1)
        # 刻度每 6 个月一格，但年份只在 1 月标一次，否则 2026/2026、2027/2027 重复出现。
        if index % 12 == 0:
            _text(draw, (x, bottom + 12), "{}".format(2026 + index // 12), 16, "#6E7B8B", "ma")
    image.save(path, "PNG")


def _find_input(workspace, name):
    """在 输入文件/ 下按文件名找上游成果（正射、掩膜、轮廓）。"""
    root = Path(workspace) / "输入文件"
    if not root.is_dir():
        return None
    hits = sorted(root.rglob(name))
    return hits[0] if hits else None


def _dashed_line(draw, p0, p1, color, width=4, dash=14, gap=9):
    x0, y0 = p0
    x1, y1 = p1
    dx, dy = x1 - x0, y1 - y0
    dist = math.hypot(dx, dy)
    if dist <= 0:
        return
    ux, uy = dx / dist, dy / dist
    pos = 0.0
    while pos < dist:
        end = min(pos + dash, dist)
        draw.line([(x0 + ux * pos, y0 + uy * pos), (x0 + ux * end, y0 + uy * end)],
                  fill=color, width=width)
        pos = end + gap


def _site_map(path, workspace):
    """图 3-1：用 OSGB 成果出空间现状图（正射＋地类＋建筑轮廓＋范围＋图例/指北针/比例尺）。

    数据源与 Status 图 2-1 同一套：ortho_northup.png / ortho_georef.json /
    class_rgba.png / class_georef.json / buildings-outline.geojson。
    缺件返回 None（契约 F3-1：不自绘无来源示意图）。
    """
    ortho = _find_input(workspace, "ortho_northup.png")
    class_png = _find_input(workspace, "class_rgba.png")
    georef = _find_input(workspace, "ortho_georef.json")
    bld = _find_input(workspace, "buildings-outline.geojson")
    if not all((ortho, class_png, georef, bld)):
        return None
    ref = json.loads(Path(georef).read_text(encoding="utf-8"))
    minx, miny, maxx, maxy = [float(x) for x in ref["extentLocalM"]]
    res = float(ref.get("resM") or 0.25)
    crs = ref.get("crs") or {}
    epsg = crs.get("epsg") or 4544
    nx = int((ref.get("grid") or {}).get("nx") or round((maxx - minx) / res))
    ny = int((ref.get("grid") or {}).get("ny") or round((maxy - miny) / res))

    ortho_im = Image.open(ortho).convert("RGB")
    class_im = Image.open(class_png).convert("RGBA")
    if ortho_im.size != (nx, ny):
        ortho_im = ortho_im.resize((nx, ny), Image.LANCZOS)
    if class_im.size != (nx, ny):
        class_im = class_im.resize((nx, ny), Image.LANCZOS)

    width, height = 1800, 1300
    l, t, r, b = 90, 110, 1310, 1000
    box_w, box_h = r - l, b - t
    mpp = max((maxx - minx) / box_w, (maxy - miny) / box_h)
    draw_w = (maxx - minx) / mpp
    draw_h = (maxy - miny) / mpp
    ox = l + (box_w - draw_w) / 2
    oy = t + (box_h - draw_h) / 2

    map_w, map_h = int(round(draw_w)), int(round(draw_h))
    alpha = class_im.split()[-1].point(lambda a: 255 if a > 8 else 0)
    base = Image.new("RGB", (map_w, map_h), (255, 255, 255))
    base.paste(ortho_im.resize((map_w, map_h), Image.LANCZOS),
               (0, 0), alpha.resize((map_w, map_h), Image.LANCZOS))
    # 正射是模型纹理：压低饱和度、提亮后再叠地类，避免纹理拉伸干扰判读
    from PIL import ImageEnhance
    base = ImageEnhance.Color(base).enhance(0.55)
    base = ImageEnhance.Brightness(base).enhance(1.18)
    land = class_im.resize((map_w, map_h), Image.LANCZOS)
    land.putalpha(land.split()[-1].point(lambda a: int(a * 0.55)))
    base = Image.alpha_composite(base.convert("RGBA"), land).convert("RGB")

    canvas = Image.new("RGB", (width, height), BG)
    canvas.paste(base, (int(ox), int(oy)))
    draw = ImageDraw.Draw(canvas)

    def px(x, y):
        return (ox + (x - minx) / mpp, oy + (maxy - y) / mpp)

    try:
        fc = json.loads(Path(bld).read_text(encoding="utf-8"))
        for feature in fc.get("features") or []:
            geom = feature.get("geometry") or {}
            if geom.get("type") != "Polygon":
                continue
            rings = geom.get("coordinates") or []
            if not rings:
                continue
            pts = [px(float(x), float(y)) for x, y in rings[0]]
            if len(pts) >= 3:
                draw.line(pts + [pts[0]], fill=(32, 48, 96), width=3)
    except (OSError, ValueError, TypeError):
        pass

    # 项目范围（模型覆盖范围）虚线框＋内外框
    for p0, p1 in (((ox, oy), (ox + draw_w, oy)), ((ox + draw_w, oy), (ox + draw_w, oy + draw_h)),
                   ((ox + draw_w, oy + draw_h), (ox, oy + draw_h)), ((ox, oy + draw_h), (ox, oy))):
        _dashed_line(draw, p0, p1, (176, 32, 48), 4)
    draw.rectangle([l, t, r, b], outline=(90, 90, 90), width=2)

    # 范围说明（图题由 docx 题注承担，图内不重复）
    _text(draw, (92, 54), "项目范围：模型覆盖范围（红色虚线）；图幅约 {:.2f} 米 × {:.2f} 米".format(
        maxx - minx, maxy - miny), 24, "#3C3C3C")

    # 指北针
    nx0, ny0 = 1240, 74
    draw.polygon([(nx0, ny0 - 34), (nx0 - 18, ny0 + 18), (nx0, ny0 + 6), (nx0 + 18, ny0 + 18)],
                 fill=(30, 30, 30))
    _text(draw, (nx0 - 8, ny0 + 24), "N", 26, "#1E1E1E")

    # 比例尺 0—30 米（三段，每段 10 米）
    sx, sy = ox + 24, oy + draw_h - 54
    seg = 10 / mpp
    for index in range(3):
        draw.rectangle([sx + index * seg, sy, sx + (index + 1) * seg, sy + 14],
                       fill=(30, 30, 30) if index % 2 == 0 else (255, 255, 255),
                       outline=(30, 30, 30))
        _text(draw, (sx + index * seg - 8, sy + 20), str(index * 10), 22, "#1E1E1E")
    _text(draw, (sx + 3 * seg - 14, sy + 20), "30", 22, "#1E1E1E")
    _text(draw, (sx + 3 * seg + 18, sy + 20), "m", 22, "#1E1E1E")

    # 图例
    lx, ly = 1350, 190
    _text(draw, (lx, ly), "图  例", 30, "#141414")
    items = [
        ((230, 126, 34), "建筑屋面（地类识别）"),
        ((158, 158, 158), "道路与硬化表面"),
        ((76, 175, 80), "绿化与植被"),
        ((255, 255, 255), "模型未覆盖 / 无数据"),
    ]
    for index, (color, label) in enumerate(items):
        yy = ly + 58 + index * 42
        draw.rectangle([lx, yy, lx + 30, yy + 24], fill=color, outline=(120, 120, 120))
        _text(draw, (lx + 42, yy - 2), label, 23, "#282828")
    yy = ly + 58 + len(items) * 42 + 12
    draw.line([(lx, yy + 10), (lx + 30, yy + 10)], fill=(176, 32, 48), width=4)
    _text(draw, (lx + 42, yy - 2), "项目范围（模型覆盖，非红线）", 23, "#282828")
    yy += 42
    draw.line([(lx, yy + 10), (lx + 30, yy + 10)], fill=(32, 48, 96), width=3)
    _text(draw, (lx + 42, yy - 2), "建筑轮廓", 23, "#282828")

    # 坐标系与来源：图下两栏
    notes_left = [
        "坐标系：CGCS2000 / 3 度带高斯—克吕格，中央经线 105°E",
        "（EPSG:{}）；高程基准 EPSG:{}。".format(epsg, crs.get("vertical", "5737")),
        "数据来源：OSGB 低空三维解析成果——正射影像、",
        "地类掩膜、建筑轮廓（同包《现状指标》依据数据）。",
    ]
    notes_right = [
        "影像地面分辨率 {:.2f} 米；底图为模型顶视正射，".format(res),
        "纹理存在拉伸、遮挡与色差，仅用于判读空间关系，",
        "不展示建筑立面与风貌细节。",
        "项目范围为模型覆盖范围，非宗地红线；四至以现状道路中心线为准。",
    ]
    ty = 1030
    for index, line in enumerate(notes_left):
        _text(draw, (90, ty + index * 34), line, 21, "#3C3C3C")
    for index, line in enumerate(notes_right):
        _text(draw, (760, ty + index * 34), line, 21, "#3C3C3C")

    canvas.save(path, "PNG")
    return path


def _series(finance, key):
    values = (((finance or {}).get("cashflows") or {}).get(key) or {}).get("cumulative") or []
    return [float(item.get("value") or 0) for item in values]


def render_figure(fig_id, out_dir, payload=None, finance=None, estimate=None, schemes=None,
                  workspace=None):
    """按图号生成 PNG；数据缺失返回 None。"""
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    path = out_dir / "图-{}.png".format(fig_id.replace("-", "_"))
    if fig_id == "3-1":
        # 空间现状图只认 OSGB 正射/地类/轮廓；缺件不画（契约 F3-1 禁止自绘无来源示意）。
        return _site_map(path, workspace) if workspace else None
    if fig_id == "3-2":
        # 现状图数据只取上游实测样区口径（建筑密度/绿地率/道路占比），
        # 其余硬化 = 100 − 三项；三项缺一就不画（契约 F3-2 禁止自绘无来源示意）。
        upstream = (payload or {}).get("upstream") or {}
        facts = (payload or {}).get("facts") or upstream.get("facts") or {}

        def _fact(key):
            item = facts.get(key) or {}
            try:
                return float(item.get("value"))
            except (TypeError, ValueError):
                return None

        density, green, road = _fact("density"), _fact("greeneryRate"), _fact("roadRate")
        if None in (density, green, road):
            # 财务输入.json 不带 facts 时，退到 05_过程/低空解析/现状指标.json 的 metrics
            # （01 把该文件挂到 upstream.statusMetrics，比率为 0–1）。
            by_name = {str(item.get("name")): item.get("value")
                       for item in (upstream.get("statusMetrics") or []) if isinstance(item, dict)}

            def _pct(name):
                try:
                    return round(float(by_name.get(name)) * 100, 2)
                except (TypeError, ValueError):
                    return None

            density = density if density is not None else _pct("建筑密度")
            green = green if green is not None else _pct("绿化率（纹理覆盖代理）")
            road = road if road is not None else _pct("道路占比（纹理覆盖代理）")
        if None in (density, green, road):
            return None
        other = round(100 - density - green - road, 2)
        _bar_chart(path, ["建筑基底", "绿地", "道路", "其他硬化"],
                   [density, green, road, other],
                   "图 3-2  实测样区现状空间构成（%）", "%", horizontal=True, digits=2)
    elif fig_id == "5-1":
        rows = ((estimate or {}).get("tables") or {}).get("yearPlan", {}).get("rows") or []
        tasks = [("第 {} 年".format(row.get("year")), (int(row.get("year", 1)) - 1) * 12, 12) for row in rows]
        _gantt(path, tasks or [("建设期", 0, 24)], "图 5-1  建设时序横道图")
    elif fig_id == "6-1":
        # C8：组织架构属方案类待业主填写，图只能画 6.1 已写明的职责分工关系，
        # 不预设「项目公司」「专业运营团队」等具体层级与岗位。
        slots = (payload or {}).get("slots") or {}
        org = slots.get("organization")
        if isinstance(org, list):
            org = "；".join(str(item) for item in org if str(item).strip())
        org_text = str(org or "").strip()
        head = org_text.split("；")[0].strip() if org_text else ""
        implementer = "实施主体：{}".format(head) if head else "实施主体（融资、建设管理与资产持有）"
        _org(path, [["政府主管部门（监管）"], [implementer], ["运营单位（招商、租赁、物业与设施维护）"],
                    ["街道、社区与商户（共建共治）"]], "图 6-1  实施与运营职责分工关系图")
    elif fig_id == "8-1":
        all_values = _series(finance, "allInvestment")
        equity_values = _series(finance, "equity")
        if not all_values:
            return None
        labels = ["第 {} 年".format(index + 1) for index in range(len(all_values))]
        _line_chart(path, [("全部投资累计净现金流", all_values), ("资本金累计净现金流", equity_values)],
                    "图 8-1  累计净现金流量曲线", labels, "万元")
    elif fig_id == "8-2":
        factors = ((finance or {}).get("sensitivity") or {}).get("factors") or []
        if not factors:
            return None
        series = []
        labels = []
        for factor in factors:
            rows = factor.get("rows") or []
            labels = [row.get("label") for row in rows]
            series.append((factor.get("name") or factor.get("key"),
                           [float(row.get("firr") or 0) * 100 for row in rows]))
        _line_chart(path, series, "图 8-2  敏感性分析（FIRR，%）", labels, "%")
    elif fig_id == "9-1":
        weights = ((payload or {}).get("ahpWeights") or {}).get("indicatorWeights") or []
        if not weights:
            return None
        ordered = sorted(weights, key=lambda item: -float(item.get("compositeWeight") or 0))
        _bar_chart(path, [item.get("indicator") for item in ordered],
                   [float(item.get("compositeWeight") or 0) for item in ordered],
                   "图 9-1  指标综合权重排序", "", horizontal=True, digits=4)
    elif fig_id == "9-2":
        scores = ((payload or {}).get("ahpScores") or {}).get("schemes") or []
        # 图例编号与第 9 章正文、表 9-5、表 9-6 同一套，且与 AHP 包一致：
        # S1 保护修缮为主 / S2 拆建结合 / S3 综合提升（方案比选原名：功能置换 / 拆除重建 / 微更新）。
        _canon = (("S1", "保护修缮为主", "功能置换"), ("S2", "拆建结合", "拆除重建"),
                  ("S3", "综合提升", "微更新"))
        compare_rows = ((schemes or {}).get("compareTable") or {}).get("rows") or []

        def _label(name):
            text = str(name or "")
            if not text:
                return text
            for scheme_id, canon_name, compare_name in _canon:
                for candidate in (canon_name, compare_name):
                    if text in candidate or candidate in text:
                        return "{} {}".format(scheme_id, canon_name)
            return text

        if scores:
            ordered = sorted(scores, key=lambda item: -float(item.get("composite") or 0))
            _bar_chart(path, [_label(item.get("name")) for item in ordered],
                       [float(item.get("composite") or 0) for item in ordered],
                       "图 9-2  方案综合得分对比（AHP 综合得分，10 分制）", "", digits=4)
        else:
            rows = compare_rows
            if not rows:
                return None
            _bar_chart(path, [_label(row.get("name")) for row in rows],
                       [float(row.get("score") or 0) for row in rows], "图 9-2  方案综合得分对比", "", digits=4)
    else:
        return None
    return path if path.is_file() else None
