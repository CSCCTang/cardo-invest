#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""check_crossnum.py -- 跨章数字一致性校验（A-07）。

契约要求：「1.3 的六个数字能否被单独摘出、且与第七/八/九章逐字相同」
（章节契约.yaml 第 1 章 Q2；acceptance 里 A-07 为 fail 级）。
但**代码里从来没有实现过**——1.3 写错、7/8/9 章写对，读者会看到全书两套数字，
而流水线照样全绿。这是投资决策报告里最伤的一类错：电梯陈述与财务章对不上。

做法：
  ① 从**引擎 JSON**（唯一数据源）取核心指标，生成每个指标在正文里可能的写法；
  ② 找出 1.3 节与 7/8/9 章；
  ③ 同一指标在不同章出现**不同的数值** → fail（真矛盾）；
     某章根本没提 → warn（漏写，不是矛盾）。

数值比较前统一归一（数值归一 + 单位别名），避免「9.57%」与「0.0957」
这种同值不同形被判成不一致（见 _lib.canon_number）。

退出码：0 无矛盾；1 有跨章矛盾。
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import _lib  # noqa: E402

# 锚点小节 → 需要核对的章
CROSS_GROUPS = [("1.3", ["7", "8", "9"])]

# 指标 → (显示名, 单位, 是否百分比)
INDICATORS = [
    ("investment", "总投资", "万元", False),
    ("firrAll", "全投资财务内部收益率", "%", True),
    ("fnpvAll", "全投资财务净现值", "万元", False),
    ("ptAll", "全投资投资回收期", "年", False),
    ("firrEquity", "资本金财务内部收益率", "%", True),
    ("fnpvEquity", "资本金财务净现值", "万元", False),
    ("ptEquity", "资本金投资回收期", "年", False),
    ("roe", "资本金净利润率", "%", True),
    ("roi", "总资产净利润率", "%", True),
    ("icr", "利息备付率", "倍", False),
    ("dscr", "偿债备付率", "倍", False),
]


def _variants(value, is_pct: bool) -> list[str]:
    """该指标在正文里可能的写法（供子串匹配）。

    FIRR 0.0957 → ["9.57", "9.57%", "0.0957"]
    """
    out = []
    try:
        dec = _lib.dec(value)
    except Exception:
        return out
    if is_pct:
        pct = _lib.dec(dec * 100)
        for text in (_lib.canon_number(pct), _lib.canon_number(pct) + "%"):
            if text and text not in out:
                out.append(text)
        raw = _lib.canon_number(dec)
        if raw and raw not in out:
            out.append(raw)
    else:
        text = _lib.canon_number(dec)
        if text and text not in out:
            out.append(text)
        try:
            grouped = "{:,.2f}".format(dec)
            if grouped not in out:
                out.append(grouped)
        except Exception:
            pass
    return out


def _all_variants(indicators: dict) -> dict:
    """{key: [写法…]}"""
    out = {}
    for key, _label, _unit, is_pct in INDICATORS:
        value = indicators.get(key)
        if value is None:
            continue
        variants = _variants(value, is_pct)
        if variants:
            out[key] = variants
    return out


def _section_body(text: str, number: str) -> str:
    """取指定编号小节的正文。"""
    lines = text.splitlines()
    out, capture, level = [], False, 0
    for ln in lines:
        m = re.match(r"^(#{1,6})\s+(.*)$", ln.strip())
        if m:
            depth = len(m.group(1))
            title = m.group(2).strip()
            if not capture and (title.startswith(number + " ") or title == number
                                or title.startswith(number + ".")):
                capture = True
                level = depth
                out.append(ln)
                continue
            if capture:
                if depth <= level:
                    break
                out.append(ln)
                continue
        if capture:
            out.append(ln)
    return "\n".join(out)


def _chapter_bodies(ch_dir: Path) -> dict:
    """{章号: 全文}"""
    bodies = {}
    for f in sorted(ch_dir.glob("*.md")):
        text = f.read_text(encoding="utf-8")
        m = re.match(r"^#\s*(\d+)\s", text.strip())
        if m:
            bodies[m.group(1)] = text
    return bodies


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="跨章数字一致性校验（A-07）")
    ap.add_argument("--project-dir", required=True)
    ap.add_argument("--chapters", default="", help="分章目录（默认 05_过程/投资决策/分章）")
    ap.add_argument("--engine-json", default="",
                    help="引擎 JSON（默认 05_过程/投资决策/财务测算.json）")
    ap.add_argument("--out", default="", help="检查报告 JSON 输出路径")
    args = ap.parse_args(argv)

    project = Path(args.project_dir)
    ch_dir = Path(args.chapters) if args.chapters else project / "05_过程" / "投资决策" / "分章"
    eng_path = (Path(args.engine_json) if args.engine_json
                else project / "05_过程" / "投资决策" / "财务测算.json")
    if not ch_dir.is_dir() or not eng_path.is_file():
        print("[跨章数字] 缺分章稿或引擎 JSON，跳过")
        return 0

    try:
        engine = json.loads(eng_path.read_text(encoding="utf-8"))
    except Exception as exc:
        print("[跨章数字] 引擎 JSON 读不了：{}".format(exc))
        return 0
    indicators = engine.get("indicators") or {}
    variants_by_key = _all_variants(indicators)
    if not variants_by_key:
        print("[跨章数字] 引擎 JSON 里没有 indicators，跳过")
        return 0

    chapters = _chapter_bodies(ch_dir)
    conflicts, missing = [], []
    label_of = {key: label for key, label, _u, _p in INDICATORS}

    for anchor, targets in CROSS_GROUPS:
        anchor_body = ""
        text1 = chapters.get("1")
        if text1:
            anchor_body = _section_body(text1, anchor)
        if not anchor_body.strip():
            continue
        for key, variants in variants_by_key.items():
            if not any(v in anchor_body for v in variants):
                continue      # 1.3 没写这个指标，不构成矛盾
            for t in targets:
                text = chapters.get(t)
                if not text:
                    continue
                # 矛盾有两种：
                # ① 同名指标在该章写了**别的值**（如 1.3 写 FIRR 9.57%，第 8 章写 11.20%）
                # ② 该章写了**另一个指标**的值，而本指标没写（如第 8 章只写资本金 FIRR）
                others = []
                # ① 同指标的不同写法：从正文里抓「指标名附近的百分数/金额」
                label = label_of.get(key, key)
                for m in re.finditer(
                        re.escape(label) + r"[^\n]{0,24}?(\d+(?:\.\d+)?)\s*%?", text):
                    val = m.group(1)
                    if val in variants or val == _lib.canon_number(indicators.get(key)):
                        continue
                    others.append((label + "（该章写法）", val))
                # ② 其它指标的值
                for other_key, other_variants in variants_by_key.items():
                    if other_key == key:
                        continue
                    for ov in other_variants:
                        if ov in text and ov not in variants:
                            others.append((label_of.get(other_key, other_key), ov))
                if others:
                    conflicts.append({
                        "anchor": anchor, "indicator": label_of.get(key, key), "key": key,
                        "expected": variants[0], "chapter": t,
                        "conflictingWith": others[:3],
                    })
                elif not any(v in text for v in variants):
                    missing.append({"anchor": anchor, "indicator": label_of.get(key, key),
                                    "chapter": t, "expected": variants[0]})

    payload = {"schema": "cardo/invest-crossnum-check/v1",
               "ok": not conflicts,
               "conflicts": conflicts, "missing": missing}
    if args.out:
        Path(args.out).parent.mkdir(parents=True, exist_ok=True)
        Path(args.out).write_text(json.dumps(payload, ensure_ascii=False, indent=2),
                                  encoding="utf-8")

    print("[跨章数字] 1.3 与 7/8/9 章：矛盾 {} 处、漏写 {} 处".format(
        len(conflicts), len(missing)))
    for c in conflicts[:8]:
        print("   冲突：1.3 的「{}」= {}，但第 {} 章写了 {}".format(
            c["indicator"], c["expected"], c["chapter"],
            "、".join("{}={}".format(lbl, val) for lbl, val in c["conflictingWith"])))
    for m in missing[:6]:
        print("   漏写：第 {} 章未出现 1.3 的「{}」= {}".format(
            m["chapter"], m["indicator"], m["expected"]))
    if conflicts:
        print("   跨章数字矛盾属 fail 级：同一指标在全书必须是同一个值"
              "（口径以 财务测算.json 为准）。")
    return 1 if conflicts else 0


if __name__ == "__main__":
    sys.exit(main())
