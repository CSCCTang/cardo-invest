#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""export_trace_sample.py -- 数字溯源抽样表。

以 05 步产出的 `依据对照表.json` 为准，给出两级溯源口径：
  · 值级：数值绑到上游 JSON 的唯一字段（status=可解析 / 标准值）；
  · 段级：数值未绑字段，但所在段落有来源行（指向上游报告或文本）。
两者都能回到原始文件；表里对每个抽样值写明「节 → 依据对照表指向 → 原始文件」。

输出：成果文件/内部件/数字溯源表/数字溯源抽样.md
退出码：0 生成成功；2 找不到依据对照表。
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

LEVEL_VALUE = ("可解析", "标准值")


def load_rows(workspace: Path) -> list[dict]:
    path = workspace / "过程文件" / "05-章节成文" / "分章" / "依据对照表.json"
    if not path.is_file():
        return []
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return []
    return data.get("rows") or []


def trace_levels(rows: list[dict]) -> dict:
    counts = {"值级": 0, "标准值": 0, "结构值": 0, "段级": 0, "无源": 0}
    for row in rows:
        has_source = bool(row.get("sources"))
        for item in row.get("valueRefs") or []:
            status = item.get("status")
            if status == "可解析":
                counts["值级"] += 1
            elif status == "标准值":
                counts["标准值"] += 1
            elif status == "结构值":
                counts["结构值"] += 1
            elif has_source:
                counts["段级"] += 1
            else:
                counts["无源"] += 1
    return counts


def pick_samples(rows: list[dict], limit: int) -> list[dict]:
    samples, seen_sections = [], set()
    for row in rows:
        for item in row.get("valueRefs") or []:
            if item.get("status") != "可解析":
                continue
            key = (row.get("section"), item.get("number"))
            if key in seen_sections:
                continue
            seen_sections.add(key)
            samples.append({
                "chapter": row.get("chapter"),
                "section": row.get("section"),
                "number": item.get("number"),
                "ref": item.get("ref"),
                "path": item.get("path"),
            })
            break                       # 每节取一个值级样本，覆盖尽量多节
        if len(samples) >= limit:
            break
    return samples


def render(counts: dict, samples: list[dict]) -> str:
    total = sum(counts.values())
    lines = [
        "# 数字溯源抽样表", "",
        "口径：以 `依据对照表.json` 为准。值级＝数值绑到上游唯一字段；"
        "段级＝数值未绑字段，但所在段落有来源行。两者都能回到原始文件。", "",
        "## 一、总体覆盖", "",
        "| 层级 | 数量 | 占比 |", "|---|---:|---:|",
    ]
    for key in ("值级", "标准值", "结构值", "段级", "无源"):
        count = counts[key]
        pct = "—" if not total else "{:.1f}%".format(100.0 * count / total)
        lines.append("| {} | {} | {} |".format(key, count, pct))
    lines += ["| **合计** | **{}** | **100%** |".format(total), ""]
    lines += ["## 二、抽样明细", "",
              "| 章 | 节 | 数值 | 依据对照表指向 | 原始文件 |", "|---|---|---|---|---|"]
    for item in samples:
        lines.append("| {} | {} | {} | {} | {} |".format(
            item["chapter"] or "", item["section"] or "", item["number"] or "",
            item["ref"] or "", item["path"] or ""))
    if not samples:
        lines.append("| — | 无值级样本 | — | — | — |")
    lines += ["", "> 抽样每节取一个值级样本；段级覆盖数见第一节。原始文件路径以本表为准，"
              "正文数字不因本表改动。"]
    return "\n".join(lines)


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="数字溯源抽样表")
    ap.add_argument("--workspace", default=".")
    ap.add_argument("--out", default="")
    ap.add_argument("--sample", type=int, default=20)
    args = ap.parse_args(argv)
    workspace = Path(args.workspace).resolve()
    rows = load_rows(workspace)
    if not rows:
        print("[数字溯源] 找不到 依据对照表.json，跳过")
        return 2
    counts = trace_levels(rows)
    samples = pick_samples(rows, args.sample)
    out = Path(args.out) if args.out else (
        workspace / "成果文件" / "内部件" / "数字溯源表" / "数字溯源抽样.md")
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(render(counts, samples), encoding="utf-8")
    print("[数字溯源] 值级 {}｜标准值 {}｜结构值 {}｜段级 {}｜无源 {}｜抽样 {}".format(
        counts["值级"], counts["标准值"], counts["结构值"], counts["段级"],
        counts["无源"], len(samples)))
    return 0


if __name__ == "__main__":
    sys.exit(main())
