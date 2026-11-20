#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""check_keyq.py -- 必答要点校验（ACC-KEYQ）。

为什么单独一个脚本：章节契约里每节都写了 `requiredQuestions`
（如「Q1：项目名称、建设地点、四至或改造范围、项目性质分别是什么」），
`acceptanceCatalog` 里 ACC-KEYQ 标了 `severity: fail`，
但**代码里从来没有执行过**——契约引用的 `check_report_v2.py` 这个文件根本不存在。
结果是：必答问题漏答也没人管。

判定口径（照契约 `requiredQuestionsSemantics`）：
  「必答问题必须在小节正文里被明确回答；允许换句式，不允许漏答；答不出写（资料缺失：…）」
故按**领域关键词命中**判（不是逐字比对）：
  · 从问题文本里抽出领域关键词（去掉 Q1:、问号、疑问词、常见虚词）；
  · 该节正文命中其中**足够比例**的关键词即视为已回答；
  · 命中不足但该节写了「（资料缺失：…）」占位 → 判 warn（已按缺项登记，不阻断）；
  · 既没答也没写缺项 → fail。

退出码：0 全部合格；1 有 fail 级漏答。
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import _lib  # noqa: E402

# 抽关键词时要丢掉的虚词与疑问词
STOP = {
    "q1", "q2", "q3", "q4", "q5", "q6", "q7", "q8",
    "什么", "哪些", "哪个", "多少", "如何", "怎样", "是否", "分别", "什么",
    "本章", "本节", "小节", "分别", "分别是什么", "与", "和", "及", "或",
    "为", "是", "在", "的", "了", "吗", "呢", "并", "其中", "是否一致",
    "允许", "写", "见", "如下", "包括", "不含", "有没有",
    # 疑问词剥离后残留的碎片
    "多长", "多大", "几个", "几处", "长", "含不含", "包不包含",
    "长不长", "是否含", "含含", "不含不含",
}


def _keywords(question: str) -> list[str]:
    """从必答问题里抽领域关键词。

    抽得太粗会全盘落空：早先按逗号切，「实施主体是谁，是否拟设项目公司、股权与治理结构如何」
    切出「实施主体是谁」这种**整块**，正文写「实施主体是某街道办」也匹配不上。
    故再细分两步：
      ① 按分隔符切出语义单元；
      ② 每个单元去掉疑问尾巴（「…是什么」「…多少」「…如何」「是谁」等），
         太短的（<2 字）丢弃。
    """
    text = str(question or "")
    text = re.sub(r"^Q\d+[:：]\s*", "", text)
    text = re.sub(r"[（(].*?[）)]", "", text)
    text = text.replace("？", "").replace("?", "").replace("。", "")
    text = re.sub(r"不允许.*$", "", text)
    parts = re.split(r"[，,、；;：:／/（）()\s]+", text)
    out = []
    for part in parts:
        part = part.strip()
        if len(part) < 2:
            continue
        low = part.lower()
        if low in STOP:
            continue
        # 去疑问尾巴：「项目性质是什么」→「项目性质」
        part = re.sub(r"(是什么|分别是多少|多少|如何|怎样|是谁|是否|哪个|哪些|吗|呢)$", "", part)
        # 去疑问词**在中间**的情况：「建设周期多长」→「建设周期」、
        # 「含不含建设期利息」→「含建设期利息」。否则整块当关键词，正文永远匹配不上。
        part = re.sub(r"(多长|多大|多少|如何|怎样|是谁|是否|几个|几处)$", "", part)
        part = re.sub(r"^(是否|含不含|包不包含|包括)", "", part)
        part = part.strip()
        if len(part) < 2 or part.lower() in STOP:
            continue
        if part not in out:
            out.append(part)
    return out


def load_contract(project: Path) -> list[dict]:
    """读章节契约的 chapters → 节点列表（含 number/title/requiredQuestions）。"""
    path = None
    for cand in (project / "05_过程" / "投资决策" / "_rules" / "章节契约.yaml",):
        if cand.is_file():
            path = cand
            break
    if path is None:
        for cand in project.rglob("章节契约.yaml"):
            path = cand
            break
    if path is None:
        return []
    data = _lib.load_yaml(str(path)) or {}
    nodes = []
    for chapter in (data.get("chapters") or []):
        cid = str(chapter.get("id") or chapter.get("number") or "")
        # 章级本身也可能有 requiredQuestions（如第 1 章的「五个核心问题」）
        if chapter.get("requiredQuestions"):
            nodes.append({
                "chapter": cid,
                "number": str(chapter.get("number") or ""),
                "title": str(chapter.get("title") or ""),
                "questions": list(chapter.get("requiredQuestions") or []),
                "scope": "chapter",
            })
        # 子节点用 `sections` 键（不是 nodes）
        for node in (chapter.get("sections") or chapter.get("nodes") or []):
            if not node.get("requiredQuestions"):
                continue
            nodes.append({
                "chapter": cid,
                "number": str(node.get("number") or ""),
                "title": str(node.get("title") or ""),
                "questions": list(node.get("requiredQuestions") or []),
                "scope": "section",
            })
    return nodes


def _section_bodies(text: str) -> dict:
    """按二级标题切出正文块：{标题: 正文}。"""
    lines = text.splitlines()
    bodies, cur_title, cur = {}, None, []
    for ln in lines:
        stripped = ln.strip()
        if stripped.startswith("## "):
            if cur_title is not None:
                bodies[cur_title] = "\n".join(cur)
            cur_title = stripped[3:].strip()
            cur = []
        elif cur_title is not None:
            cur.append(ln)
    if cur_title is not None:
        bodies[cur_title] = "\n".join(cur)
    return bodies


def _match_section(node: dict, bodies: dict) -> tuple[str, str]:
    """按节点号或标题找到对应小节正文。"""
    number = node.get("number") or ""
    title = node.get("title") or ""
    for heading, body in bodies.items():
        if number and heading.startswith(number):
            return heading, body
    for heading, body in bodies.items():
        if title and title in heading:
            return heading, body
    return "", ""


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="必答要点校验（ACC-KEYQ）")
    ap.add_argument("--project-dir", required=True)
    ap.add_argument("--chapters", default="", help="章稿目录（默认 05_过程/投资决策/分章）")
    ap.add_argument("--out", default="", help="检查报告 JSON 输出路径")
    ap.add_argument("--min-hit-ratio", type=float, default=0.34,
                    help="每道必答问题判定为「已回答」所需的关键词命中比例")
    args = ap.parse_args(argv)

    project = Path(args.project_dir)
    ch_dir = Path(args.chapters) if args.chapters else project / "05_过程" / "投资决策" / "分章"
    nodes = load_contract(project)
    if not nodes:
        print("[必答要点] 未找到章节契约的 requiredQuestions，跳过")
        return 0
    files = sorted(ch_dir.glob("*.md")) if ch_dir.is_dir() else []
    if not files:
        print("[必答要点] {} 里没有章稿，跳过".format(ch_dir))
        return 0

    problems, checked = [], 0
    for f in files:
        text = f.read_text(encoding="utf-8")
        bodies = _section_bodies(text)
        for node in nodes:
            heading, body = _match_section(node, bodies)
            if not heading:
                continue      # 该节还没写（别的问题会报）
            checked += 1
            has_gap = "（资料缺失" in body
            for qi, question in enumerate(node["questions"], start=1):
                kws = _keywords(question)
                if not kws:
                    continue
                hit = [k for k in kws if k in body]
                # 关键词是「领域词」而非原句：正文写「实施主体是某街道办」，
                # 关键词「实施主体」能命中；再放宽一档——关键词去掉尾字后
                # 在正文里出现（应对「项目性质」vs「项目性质为改造」这类写法）。
                if not hit:
                    for k in kws:
                        stem = k[:-1] if len(k) > 3 else k
                        if stem and stem in body:
                            hit.append(k)
                ratio = len(hit) / len(kws)
                if ratio >= args.min_hit_ratio:
                    continue
                # 章级必答问题（如「本章五个核心问题是否齐全」）是**元问题**，
                # 不是让正文复述的题干，用关键词命中判会永远落空 → 降为提示。
                # 小节级才是真必答（漏答要 fail）。
                if node.get("scope") == "chapter":
                    severity = "warn"
                elif has_gap:
                    severity = "warn"      # 已按缺项登记占位
                else:
                    severity = "fail"
                problems.append({
                    "file": f.name, "node": node["number"], "title": node["title"],
                    "question": question,
                    "hitRatio": round(ratio, 3),
                    "missed": [k for k in kws if k not in body][:6],
                    "severity": severity,
                })

    fails = [p for p in problems if p["severity"] == "fail"]
    warns = [p for p in problems if p["severity"] == "warn"]
    payload = {"schema": "cardo/invest-keyq-check/v1",
               "checkedFiles": len(files), "checkedSections": checked,
               "minHitRatio": args.min_hit_ratio,
               "ok": not fails,
               "fails": fails, "warns": warns}
    if args.out:
        Path(args.out).parent.mkdir(parents=True, exist_ok=True)
        Path(args.out).write_text(json.dumps(payload, ensure_ascii=False, indent=2),
                                  encoding="utf-8")

    print("[必答要点] {} 章 / {} 小节：漏答 {} 处（fail {}、warn {}）".format(
        len(files), checked, len(problems), len(fails), len(warns)))
    for p in fails[:10]:
        print("   FAIL {} {} 「{}」 命中 {}{}".format(
            p["file"], p["node"], p["question"][:44], p["hitRatio"],
            "　未命中：{}".format("、".join(p["missed"])) if p["missed"] else ""))
    for p in warns[:5]:
        print("   WARN {} {} 「{}」 已按缺项登记占位，仅提示".format(
            p["file"], p["node"], p["question"][:40]))
    if fails:
        print("   漏答的必答问题要么在正文里答，要么写「（资料缺失：…，见缺项表）」。")
    return 1 if fails else 0


if __name__ == "__main__":
    sys.exit(main())
