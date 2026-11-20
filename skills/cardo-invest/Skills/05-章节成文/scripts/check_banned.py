#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""check_banned.py -- 禁用词校验（ACC-NOBAN）。

为什么单独一个脚本：写作规范里定义了 14 条禁用词（bannedWords），
且每条都标了 `severity: fail`，但**代码里从来没有执行过**——
规则只被写进任务包当提示词，模型写完没人查。
实测结果：禁用词可以自由出现在成品里，流水线照样全绿。

本脚本从 `写作规范.yaml` 读禁用词清单（不硬编码），逐条扫章稿：
  · 一般规则：整篇命中即 fail；
  · BW-05（无|暂无）：**只在来源行判 fail**（规范里的豁免规则：
    「无风险」「无遮挡」等正文用法不误杀，只在来源行才算空话）；
  · BW-02（待补|待补充|待定|待确认）：同段出现「旧口径」或「模板」时豁免。

退出码：0 全部合格；1 有命中（逐条报出位置）。
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import _lib  # noqa: E402


def load_rules(project: Path) -> tuple[list[dict], dict]:
    """从写作规范.yaml 读禁用词清单与豁免规则。

    查找顺序：项目内 _rules 副本 → 工作区内任意副本 → **技能包自带规则**。
    最后一级必须有：早先只找前两级，工作区里没有规则副本时整条闸门静默跳过
    （exit=0），实测放过了 3 处禁用词。规则读不到就是闸门无效，不能当通过。
    """
    path = project / "05_过程" / "投资决策" / "_rules" / "写作规范.yaml"
    if not path.is_file():
        # 回退：工作区里的规则副本
        for cand in project.rglob("写作规范.yaml"):
            path = cand
            break
    if not path.is_file():
        # 回退：技能包自带规则（Skills/05-章节成文/规则/写作规范.yaml）
        path = Path(__file__).resolve().parent.parent / "规则" / "写作规范.yaml"
    if not path.is_file():
        return [], {}
    data = _lib.load_yaml(str(path)) or {}
    words = data.get("bannedWords") or []
    if isinstance(words, dict):
        words = words.get("items") or []
    overrides = data.get("bannedWordScopeOverride") or {}
    return list(words), overrides


def _exempt(text: str, rule_id: str, overrides: dict) -> bool:
    """按规范里的豁免规则判断是否放行。"""
    scope = overrides.get(rule_id) if isinstance(overrides, dict) else None
    if rule_id == "BW-02":
        # 「待确认」仅在标注模板旧口径时豁免
        return ("旧口径" in text or "模板" in text)
    if rule_id == "BW-05":
        # 只在来源行判 fail → 正文里的「无」不误杀
        return False
    return False


def check_file(path: Path, words: list[dict], overrides: dict) -> list[dict]:
    hits = []
    try:
        text = path.read_text(encoding="utf-8")
    except OSError as exc:
        return [{"rule": "READ", "pattern": "", "line": 0, "snippet": str(exc)}]
    lines = text.splitlines()
    for rule in words:
        rid = str(rule.get("id") or "")
        pattern = str(rule.get("pattern") or "")
        if not pattern:
            continue
        try:
            rx = re.compile(pattern)
        except re.error:
            continue
        for no, ln in enumerate(lines, start=1):
            stripped = ln.strip()
            if not stripped:
                continue
            if rid == "BW-05" and not stripped.startswith("来源："):
                continue          # 只在来源行判
            m = rx.search(stripped)
            if not m:
                continue
            if _exempt(stripped, rid, overrides):
                continue
            hits.append({"rule": rid, "pattern": pattern, "line": no,
                         "snippet": stripped[:100],
                         "reason": str(rule.get("reason") or ""),
                         "replacement": str(rule.get("replacement") or "")})
            break        # 每条规则报一次即可
    return hits


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="禁用词校验（ACC-NOBAN）")
    ap.add_argument("--project-dir", required=True)
    ap.add_argument("--chapters", default="", help="章稿目录（默认 05_过程/投资决策/分章）")
    ap.add_argument("--out", default="", help="检查报告 JSON 输出路径")
    args = ap.parse_args(argv)

    project = Path(args.project_dir)
    ch_dir = Path(args.chapters) if args.chapters else project / "05_过程" / "投资决策" / "分章"
    words, overrides = load_rules(project)
    if not words:
        print("[禁用词] 读不到 写作规范.yaml 的 bannedWords，闸门无法执行——按不通过处理")
        print("   期望位置：{}".format(
            Path(__file__).resolve().parent.parent / "规则" / "写作规范.yaml"))
        return 1

    files = sorted(ch_dir.glob("*.md")) if ch_dir.is_dir() else []
    if not files:
        print("[禁用词] {} 里没有章稿，跳过".format(ch_dir))
        return 0

    all_hits = []
    for f in files:
        for h in check_file(f, words, overrides):
            all_hits.append({"file": f.name, **h})

    payload = {"schema": "cardo/invest-banned-check/v1",
               "checkedFiles": len(files),
               "ruleCount": len(words),
               "ok": not all_hits,
               "hits": all_hits}
    if args.out:
        Path(args.out).parent.mkdir(parents=True, exist_ok=True)
        Path(args.out).write_text(json.dumps(payload, ensure_ascii=False, indent=2),
                                  encoding="utf-8")

    if not all_hits:
        print("[禁用词] {} 章、{} 条规则：全部合格".format(len(files), len(words)))
        return 0
    print("[禁用词] 发现 {} 处命中（规则 {} 条）：".format(len(all_hits), len(words)))
    for h in all_hits[:12]:
        print("   {}:{}  {} 「{}」 —— {}".format(
            h["file"], h["line"], h["rule"], h["snippet"][:50], h["reason"]))
    if len(all_hits) > 12:
        print("   …另有 {} 处".format(len(all_hits) - 12))
    print("   按规范改成可核的表述（replacement 字段给了替换建议），或确属例外则说明理由。")
    return 1


if __name__ == "__main__":
    sys.exit(main())
