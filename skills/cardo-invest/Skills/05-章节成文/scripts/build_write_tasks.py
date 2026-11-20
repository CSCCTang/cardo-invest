#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""build_write_tasks.py -- 生成「章节写作任务包」（cardo/invest-write-tasks/v1）。

为什么要有这个环节：
  流水线的脚本必须是**确定性、可审计、零 LLM 调用**——同一输入必须产出同一结果，
  否则 07 的回读自检、08 的独立复算与基线比对都失去意义。
  但「把事实与测算结果组织成可交付的论证性正文」需要语言模型：
  纯脚本只能填空，产出的段是字段罗列（结论—证据—判断三段式里的「判断」缺失）。

  因此采用与 CardoStatus/06 步同一套模式（也与「联网检索补料」同范式）：
      脚本生成任务包（输入） → 人/agent 用大模型成文并回填（输出） → 流水线消费回填稿。
  这样既有大模型的成文质量，又不破坏流水线的可审计性与离线可跑性。

与 CardoStatus/06 的差异（本包特有，必须讲清）：
  - Status 的输入是**散材料**，语言能力用在「组织材料」上；
  - Invest 的输入是**引擎已算好的数字**（投资估算/财务测算/方案比选），
    语言能力只用于「组织论证」，**不产生数字**。
  - 因此本任务包把「可引用的事实」与「引擎已定的数字」分两块给模型：
      事实（facts）来自上游现状分析报告/槽位/政策，可引用；
      数字（figures）来自财务测算.json 等，**只能照抄，不许改写、不许推算**。

产物（默认 过程文件/05-章节成文/写作任务/）：
  * 写作任务.md      —— 人可读的提示词与素材清单
  * 写作任务.json    —— 机器可读任务包（供 agent 批量处理）
  * 分章初稿_待回填/  —— 每章一个空 .md 模板，agent 直接写入即可

用法：
    python build_write_tasks.py --project-dir <工作区> [--out <目录>] [--chapters CH-01,CH-02]
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPT_DIR))

import _lib  # noqa: E402

SKILL_DIR = SCRIPT_DIR.parent
RULES_DIR = SKILL_DIR / "规则"
DEC = Path("05_过程") / "投资决策"

# 各章成文要点（人话版；硬判据以 章节契约.yaml 为准）
CHAPTER_BRIEF = {
    "CH-01": "全书电梯陈述：项目是什么 + 投多少 + 财务是否可行 + AHP 评级 + 建议。"
             "1.1 五项必答（名称/地点与范围/性质/周期/主体）+ 投资控制目标；"
             "1.3 六个数字必须与第七、八、九章逐字相同。",
    "CH-02": "论证必要性：政策背景（列已取得的政策文件，不引未取得的）、区域发展背景、"
             "现存主要问题、必要性的宏观/中观/微观三层论证。",
    "CH-03": "归纳上游现状分析报告的结论：空间、建筑、交通、市政、业态五方面，"
             "作为 5.1 范围界定与 5.3 改造内容的共同依据。",
    "CH-04": "市场与定位：区域市场简述、总体功能定位、业态规划方案、运营收入预判。"
             "定位须与 6.x 运营模式自洽。",
    "CH-05": "建设方案：范围与用地、改造内容与工程量（逐项对应）、技术方案、"
             "建设时序安排。工程量必须来自上游现状分析报告，不得推算。",
    "CH-06": "组织架构与运营模式：实施主体与职责分工、长效运营收益模式、"
             "运营期收入与成本构成（数字取自财务测算，不得改写）。",
    "CH-07": "投资估算与资金筹措：总投资构成（工程费/其他费/预备费/建设期利息）、"
             "分年投资使用计划、资金筹措方案。三者与第 8 章同源。",
    "CH-08": "财务评价：评价说明与参数、经营收入、年成本支出、盈利能力、"
             "债务清偿能力、可持续性与敏感性。数字全部取自 财务测算.json。",
    "CH-09": "AHP 量化分析：指标体系、判别矩阵与权重、一致性检验、方案打分与综合评分、"
             "核心结论。AHP 结论与现金流结论不得互相替代。",
    "CH-10": "附件清单：列出支撑材料及其与章节的对应关系，并说明尚未附送的材料。",
}

# 引擎已算定的数字：模型只能照抄，不得改写/推算
FIGURE_SOURCES = [
    ("投资估算.json", "总投资构成、分年投资计划、资金筹措"),
    ("财务测算.json", "收入/成本/税费/现金流/全部指标/敏感性/偿债"),
    ("方案比选.json", "三方案投资与指标、推荐方案"),
]


def has_real_body(text) -> bool:
    """回填稿里有没有**实际正文**（剔除 HTML 注释、标题行与引用行后的实质文字量）。"""
    body = []
    for line in (text or "").splitlines():
        stripped = line.strip()
        if not stripped or stripped.startswith("<!--") or stripped.startswith("#") or stripped.startswith(">"):
            continue
        body.append(stripped)
    return len("".join(body)) > 200


def load_yaml(path):
    """读规则 YAML。解析失败**必须报错**——规则文件坏了却静默返回空字典，
    会让禁用词、来源行格式、章节要求全部退化为空，流水线表面正常、实则失守。"""
    import yaml
    try:
        with open(path, encoding="utf-8-sig") as handle:
            return yaml.safe_load(handle) or {}
    except FileNotFoundError:
        raise SystemExit("[写作任务] 缺规则文件：{}".format(path))
    except Exception as error:
        raise SystemExit("[写作任务] 规则文件解析失败：{}（{}）".format(path, error))


def norm(text):
    return re.sub(r"\s+", "", str(text or ""))


def chapter_nodes(contract, chapter_id):
    """取该章的契约条目（章 + 其下各节）。"""
    out = []
    for chapter in contract.get("chapters") or []:
        if str(chapter.get("id")) != chapter_id:
            continue
        out.append({"kind": "chapter", "id": chapter.get("id"), "number": chapter.get("number"),
                    "title": chapter.get("title"), "goal": chapter.get("goal"),
                    "wordMin": chapter.get("wordMin"),
                    "requiredQuestions": chapter.get("requiredQuestions") or [],
                    "paragraphPlan": chapter.get("paragraphPlan") or [],
                    "requiredTables": chapter.get("requiredTables") or [],
                    "requiredFigures": chapter.get("requiredFigures") or [],
                    "antiPatterns": chapter.get("antiPatterns") or []})
        for section in chapter.get("sections") or []:
            out.append({"kind": "section", "id": section.get("id"), "number": section.get("number"),
                        "title": section.get("title"), "goal": section.get("goal"),
                        "wordMin": section.get("wordMin"),
                        "requiredQuestions": section.get("requiredQuestions") or [],
                        "paragraphPlan": section.get("paragraphPlan") or [],
                        "requiredTables": section.get("requiredTables") or [],
                        "requiredFigures": section.get("requiredFigures") or [],
                        "antiPatterns": section.get("antiPatterns") or []})
        break
    return out


def collect_facts(upstream):
    """上游事实：模型可引用（带证据与来源）。"""
    facts = (upstream or {}).get("facts") or {}
    rows = []
    for key, item in facts.items():
        if not isinstance(item, dict):
            continue
        value = item.get("value")
        if value is None:
            continue
        unit = item.get("unit") or ""
        rows.append({
            "key": key,
            "name": item.get("name") or key,
            "value": value if isinstance(value, (int, float)) else str(value),
            "unit": unit,
            "evidence": str(item.get("evidence") or "")[:300],
            "source": str(item.get("source") or "")[:200],
        })
    return rows


def collect_figures(dec, project):
    """引擎已算定的数字：模型只能照抄。"""
    blocks = []
    estimate = _lib.read_json(dec / "投资估算.json") or {}
    finance = _lib.read_json(dec / "财务测算.json") or {}
    schemes = _lib.read_json(dec / "方案比选.json") or {}

    if estimate:
        indicators = finance.get("indicators") or {}
        period = finance.get("period") or {}
        rows = [
            ("项目总投资", estimate.get("totalInvestment"), "万元"),
            ("工程费用", estimate.get("engineering"), "万元"),
            ("工程建设其他费", estimate.get("otherFeeTotal"), "万元"),
            ("预备费", estimate.get("contingency"), "万元"),
            ("建设期利息", estimate.get("constructionInterest"), "万元"),
            ("资本金", (estimate.get("funding") or {}).get("equity"), "万元"),
            ("专项融资（借款）", (estimate.get("funding") or {}).get("loan"), "万元"),
            ("全部投资FIRR(税后)", indicators.get("firrAll"), "小数"),
            ("全部投资FNPV(税后)", indicators.get("fnpvAll"), "万元"),
            ("全部投资回收期", indicators.get("ptAll"), "年"),
            ("资本金FIRR(税后)", indicators.get("firrEquity"), "小数"),
            ("资本金FNPV(税后)", indicators.get("fnpvEquity"), "万元"),
            ("资本金回收期", indicators.get("ptEquity"), "年"),
            ("投资收益率ROI", indicators.get("roi"), "小数"),
            ("资本金净利润率ROE", indicators.get("roe"), "小数"),
            ("利息备付率ICR", indicators.get("icr"), "倍"),
            ("偿债备付率DSCR", indicators.get("dscr"), "倍"),
            ("财务评价建设期", period.get("constructionYears"), "年"),
            ("财务评价运营期", period.get("operationYears"), "年"),
            ("计算期", period.get("totalYears"), "年"),
            ("折现率", period.get("discountRate"), "小数"),
        ]
        for label, value, unit in rows:
            if value is None:
                continue
            blocks.append({"label": label, "value": value, "unit": unit,
                           "source": "财务测算.json" if label not in (
                               "项目总投资", "工程费用", "工程建设其他费", "预备费", "建设期利息",
                               "资本金", "专项融资（借款）") else "投资估算.json"})
    if schemes:
        winner = ((schemes.get("compareTable") or {}).get("recommendation") or {}).get("winner")
        for row in (schemes.get("compareTable") or {}).get("rows") or []:
            blocks.append({
                "label": "方案 {}（{}）FIRR".format(row.get("id"), row.get("name")),
                "value": row.get("firrPercent"), "unit": "%", "source": "方案比选.json",
            })
        if winner:
            blocks.append({"label": "现金流口径推荐方案", "value": winner, "unit": "", "source": "方案比选.json"})
    return blocks


# 各章可用的上游小节（按「章节契约」的章意匹配上游现状分析报告的小节标题）
# 匹配不上就给出全部上游小节标题清单，让模型自行判断可用哪节。
CHAPTER_UPSTREAM_HINTS = {
    "CH-01": ["项目概况", "项目编制依据", "项目调研范围与期限"],
    "CH-02": ["项目编制依据", "片区基础概况分析", "现状详细分析", "核心问题汇总"],
    "CH-03": ["现状详细分析", "片区基础概况分析"],
    "CH-04": ["片区基础概况分析", "现状详细分析", "核心问题汇总"],
    "CH-05": ["项目概况", "现状详细分析", "核心问题汇总"],
    "CH-06": ["核心问题汇总", "现状综合研判与更新思路"],
    "CH-07": ["项目概况"],
    "CH-08": [],
    "CH-09": [],
    "CH-10": [],
}


def collect_figures_index(contract, chapter_id):
    """本章的图清单（id → 标题/类型/数据来源），供模型写图题占位。"""
    rows = []
    for item in contract.get("figuresIndex") or []:
        if str(item.get("chapter")) != chapter_id:
            continue
        rows.append({"id": item.get("id"), "number": item.get("displayNumber"),
                     "title": item.get("title"), "type": item.get("type"),
                     "data": item.get("data"), "note": item.get("note") or ""})
    return rows


def collect_tables_index(contract, chapter_id):
    """本章的表清单（id → 标题），供模型写表题占位。"""
    rows = []
    for item in contract.get("tablesIndex") or []:
        if str(item.get("chapter")) != chapter_id:
            continue
        rows.append({"id": item.get("id"), "number": item.get("displayNumber"),
                     "title": item.get("title"), "note": item.get("note") or ""})
    return rows


def collect_upstream_text(upstream, chapter_id, max_sections=6, max_paras=8):
    """上游原文段落：按章意挑小节，给模型可直接引用的素材（含小节标题）。"""
    sections = [item for item in (upstream or {}).get("sections") or []
                if isinstance(item, dict)]
    if not sections:
        return []
    hints = CHAPTER_UPSTREAM_HINTS.get(chapter_id) or []
    picked = []
    if hints:
        for hint in hints:
            for item in sections:
                title = str(item.get("title") or "")
                if hint in title and item not in picked:
                    picked.append(item)
                if len(picked) >= max_sections:
                    break
            if len(picked) >= max_sections:
                break
    if not picked:
        picked = sections[:max_sections]
    rows = []
    for item in picked:
        title = str(item.get("title") or "")
        for paragraph in (item.get("paragraphs") or [])[:max_paras]:
            text = str(paragraph).strip()
            if len(text) < 12:
                continue
            rows.append({"section": title, "text": text[:600]})
    return rows


def build_tasks(project_dir: Path, out_dir: Path, chapters=None):
    dec = project_dir / DEC
    contract = load_yaml(RULES_DIR / "章节契约.yaml")
    spec = load_yaml(RULES_DIR / "写作规范.yaml")
    upstream = _lib.read_json(dec / "上游摘要.json") or {}
    payload_in = _lib.read_json(dec / "财务输入.json") or {}

    facts = collect_facts(upstream)
    figures = collect_figures(dec, project_dir)
    slots = (payload_in.get("slots") or {})

    all_chapters = {str(item.get("id")): item for item in (contract.get("chapters") or [])}
    wanted = [c.strip() for c in (chapters or "").split(",") if c.strip()] or list(all_chapters.keys())

    out_dir.mkdir(parents=True, exist_ok=True)
    draft_dir = out_dir / "分章初稿_待回填"
    draft_dir.mkdir(parents=True, exist_ok=True)

    tone = ((spec.get("tone") or {}).get("principle") or [])
    banned = [str(item.get("pattern")) for item in (spec.get("bannedWords") or [])
              if isinstance(item, dict) and item.get("pattern")]
    banned_reasons = [{"pattern": str(item.get("pattern")), "reason": str(item.get("reason") or ""),
                       "replacement": str(item.get("replacement") or "")}
                      for item in (spec.get("bannedWords") or []) if isinstance(item, dict)]
    source_format = ((spec.get("sourceLine") or {}).get("format")
                     or "来源：<文件相对路径>#<字段或小节>")
    paragraph_triplet = (((spec.get("tone") or {}).get("paragraphTriplet") or {}).get("rule")
                         or "结论句 + 证据/测算 + 判断或建议")

    tasks = []
    for chapter_id in wanted:
        chapter = all_chapters.get(chapter_id)
        if chapter is None:
            continue
        nodes = chapter_nodes(contract, chapter_id)
        chapter_facts = facts
        upstream_text = collect_upstream_text(upstream, chapter_id)
        figure_slots = collect_figures_index(contract, chapter_id)
        table_slots = collect_tables_index(contract, chapter_id)
        brief = CHAPTER_BRIEF.get(chapter_id, "")
        requirements = []
        for node in nodes:
            for question in node.get("requiredQuestions") or []:
                requirements.append("[{} {}] {}".format(node.get("number") or node.get("id"),
                                                        node.get("title"), question))
            for plan in node.get("paragraphPlan") or []:
                requirements.append("[{} {}] {}".format(node.get("number") or node.get("id"),
                                                        node.get("title"), plan))
        if chapter.get("requiredTables"):
            requirements.append("本章须含表格：{}".format("、".join(chapter["requiredTables"])))
        if chapter.get("requiredFigures"):
            requirements.append("本章须含图：{}".format("、".join(chapter["requiredFigures"])))
        if chapter.get("wordMin"):
            requirements.append("本章正文不少于 {} 字".format(chapter["wordMin"]))

        task = {
            "chapter": chapter_id,
            "number": chapter.get("number"),
            "title": chapter.get("title"),
            "goal": chapter.get("goal"),
            "brief": brief,
            "requirements": requirements,
            "sections": [{
                "number": node.get("number"), "title": node.get("title"),
                "goal": node.get("goal"), "wordMin": node.get("wordMin"),
                "kind": node.get("kind"),
            } for node in nodes if node.get("kind") == "section"],
            "facts": chapter_facts,
            "upstreamText": upstream_text,
            "figures": figures,
            "figureSlots": figure_slots,
            "figureCaptionLines": ["图 {}  {}".format(item.get("number") or "", item.get("title") or "")
                                   for item in figure_slots],
            "tableSlots": table_slots,
            "slots": slots,
            "bannedPhrases": banned,
            "bannedWordDetail": banned_reasons,
            "numbersRule": (spec.get("numbers") or {}),
            "tone": tone,
            "sourceLineFormat": source_format,
            "paragraphTriplet": paragraph_triplet,
            "antiPatterns": chapter.get("antiPatterns") or [],
            "outputFile": "{}_{}.md".format(chapter_id, chapter.get("title")),
        }
        tasks.append(task)

        template = ["# {} {}".format(chapter.get("number"), chapter.get("title")), "",
                    "<!-- 请由大模型成文后写入本文件；完成后删除本注释。 -->", "",
                    "<!-- 章节：{} -->".format(chapter_id), ""]
        if brief:
            template += ["> 本章任务：{}".format(brief), ""]
        template += ["## 各节内容", "",
                     "<!-- 按 sections 逐节成文；每节写「结论句 → 证据/测算 → 判断或建议」三段；",
                     "     节末写一行来源行，格式：{} -->".format(source_format), ""]
        if figure_slots or table_slots:
            # 图题必须是**独立行**「图 <章>-<序> <标题>」——06 渲染器按这个正则识别并画图；
            # 写成 HTML 注释则渲染器认不出，成品会 0 张图、自检 fail。
            template += ["## 本章图题与表题（图题写成独立行，渲染器据此画图）", "",
                         "<!-- 把下列图题原样写成正文里的独立行（不要加注释符），",
                         "     并放在对应论述段落之后。 -->", ""]
            for item in table_slots:
                template += ["<!-- 表位：{} {} {} -->".format(
                    item["id"], item.get("number") or "", item.get("title") or "")]
            for item in figure_slots:
                template += ["图 {}  {}".format(item.get("number") or "", item.get("title") or "")]
            template += [""]
        for node in nodes:
            if node.get("kind") != "section":
                continue
            template += ["### {} {}".format(node.get("number"), node.get("title")), ""]
            if node.get("goal"):
                template += ["<!-- 本节目标：{} -->".format(node["goal"]), ""]
            if node.get("requiredQuestions"):
                template += ["<!-- 本节必答：", ]
                for question in node["requiredQuestions"]:
                    template += ["     - {} -->".format(question)]
                template += [""]
        # **绝不覆盖已有回填稿**：agent/大模型写好的正文是人工/模型投入的成果，
        # 重新生成任务包时若把模板覆盖回去，等于白干一轮。
        # 只有文件不存在（或还是空模板）时才写新模板。
        draft_path = draft_dir / task["outputFile"]
        if draft_path.is_file():
            existing = draft_path.read_text(encoding="utf-8-sig", errors="ignore")
            if has_real_body(existing):
                continue                      # 已有成文稿：保留，不动
        draft_path.write_text("\n".join(template), encoding="utf-8")

    payload = {
        "schema": "cardo/invest-write-tasks/v1",
        "project": (payload_in.get("project") or {}).get("name", ""),
        "chapters": tasks,
        "_说明": "每章的成文要求见 requirements/brief；可用事实见 facts；"
                 "引擎已算定的数字见 figures（只能照抄，不得改写或推算）；"
                 "成文结果写入 分章初稿_待回填/<章>_<标题>.md 后重跑流水线 05。",
    }
    (out_dir / "写作任务.json").write_text(
        json.dumps(payload, ensure_ascii=False, indent=1), encoding="utf-8")
    _write_md(out_dir / "写作任务.md", payload)
    return payload


def _fmt(value, unit):
    if value is None:
        return ""
    if isinstance(value, float):
        text = "{:g}".format(value)
    else:
        text = str(value)
    return "{}{}".format(text, unit or "")


def _write_md(path, payload):
    lines = ["# 章节写作任务 · {}".format(payload.get("project") or "本项目"), "",
             "> 本文件是**给大模型的输入**。请按每章的「成文要求 + 可用事实 + 引擎数字」生成章节正文，",
             "> 写入 `分章初稿_待回填/<章>_<标题>.md`，然后重跑流水线（05 会自动采用回填稿）。", "",
             "> 硬约束：",
             "> - **不得新增事实、不得改数字、不得用行业经验值补空缺**；缺什么写什么并登记缺项编号；",
             "> - `figures` 是引擎已算定的数字，**只能照抄**，不得改写、不得四舍五入、不得推算；",
             "> - 每节写「结论句 → 证据/测算 → 判断或建议」三段，节末写一行来源行；",
             "> - 禁用词（出现即判不合格）：{}；".format(
                 "、".join(payload["chapters"][0]["bannedPhrases"]) if payload["chapters"] else "无"),
             "> - 来源行格式：{}（须具体到「文件相对路径#字段或小节」，不得只写文件名）。".format(
                 (payload["chapters"][0]["sourceLineFormat"] if payload["chapters"] else "")),
             "> - 每节写「结论句 → 证据/测算 → 判断或建议」三段，缺一即判「只罗列不分析」。",
             "> - 标题层级：章用 `#`，小节用契约给定的标题逐字（不要自造层级、不要加三级标题）。",
             "",
             "> **可直接照抄的样例**（格式照抄，内容换成你的）：",
             ">",
             "> ```markdown",
             "> # 1 项目概述",
             ">",
             "> ## 1.1 项目基本信息",
             ">",
             "> 本项目名称为 XX，改造范围为 XX 片区，项目性质为老旧街区改造。",
             "> 建设周期 3 年，其中建设期 2 年、运营期 1 年，合计计算期 3 年。",
             "> 实施主体为 XX 街道办事处；项目总投资 20000 万元（口径：财务测算.json#indicators.investment）。",
             ">",
             "> 来源：财务测算.json#indicators.investment；上游摘要.md#项目基本情况",
             "> ```",
             ">",
             "> 缺材料的小节照抄这个（**编号格式别写错**，自检按这个格式认）：",
             ">",
             "> ```markdown",
             "> 本节材料缺失，无法说明 XX，缺项编号【缺项-C03：见本节缺项表】。",
             "> ```",
             ">",
             "> 禁用词自查：上面样例里的「项目资料」「大概」「待确认」等词**一律不许出现**",
             "> （自检按 写作规范.yaml 的 bannedWords 逐条判 fail）。",
             "> **必答问题必须逐条回答**：每节的 requiredQuestions 都列在下面「必答要点」，",
             "> 要么在正文里明确回答，要么写「（资料缺失：…，见缺项表）」——漏答判 fail。",
             ""]
    for task in payload["chapters"]:
        lines += ["## {} {}".format(task["chapter"], task["title"]), "",
                  "**本章任务**：{}".format(task["brief"]), ""]
        if task["requirements"]:
            lines += ["**成文要求**：", ""]
            lines += ["- {}".format(item) for item in task["requirements"]]
            lines += [""]
        lines += ["**可用事实（可引用，带证据与来源）**：", ""]
        if task["facts"]:
            for fact in task["facts"]:
                lines.append("- {} {}｜{}｜来源：{}".format(
                    fact["name"], fact["unit"], fact["evidence"][:110], fact["source"]))
        else:
            lines.append("- （无上游事实：涉及处写「（资料缺失：…，见缺项表）」并登记缺项编号）")
        lines += [""]
        if task.get("upstreamText"):
            lines += ["**上游原文段落（可摘引，勿改写数字）**：", ""]
            for row in task["upstreamText"]:
                lines.append("- [{}] {}".format(row["section"], row["text"][:220]))
            lines += [""]
        lines += ["**引擎已算定的数字（只能照抄）**：", ""]
        for figure in task["figures"]:
            lines.append("- {} = {}｜来源：{}".format(
                figure["label"], _fmt(figure["value"], figure["unit"]), figure["source"]))
        lines += [""]
        if task["slots"]:
            lines += ["**用户槽位**：", ""]
            for key, value in sorted(task["slots"].items()):
                if isinstance(value, list):
                    value = "、".join(str(item) for item in value)
                lines.append("- {}：{}".format(key, value))
            lines += [""]
        lines += ["**各节**：", ""]
        for section in task["sections"]:
            lines.append("- {} {}（不少于 {} 字）{}".format(
                section["number"], section["title"], section.get("wordMin") or "-",
                section.get("goal") or ""))
        lines += [""]
    path.write_text("\n".join(lines), encoding="utf-8")


def main(argv=None):
    parser = argparse.ArgumentParser(description="生成章节写作任务包（大模型成文输入）")
    parser.add_argument("--project-dir", required=True)
    parser.add_argument("--out", default="")
    parser.add_argument("--chapters", default="", help="只生成指定章，如 CH-01,CH-02")
    args = parser.parse_args(argv)

    project_dir = Path(args.project_dir).resolve()
    out_dir = Path(args.out) if args.out else project_dir / DEC / "写作任务"
    chapters = [c.strip() for c in args.chapters.split(",") if c.strip()] or None
    payload = build_tasks(project_dir, out_dir, chapters)
    print("[写作任务] {} 章 → {}".format(len(payload["chapters"]), _lib.rel_posix(out_dir, project_dir)))
    print("[写作任务] 提示词：{}".format(out_dir / "写作任务.md"))
    print("[写作任务] 回填：{}".format(out_dir / "分章初稿_待回填"))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
