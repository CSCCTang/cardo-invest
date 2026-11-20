"""00 · Cardo投资决策总控（唯一入口）。

按 01→08 顺序调用编号技能，每步跑完过闸门，不通过即停并说明原因。
所有中间产物落「过程文件/<编号-名称>/」，成果只落「成果文件/」（Office）。

用法（工作区根执行）：
  python Skills/00-Cardo投资决策总控/scripts/pipeline.py status --workspace .
  python Skills/00-Cardo投资决策总控/scripts/pipeline.py run    --workspace .
  python Skills/00-Cardo投资决策总控/scripts/pipeline.py run    --workspace . --from 03
  python Skills/00-Cardo投资决策总控/scripts/pipeline.py run    --workspace . --scenario gaps
  python Skills/00-Cardo投资决策总控/scripts/pipeline.py scaffold --target <新目录> --project-name <名称>

退出码：0 全过；1 某步失败（脚本报错或闸门未过）；2 参数/工作区/前置产物问题。
"""
from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path

SCHEMA_LOG = "cardo-invest/pipeline-log/v1"
# 技能链根：优先环境变量（工作区可在技能包之外），缺省为本包 Skills/
SKILLS = Path(os.environ.get("CARDO_SKILLS_ROOT")
              or Path(__file__).resolve().parents[2]).resolve()
STEP_NAMES = {
    "01": "上游取料",
    "02": "匡算",
    "03": "财务评价",
    "04": "三方案比选",
    "05": "章节成文",
    "06": "渲染docx",
    "07": "回读自检",
    "08": "验证与边界",
}
STEP_DIRS = {key: "{}-{}".format(key, name) for key, name in STEP_NAMES.items()}
# 01b 是 01 的前置子步骤（自动从可研/估算材料提取并补全 财务参数.xlsx），
# 不作为独立流水线步骤，只提供脚本路径解析。
STEP_DIRS["01b"] = "01b-可研财务提取"
DEFAULT_PATHS = {
    "inputRoot": "输入文件",
    "upstreamJson": "过程文件/00-总控/上游机读件",
    "financialParams": "输入文件/财务参数.xlsx",
    "template": "../提资数据/模板/建议书模板.docx",
    "stagingProject": "过程文件/00-总控/引擎项目",
}
ARTIFACTS = {
    "01": ["财务输入.json", "取料清单.json", "缺项表.md"],
    "02": ["投资估算.json"],
    "03": ["财务测算.json", "计算轨迹.json"],
    "04": ["方案比选.json", "方案对比表.md"],
    "05": ["分章/CH-01.md", "分章/来源映射.json", "缺项表.md"],
    "06": ["自检-渲染.json"],
    "07": ["自检.json"],
    "08": ["独立验证结果.json", "Office导出清单.json", "产物指纹.json"],
}


# ---------------------------------------------------------------- 基础

def read_json(path, default=None):
    try:
        return json.loads(Path(path).read_text(encoding="utf-8-sig"))
    except (OSError, json.JSONDecodeError):
        return default


def write_json(path, payload):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2, sort_keys=True), encoding="utf-8")


def rel(path, base):
    try:
        return Path(path).resolve().relative_to(Path(base).resolve()).as_posix()
    except ValueError:
        return str(path)


def now():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def workspace_config(workspace):
    config = read_json(workspace / "工作区参数.json") or {}
    paths = dict(DEFAULT_PATHS)
    paths.update(config.get("paths") or {})
    return config, paths


def step_dir(workspace, step):
    return workspace / "过程文件" / STEP_DIRS[step]


def copy_artifact(source, target):
    """先删旧再复制：闸门不静默，旧产物不残留。"""
    source, target = Path(source), Path(target)
    if not source.is_file():
        return False
    target.parent.mkdir(parents=True, exist_ok=True)
    if target.exists():
        target.unlink()
    shutil.copy2(source, target)
    return True


def remove_tree(path):
    if Path(path).exists():
        shutil.rmtree(path)


# ---------------------------------------------------------------- 执行与闸门

class StepFailure(Exception):
    def __init__(self, message, detail=None):
        super().__init__(message)
        self.detail = detail or {}


class Gate(StepFailure):
    """内容闸门不通过（禁用词 / 跨章数字 / 必答要点）。

    与「脚本跑挂」区分开：闸门是成品不合格，退出码同为 1，但提示语指向具体规范。
    这里必须真定义——早先只写 `raise Gate(...)` 而未定义类，闸门一触发就变成
    NameError，把「成品有禁用词」报成「程序内部错误」。
    """


def run_cmd(workspace, argv, label):
    started = time.time()
    env = dict(os.environ)
    env["PYTHONIOENCODING"] = "utf-8"
    completed = subprocess.run(
        [sys.executable] + argv,
        cwd=str(workspace),
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        env=env,
    )
    record = {
        "label": label,
        "argv": [Path(item).name if item.endswith(".py") else item for item in argv],
        "exit": completed.returncode,
        "stdout": (completed.stdout or "").strip().splitlines()[-12:],
        "stderr": (completed.stderr or "").strip().splitlines()[-12:],
        "seconds": round(time.time() - started, 2),
    }
    return record


def fail_on_exit(record, allowed=(0,)):
    if record["exit"] not in allowed:
        raise StepFailure("{} 退出码 {}（允许 {}）".format(record["label"], record["exit"], list(allowed)),
                          {"stdout": record["stdout"], "stderr": record["stderr"]})


def script(step, name):
    return str(SKILLS / STEP_DIRS[step] / "scripts" / name)


# 「缺数就问用户」：把缺口翻译成可直接照填的表格（表名/单元格/填什么/示例），
# 用户填完重跑同一条命令即可续跑，不需要重新准备投递区。
FILL_HINTS = {
    "project.direction": ("项目", "B3", "三选一：老旧街区 / 老旧小区 / 老旧厂区", "老旧街区"),
    "project.constructionYears": ("项目", "B5", "整数，建设期年数", "2"),
    "project.operationYears": ("项目", "B6", "整数，运营期年数", "20"),
    "project.discountRate": ("项目", "B7", "小数；3.5% 写 0.035", "0.035"),
    "estimate.contingency.rate": ("预备费", "B2", "小数；可由「预备费÷（工程费+其他费）」反算", "0.05"),
}


def fill_request_path(workspace):
    return workspace / "过程文件" / "00-总控" / "待填-财务参数.md"


def finance_source(manifest):
    for item in (manifest or {}).get("placed") or []:
        if str(item.get("target") or "").endswith("财务参数.json"):
            return item.get("source")
    return "输入文件/财务参数.xlsx"


def write_fill_request(workspace, payload, manifest):
    """把缺失的财务字段写成待填清单；没有缺失则返回 None。"""
    project = payload.get("project") or {}
    estimate = payload.get("estimate") or {}

    def present(value):
        return value is not None and str(value).strip() != ""

    missing = []
    if not present(project.get("direction")):
        missing.append("project.direction")
    if not present(project.get("constructionYears")):
        missing.append("project.constructionYears")
    if not present(project.get("operationYears")):
        missing.append("project.operationYears")
    if not present(project.get("discountRate")):
        missing.append("project.discountRate")
    if not present((estimate.get("contingency") or {}).get("rate")):
        missing.append("estimate.contingency.rate")
    if not missing:
        return None

    source = finance_source(manifest)
    lines = [
        "# 待填财务参数（缺数问用户）",
        "",
        "- 项目：{}".format(project.get("name") or "（未识别）"),
        "- 填写文件：`{}`".format(source),
        "- 填完重跑：`python 开工.py` 或 `python Skills/00-Cardo投资决策总控/scripts/pipeline.py run --workspace . --from 01`",
        "- 说明：只补下表空格，不要改已填单元格；填完 01 会重新取件并续跑。",
        "",
        "| 字段 | sheet | 单元格 | 要填什么 | 填写示例 |",
        "|---|---|---|---|---|",
    ]
    for key in missing:
        sheet, cell, need, example = FILL_HINTS[key]
        lines.append("| `{}` | {} | {} | {} | `{}` |".format(key, sheet, cell, need, example))
    lines += [
        "",
        "> 这些字段为空时，流水线不会静默补数；缺项会进入 `缺项表.md`，正文对应位置标「待定」。",
    ]
    path = fill_request_path(workspace)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return path



# ---------------------------------------------------------------- 各步

def step_01(workspace, paths, scenario, log):
    staging = workspace / paths["stagingProject"]
    out_dir = step_dir(workspace, "01")
    remove_tree(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    # 01b 先跑：从输入材料（可研 PDF / 投资估算 xlsx）自动提取并补全
    #      输入文件/财务参数.xlsx。用户没提供时由它生成，有则原地补全、
    #      不覆盖用户已填的格子。失败不阻断（01 会照常报「缺财务参数」）。
    record = run_cmd(workspace, [script("01b", "build_params.py"), "--workspace", ".", "--no-timestamp"],
                     "01b 可研财务提取（自动补全财务参数）")
    log.append(record)
    if record["exit"] not in (0, 2):
        raise StepFailure("01b 财务参数自动提取失败：{}".format(
            record.get("stderr") or record.get("stdout")))

    record = run_cmd(workspace, [script("01", "prepare_inputs.py"), "--workspace", ".", "--scenario", scenario,
                                 "--no-timestamp"], "01 装配引擎项目")
    log.append(record)
    fail_on_exit(record, allowed=(0, 3))
    manifest = read_json(out_dir / "取料清单.json") or {}
    if record["exit"] == 3 or manifest.get("status") == "need-input":
        raise StepFailure("01 缺必需财务参数（待补，不补数）", {"manifest": manifest})

    record = run_cmd(workspace, [script("01", "collect_inputs.py"),
                                 "--project-dir", rel(staging, workspace), "--no-timestamp"], "01 取料与缺项")
    log.append(record)
    fail_on_exit(record)

    # 上游正文结构化：现状分析报告按 H1/H2 分节 + 抽表 → 上游摘要.json（05 的事实底稿）
    # 报告缺失不阻断（05 出缺口版），但必须留痕，不允许静默跳过
    record = run_cmd(workspace, [script("01", "extract_upstream.py"),
                                 "--project-dir", rel(staging, workspace)], "01 上游正文结构化")
    log.append(record)
    if record["exit"] not in (0, 3):
        raise StepFailure("01 上游摘要抽取失败：{}".format(record.get("stderr") or record.get("stdout")))
    payload = read_json(staging / "05_过程/投资决策/财务输入.json", {})
    if not payload:
        raise StepFailure("01 未产出财务输入.json")

    gaps = payload.get("gaps") or []
    for item in gaps:
        if not all(item.get(key) for key in ("code", "what", "need", "owner")):
            raise StepFailure("01 gaps 四要素不全：{}".format(item))
    fill_request = write_fill_request(workspace, payload, manifest)
    ahp = payload.get("ahpWeights") or {}
    temporary = bool(ahp.get("temporary")) or ahp.get("kind") == "default"
    if not (staging / "05_过程/AHP/权重表.json").is_file() and not temporary:
        raise StepFailure("01 AHP 缺失但未标临时权重")

    copied = []
    for pair in (("财务输入.json", "财务输入.json"), ("缺项表.md", "缺项表.md"),
                 ("上游摘要.json", "上游摘要.json")):
        if copy_artifact(staging / "05_过程/投资决策" / pair[0], out_dir / pair[1]):
            copied.append(pair[1])
    if (out_dir / "取料清单.json").is_file():
        copied.append("取料清单.json")
    gate = {
        "gaps": [item["code"] for item in gaps],
        "gapsOwner": sorted({item["owner"] for item in gaps}),
        "ahpTemporary": temporary,
        "sources": len(payload.get("sources") or []),
        "placeholders": len(payload.get("placeholders") or []),
        "fillRequest": rel(fill_request, workspace) if fill_request else None,
    }
    return gate, copied, payload


def step_02(workspace, paths, log):
    staging = workspace / paths["stagingProject"]
    dec = staging / "05_过程/投资决策"
    out_dir = step_dir(workspace, "02")
    remove_tree(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    if not (dec / "财务输入.json").is_file():
        raise StepFailure("02 前置缺失：01 未产出财务输入.json")

    record = run_cmd(workspace, [script("02", "calc_invest.py"),
                                 "--input", rel(dec / "财务输入.json", workspace),
                                 "--out", rel(dec, workspace),
                                 "--format", "json", "--no-timestamp"], "02 匡算")
    log.append(record)
    # 允许 exit 3：缺项已登记（gaps）但分项能算多少算多少，属「带缺项继续」，
    # 不能把「有缺项」当脚本失败。真正的硬停交给下游：只有结构性算不出的数
    # （如总投资为空导致现金流无法建立）才在 03 停。
    fail_on_exit(record, allowed=(0, 3))
    estimate = read_json(dec / "投资估算.json", {})
    if not estimate:
        raise StepFailure("02 未产出投资估算.json")

    total = estimate.get("totalInvestment")
    trace = estimate.get("trace") or {}
    steps = {item.get("id"): item.get("value") for item in trace.get("steps") or []}
    parts = [estimate.get("engineering"), estimate.get("otherFeeTotal"), estimate.get("contingency"),
             estimate.get("constructionInterest")]
    if any(value is not None for value in parts) and not steps:
        raise StepFailure("02 缺计算轨迹（分项不可复核）")
    funding = estimate.get("funding") or {}
    funding_total = estimate.get("fundingTotal")
    # 资金筹措不平由 02 自己登记为 FUNDING-MISMATCH 缺项（calc_invest 退出码 3），
    # 这里只兜底「引擎没登记却对不上」的情况，不再把已登记的不平当硬失败。
    gaps = estimate.get("gaps") or []
    gap_codes = {item.get("code") for item in gaps}
    if total is not None and funding_total is not None \
            and abs(float(funding_total) - float(total)) > 1 \
            and "FUNDING-MISMATCH" not in gap_codes:
        raise StepFailure("02 资金筹措与总投资不平且未登记缺项", {"fundingTotal": funding_total, "total": total})

    # 缺项只登记、不补数：总投资为空时仍保留分项估算与缺口清单，下游正文标「待定」。
    if total is None and not gaps:
        raise StepFailure("02 总投资为空但未登记缺项（缺项必须可追溯）", {"note": estimate.get("note")})

    copied = ["投资估算.json"] if copy_artifact(dec / "投资估算.json", out_dir / "投资估算.json") else []
    gate = {
        "totalInvestment": total,
        "incomplete": total is None,
        "gaps": [item.get("code") for item in gaps],
        "funding": funding,
        "engineVersion": trace.get("engineVersion"),
        "steps": len(trace.get("steps") or []),
    }
    if total is None:
        print("[总控] 02 带缺项继续：总投资待定（{}），分项估算与缺口清单已落盘".format(
            "、".join(item.get("code", "") for item in gaps)))
    elif gaps:
        print("[总控] 02 带缺项继续：{}（总投资 {} 万元）".format(
            "、".join(item.get("code", "") for item in gaps), total))
    return gate, copied, estimate


def step_03(workspace, paths, log):
    staging = workspace / paths["stagingProject"]
    dec = staging / "05_过程/投资决策"
    out_dir = step_dir(workspace, "03")
    remove_tree(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    for name in ("财务输入.json", "投资估算.json"):
        if not (dec / name).is_file():
            raise StepFailure("03 前置缺失：{}".format(name))

    # 结构性硬停：总投资为空 → 折旧/摊销/现金流/指标全都没有基数，
    # 再往下算只能编数。这不是「缺项可带」，而是整章算不出来，必须停下来补料。
    estimate = read_json(dec / "投资估算.json", {}) or {}
    if estimate.get("totalInvestment") is None:
        gaps = [item.get("code") for item in estimate.get("gaps") or []]
        raise StepFailure(
            "03 结构性缺数：总投资为空，财务评价无法计算（缺项：{}）".format(
                "、".join(gaps) if gaps else "未登记"),
            {"estimate": {"engineering": estimate.get("engineering"),
                          "otherFeeTotal": estimate.get("otherFeeTotal"),
                          "contingency": estimate.get("contingency"),
                          "constructionInterest": estimate.get("constructionInterest"),
                          "totalInvestment": None},
             "gaps": estimate.get("gaps")})

    record = run_cmd(workspace, [script("03", "calc_finance.py"), "--project-dir", rel(staging, workspace),
                                 "--no-timestamp"], "03 收入成本税费与现金流")
    log.append(record)
    # 允许 exit 3（= 算得出但有缺项待确认，如无经营收入的政府投资项目）。
    # 口径：缺数**只报缺项不补数**，缺项由后面的 gaps 校验把关
    #（缺指标必须登记 gaps，否则本步失败）——不能因为「有缺项」就阻断整条链。
    # 实测：某企业饮品可研是纯财政投资项目（无借款、无经营收入），
    # 原先 exit 3 被当成失败，链路卡在 03。
    fail_on_exit(record, allowed=(0, 3))
    record = run_cmd(workspace, [script("03", "calc_metrics.py"), "--project-dir", rel(staging, workspace),
                                 "--no-timestamp"], "03 指标与轨迹")
    log.append(record)
    fail_on_exit(record, allowed=(0, 3))     # 3 = 指标缺项（如资本金现金流不能回收），产物仍落盘

    finance = read_json(dec / "财务测算.json", {})
    if not finance:
        raise StepFailure("03 未产出财务测算.json")
    indicators = finance.get("indicators") or {}
    required = ("firrAll", "fnpvAll", "ptAll", "firrEquity", "fnpvEquity", "ptEquity", "roe", "roi", "icr", "dscr")
    # 回收期为 null 是结论（计算期内累计净现金流未转正），不是缺料：
    # calc_metrics 已按此口径显式约定「不算 gap、不补数」，这里不能用 gap 去逼一个数字出来。
    missing = [key for key in required if indicators.get(key) is None and key not in ("ptAll", "ptEquity")]
    declared = {item.get("what", "") for item in finance.get("gaps") or []}
    if missing:
        # 允许“算不出”的指标，但必须在 gaps 里显式登记原因，不允许静默留空
        if not finance.get("gaps"):
            raise StepFailure("03 指标缺项但未登记 gaps：{}".format(missing), {"gaps": finance.get("gaps")})
        if indicators.get("firrAll") is None or indicators.get("fnpvAll") is None:
            raise StepFailure("03 核心指标（全部投资 FIRR/FNPV）算不出", {"gaps": finance.get("gaps")})
    tables = finance.get("tables") or {}
    cashflow_rows = ((tables.get("cashflows") or {}).get("allInvestment") or {}).get("rows") or []
    if not cashflow_rows:
        raise StepFailure("03 缺现金流表（表位 1/2 无数据源）")

    copied = []
    for name in ("财务测算.json", "计算轨迹.json"):
        if copy_artifact(dec / name, out_dir / name):
            copied.append(name)
    gate = {
        "indicators": {key: indicators[key] for key in required},
        "indicatorsGap": missing,
        "gaps": [item.get("code") for item in finance.get("gaps") or []],
        "cashflowRows": len(cashflow_rows),
        "period": finance.get("period"),
    }
    return gate, copied, finance


def step_04(workspace, paths, log):
    staging = workspace / paths["stagingProject"]
    dec = staging / "05_过程/投资决策"
    out_dir = step_dir(workspace, "04")
    remove_tree(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    for name in ("财务输入.json", "投资估算.json", "财务测算.json"):
        if not (dec / name).is_file():
            raise StepFailure("04 前置缺失：{}".format(name))

    record = run_cmd(workspace, [script("04", "gen_schemes.py"), "--project-dir", rel(staging, workspace),
                                 "--no-timestamp"], "04 三方案比选")
    log.append(record)
    # 允许 exit 3（= 方案关键参数缺口待确认，产物仍落盘；见 04 SKILL 退出码定义）。
    # 缺口由本步后面的方案数/id/推荐方案校验把关，不能因「有缺口」就阻断整条链。
    fail_on_exit(record, allowed=(0, 3))
    schemes = read_json(dec / "方案比选.json", {})
    if not schemes:
        raise StepFailure("04 未产出方案比选.json")
    rows = (schemes.get("compareTable") or {}).get("rows") or []
    if len(rows) != 3:
        raise StepFailure("04 方案数不是 3：{}".format(len(rows)))
    if len({row.get("id") for row in rows}) != 3:
        raise StepFailure("04 方案 id 重复")
    winner = ((schemes.get("compareTable") or {}).get("recommendation") or {}).get("winner")
    if not winner:
        raise StepFailure("04 无推荐结论（第九章无法回写）")

    copied = []
    for name in ("方案比选.json", "方案对比表.md"):
        if copy_artifact(dec / name, out_dir / name):
            copied.append(name)
    gate = {
        "schemes": [{"id": row.get("id"), "name": row.get("name"), "score": row.get("score")} for row in rows],
        "winner": winner,
        "weightsSource": (schemes.get("compareTable") or {}).get("recommendation", {}).get("why", {}).get("weightsSource"),
    }
    return gate, copied, schemes


CHAPTER_SKILLS = [
    ("CH-01", "05-01-项目概述"), ("CH-02", "05-02-项目背景与建设必要性"),
    ("CH-03", "05-03-现状调研分析"), ("CH-04", "05-04-市场分析与功能定位"),
    ("CH-05", "05-05-项目建设范围、改造内容与建设方案"), ("CH-06", "05-06-组织架构与运营模式"),
    ("CH-07", "05-07-投资估算与资金筹措方案"), ("CH-08", "05-08-财务评价"),
    ("CH-09", "05-09-基于AHP层次分析法的投资决策指标体系量化分析"), ("CH-10", "05-10-附件清单"),
]


FILL_MARK = "<!-- 请由大模型成文后写入本文件"


def _has_real_body(text) -> bool:
    """回填稿里有没有**实际正文**。

    判据不能只看「有没有 H1」「字数够不够」「提示注释删没删」——
    大模型正常成文时也常把模板的「## 各节内容」「本节目标：」等占位注释留在文件里。
    可靠做法：**剔除全部 HTML 注释与空行后**，看还剩不剩实质文字段落。

    另有一条硬判据：模板里的 FILL_MARK 还在，就一定是没回填的空白稿。
    只按字数判会把「模板提示语本身够长」的空白稿当成已回填（实测：10 章全降级
    却被报成「已回填 8/10」），后面的说明就全错了。
    """
    if FILL_MARK in (text or ""):
        return False
    body = []
    for line in (text or "").splitlines():
        stripped = line.strip()
        if not stripped:
            continue
        if stripped.startswith("<!--"):
            continue
        if stripped.startswith("#"):
            continue
        if stripped.startswith(">"):
            continue
        body.append(stripped)
    return len("".join(body)) > 200

def list_filled_drafts(filled_dir) -> list:
    """已回填的章节号列表。

    判定条件：存在 CH-*.md，且已写入实际正文（模板提示注释已清除、含 H1）。
    支持**部分章节回填**：回填了几章就用几章，其余章走确定性成文降级，
    不因为只填了一章就卡住整条流水线。
    """
    directory = Path(filled_dir)
    if not directory.is_dir():
        return []
    filled = []
    for path in sorted(directory.glob("CH-*.md")):
        try:
            text = path.read_text(encoding="utf-8-sig")
        except OSError:
            continue
        if not text.strip():
            continue
        # 判定「已回填」只看**有没有实际正文**：有 H1 且有实质内容即算。
        # 不要求大模型删掉模板里的提示注释——那是它的疏忽，不该让它把整条流水线卡住；
        # 提示注释残留由 07 的 DOC-RESIDUE 检查去抓（成品里出现提示行即 fail）。
        # 判「有没有实际正文」：剔除 HTML 注释与标题行后，剩下的是否还有实质文字。
        # 不能只看「提示注释删没删」——大模型正常成文时也常把模板注释留在文件里。
        if "# " in text and _has_real_body(text):
            filled.append(path.stem.split("_")[0].strip())
    return filled


def step_05(workspace, paths, log):
    staging = workspace / paths["stagingProject"]
    dec = staging / "05_过程/投资决策"
    out_dir = step_dir(workspace, "05")
    remove_tree(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    for name in ("财务输入.json", "投资估算.json", "财务测算.json"):
        if not (dec / name).is_file():
            raise StepFailure("05 前置缺失：{}".format(name))

    # ── 两段式：先出「章节写作任务包」，等 agent/大模型成文回填，再消费回填稿 ──
    # 与 CardoStatus/06 步同一套模式（见 --research 的范式）：
    #   脚本保持零 LLM 调用、离线可跑；成文质量由大模型按 章节契约.yaml 保证。
    # 未回填时走确定性成文降级（不阻断）；已回填则优先采用回填稿。
    task_dir = dec / "写作任务"
    filled_dir = task_dir / "分章初稿_待回填"
    filled = list_filled_drafts(filled_dir)
    all_ids = ["CH-{:02d}".format(index) for index in range(1, 11)]
    degraded = [cid for cid in all_ids if cid not in filled]
    if not filled:
        # 一条命令跑到底：任务包照出，但不阻断——真实用户只会敲一次 开工.py，
        # 停在这里要求他再敲 --from 05，会被当成脚本坏了。
        # 这里按确定性降级成文跑完，并**明说 0 章回填**（不许把降级稿说成「已回填」），
        # 需要大模型成文只作为提示打印。
        if not (task_dir / "写作任务.md").is_file():
            record = run_cmd(workspace, [script("05", "build_write_tasks.py"),
                                         "--project-dir", rel(staging, workspace),
                                         "--out", rel(task_dir, workspace)],
                             "05 章节写作任务包（可选：大模型成文）")
            log.append(record)
            fail_on_exit(record)
        print("[总控] 05 本次走确定性成文（本包不启用大模型回填，见 SKILL.md「章节成文」）")
        print("[总控] 如需人工改写：按 {} 写入 {} 后重跑即可（不重跑也已出完整成果）".format(
            rel(task_dir / "写作任务.md", workspace), rel(filled_dir, workspace)))
    else:
        print("[总控] 05 已回填 {}/10 章（{}）；降级成文 {} 章（{}）".format(
            len(filled), "、".join(filled), len(degraded),
            "、".join(degraded) if degraded else "无"))

    record = run_cmd(workspace, [script("05", "assemble_chapters.py"), "--project-dir", rel(staging, workspace),
                                 "--no-timestamp"], "05-00 分章引擎（10 章一次成稿）")
    log.append(record)
    fail_on_exit(record)

    chapters = sorted((dec / "分章").glob("CH-*.md"))
    expected = ["CH-{:02d}.md".format(index) for index in range(1, 11)]
    names = [chapter.name for chapter in chapters]
    if names != expected:
        raise StepFailure("05 章数/章序不符：{}".format(names))
    empty = [chapter.name for chapter in chapters if not chapter.read_text(encoding="utf-8").strip()]
    if empty:
        raise StepFailure("05 分章稿为空：{}".format(empty))

    # 按章节 skill 顺序逐章复核；任一章不通过，停在本章
    check_dir = out_dir / "分章自检"
    check_dir.mkdir(parents=True, exist_ok=True)
    chapter_gates = []
    for cid, skill_dir in CHAPTER_SKILLS:
        out_json = check_dir / "{}.json".format(cid)
        record = run_cmd(workspace, [script("05", "check_chapter.py"),
                                     "--project-dir", rel(staging, workspace),
                                     "--chapter", cid, "--out", rel(out_json, workspace),
                                     "--no-timestamp"], "05 {} {}".format(cid, skill_dir))
        log.append(record)
        fail_on_exit(record)
        payload = read_json(out_json, {}) or {}
        chapter_gates.append({"chapter": cid, "skill": skill_dir, "ok": payload.get("ok") is True})

    placeholders = sum(chapter.read_text(encoding="utf-8").count("（资料缺失") for chapter in chapters)

    # ACC-NOBAN 禁用词 + ACC-KEYQ 必答要点：
    # 两条都在规则里标了 severity: fail，但**代码里从来没执行过**
    # （契约引用的 check_report_v2.py 根本不存在）。这里补上实际校验。
    # 禁用词是 fail 级；必答要点里「已写缺项占位」的按 warn 处理。
    banned_out = check_dir / "禁用词检查.json"
    record = run_cmd(workspace, [script("05", "check_banned.py"),
                                 "--project-dir", rel(staging, workspace),
                                 "--chapters", rel(out_dir / "分章", workspace),
                                 "--out", rel(banned_out, workspace)],
                      "05 禁用词检查")
    log.append(record)
    if record.get("exit") not in (0, None):
        print("[05] 禁用词检查未通过：见 {}".format(rel(banned_out, workspace)))
        raise Gate("成品命中禁用词（ACC-NOBAN，severity=fail）："
                   "按 写作规范.yaml 的 bannedWords 逐条改，"
                   "或确属例外则说明理由。详见 {}".format(rel(banned_out, workspace)))

    # A-07 跨章数字一致性：1.3 的核心数字必须与 7/8/9 章**逐字相同**。
    # 契约里标了 fail 级，但代码里从来没有实现过——1.3 与财务章对不上时
    # 读者会看到全书两套数字，流水线却全绿。
    cross_out = check_dir / "跨章数字检查.json"
    record = run_cmd(workspace, [script("05", "check_crossnum.py"),
                                 "--project-dir", rel(staging, workspace),
                                 "--chapters", rel(out_dir / "分章", workspace),
                                 "--out", rel(cross_out, workspace)],
                      "05 跨章数字一致性检查")
    log.append(record)
    cross = read_json(cross_out, {}) or {}
    if cross.get("conflicts"):
        print("[05] 跨章数字矛盾 {} 处：".format(len(cross["conflicts"])))
        for item in cross["conflicts"][:6]:
            print("   1.3 的「{}」= {}，但第 {} 章写了 {}".format(
                item.get("indicator"), item.get("expected"), item.get("chapter"),
                "、".join("{}={}".format(lbl, val)
                          for lbl, val in (item.get("conflictingWith") or []))))
        raise Gate("跨章数字矛盾（A-07，severity=fail）："
                   "同一指标在全书必须是同一个值——口径以 财务测算.json 为准，"
                   "改错的那一章。详见 {}".format(rel(cross_out, workspace)))

    keyq_out = check_dir / "必答要点检查.json"
    record = run_cmd(workspace, [script("05", "check_keyq.py"),
                                 "--project-dir", rel(staging, workspace),
                                 "--chapters", rel(out_dir / "分章", workspace),
                                 "--out", rel(keyq_out, workspace)],
                      "05 必答要点检查")
    log.append(record)
    keyq = read_json(keyq_out, {}) or {}
    if keyq.get("fails"):
        print("[05] 必答要点漏答 {} 处：".format(len(keyq["fails"])))
        for item in keyq["fails"][:6]:
            print("   {} {} 「{}」".format(item.get("file"), item.get("node"),
                                          str(item.get("question"))[:50]))
        raise Gate("必答问题漏答（ACC-KEYQ，severity=fail）："
                   "要么在小节正文里明确回答，要么写「（资料缺失：…，见缺项表）」。"
                   "详见 {}".format(rel(keyq_out, workspace)))

    remove_tree(out_dir / "分章")
    (out_dir / "分章").mkdir(parents=True, exist_ok=True)
    copied = []
    for chapter in chapters:
        if copy_artifact(chapter, out_dir / "分章" / chapter.name):
            copied.append("分章/" + chapter.name)
    if copy_artifact(dec / "分章/来源映射.json", out_dir / "分章/来源映射.json"):
        copied.append("分章/来源映射.json")
    if copy_artifact(dec / "缺项表.md", out_dir / "缺项表.md"):
        copied.append("缺项表.md")

    # 依据对照表：把分章稿的「来源：」行逐块登记，并核每条引用查不查得到
    record = run_cmd(workspace, [script("05", "build_evidence_register.py"),
                                 "--project-dir", rel(workspace, workspace),
                                 "--chapters", rel(out_dir / "分章", workspace),
                                 "--out", rel(out_dir / "分章/依据对照表.json", workspace)],
                      "05 依据对照表")
    log.append(record)
    if record.get("exit") not in (0,):
        raise StepFailure("05 依据对照表生成失败", {"exit": record.get("exit")})
    copied.append("分章/依据对照表.json")
    gate = {"chapters": len(chapters), "placeholders": placeholders,
            "chapterChecks": len([item for item in chapter_gates if item["ok"]]),
            "sourceMap": len((read_json(dec / "分章/来源映射.json", {}) or {}).get("chapters") or [])}
    return gate, copied, None


def step_06(workspace, paths, log):
    staging = workspace / paths["stagingProject"]
    dec = staging / "05_过程/投资决策"
    template = workspace / paths["template"]
    out_dir = step_dir(workspace, "06")
    remove_tree(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    if not template.is_file():
        raise StepFailure("06 缺模板：{}".format(paths["template"]))
    if not (dec / "分章").is_dir():
        raise StepFailure("06 前置缺失：05 未产出分章稿")

    template_arg = os.path.relpath(str(template), str(workspace)).replace("\\", "/")
    record = run_cmd(workspace, [script("06", "build_invest_docx.py"),
                                 "--project-dir", rel(staging, workspace),
                                 "--template", template_arg, "--no-timestamp"], "06 渲染建议书 docx")
    log.append(record)
    fail_on_exit(record)

    produced = sorted((staging / "04_成果报告").glob("*建议书*.docx"))
    if not produced:
        raise StepFailure("06 未产出建议书 docx")
    source = produced[-1]
    check_path = source.with_name(source.stem + ".自检.json")
    render_check = read_json(check_path, {}) or {}
    if render_check.get("result") not in (True, "pass", "通过"):
        raise StepFailure("06 渲染自检未通过", {"自检": render_check.get("result")})

    # 一成果一子目录：成果文件/成品版/<成果名>/<成果名>.docx|.pdf
    deliverable_dir = workspace / "成果文件" / "成品版" / "项目投资决策建议书"
    deliverable_dir.mkdir(parents=True, exist_ok=True)
    deliverable = deliverable_dir / "项目投资决策建议书.docx"
    if deliverable.exists():
        deliverable.unlink()
    shutil.copy2(source, deliverable)
    copied = ["成品版/项目投资决策建议书/项目投资决策建议书.docx"]
    if copy_artifact(check_path, out_dir / "自检-渲染.json"):
        copied.append("自检-渲染.json")

    # 成品 PDF：与 docx 同目录。排版不得靠 Office，故复用 Status 的
    # docx → A4 HTML → Edge 无头打印 → 目录页码回填 → 页脚链路。
    pdf_path = deliverable_dir / "项目投资决策建议书.pdf"
    record = run_cmd(workspace, [script("08", "render_pdf.py"),
                                 "--docx", rel(deliverable, workspace),
                                 "--out", rel(pdf_path, workspace)], "06 成品 PDF")
    log.append(record)
    fail_on_exit(record)
    if not pdf_path.is_file():
        raise StepFailure("06 未产出成品 PDF")
    copied.append("成品版/项目投资决策建议书/项目投资决策建议书.pdf")

    # 依据对照表 xlsx（与人读的对照表；数据来自 05 的登记结果，不另起一套）
    evidence_xlsx = workspace / "成果文件" / "内部件" / "依据对照表" / "依据对照表_投资建议书.xlsx"
    record = run_cmd(workspace, [script("06", "build_evidence_xlsx.py"),
                                 "--project-dir", ".",
                                 "--out", rel(evidence_xlsx, workspace)],
                     "06 依据对照表 xlsx")
    log.append(record)
    fail_on_exit(record)
    if copy_artifact(evidence_xlsx, out_dir / "依据对照表_投资建议书.xlsx"):
        copied.append("依据对照表_投资建议书.xlsx")

    # 成品版两张表：匡算表、财务评价表独立成件（用户意见第 5 条）。
    # 与正文 7/8 章同源，只读 02/03 过程 JSON，不做二次计算。
    record = run_cmd(workspace, [script("08", "export_office.py"), "--workspace", ".", "--deliverables"],
                     "06 成品表（匡算表 / 财务评价表）")
    log.append(record)
    fail_on_exit(record)
    table_manifest = read_json(out_dir.parent / "06-渲染docx" / "成品表导出清单.json", {}) or {}
    table_exports = table_manifest.get("exports") or []
    if len(table_exports) != 2:
        raise StepFailure("06 成品表未产出两张", {"exports": [item.get("file") for item in table_exports]})
    for item in table_exports:
        copied.append(item.get("path"))

    counts = render_check
    gate = {
        "docx": rel(deliverable, workspace),
        "pdf": rel(pdf_path, workspace),
        "chapters": (counts.get("chapters") or {}).get("count"),
        "tables": (counts.get("tables") or {}).get("count"),
        "placeholders": render_check.get("placeholders"),
        "explicitGaps": render_check.get("explicitGaps"),
        "result": render_check.get("result"),
    }
    return gate, copied, None


def step_07(workspace, paths, log):
    staging = workspace / paths["stagingProject"]
    out_dir = step_dir(workspace, "07")
    remove_tree(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    if not any((staging / "04_成果报告").glob("*建议书*.docx")):
        raise StepFailure("07 前置缺失：06 未产出成品 docx")

    out_json = out_dir / "自检.json"
    record = run_cmd(workspace, [script("07", "check_invest.py"),
                                 "--project-dir", rel(staging, workspace),
                                 "--format", "json", "--out", rel(out_json, workspace),
                                 "--no-timestamp"], "07 回读自检")
    log.append(record)
    fail_on_exit(record)
    payload = read_json(out_json, {}) or {}
    summary = payload.get("summary") or {}
    if summary.get("failed"):
        failed = [item for item in payload.get("checks") or [] if not item.get("ok")]
        raise StepFailure("07 自检失败 {} 项".format(summary.get("failed")), {"failed": failed})
    if not payload.get("ok"):
        raise StepFailure("07 自检未通过")

    # 依据核对：分章稿每段出处是否真查得到（不通过即停，不允许"写得出处但查不到"）
    ev_json = out_dir / "依据核对.json"
    record = run_cmd(workspace, [script("07", "check_evidence.py"),
                                 "--project-dir", ".",
                                 "--format", "json", "--out", rel(ev_json, workspace)],
                     "07 依据核对")
    log.append(record)
    evidence = read_json(ev_json, {}) or {}
    if not evidence.get("ok"):
        failed_ev = [item for item in evidence.get("checks") or [] if not item.get("ok")]
        raise StepFailure("07 依据核对未通过：段段有源 / 出处可查 / 数值不裸奔",
                          {"failed": failed_ev,
                           "problems": (evidence.get("problems") or [])[:10]})

    # 自检 JSON 落到 07 后，把成品 docx 的指纹登记进基线候选
    gate = {"passed": summary.get("passed"), "total": summary.get("total"),
            "checks": sorted({item.get("id") for item in payload.get("checks") or []}),
            "evidence": [item.get("id") for item in evidence.get("checks") or []]}
    return gate, ["自检.json", "依据核对.json"], payload


def step_08(workspace, paths, log, baseline_mode):
    out_dir = step_dir(workspace, "08")
    remove_tree(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    record = run_cmd(workspace, [script("08", "export_office.py"), "--workspace", "."], "08 Office 导出")
    log.append(record)
    fail_on_exit(record)

    record = run_cmd(workspace, [script("08", "independent_verify.py"), "--workspace", ".",
                                   "--baseline", baseline_mode], "08 独立复算与边界")
    log.append(record)
    fail_on_exit(record)

    report = read_json(out_dir / "独立验证结果.json", {}) or {}
    summary = report.get("summary") or {}
    if summary.get("failed"):
        failed = [item for item in report.get("checks") or [] if not item.get("ok")]
        raise StepFailure("08 验证失败 {} 项".format(summary.get("failed")), {"failed": failed})
    violations = (report.get("boundary") or {}).get("violations") or []
    if violations:
        raise StepFailure("08 边界违规 {} 条".format(len(violations)), {"violations": violations})

    # 成品闸门：对着 成果文件/成品版 验，不看过程稿
    record = run_cmd(workspace, [script("08", "final_gates.py"), "--workspace", "."], "08 成品闸门")
    log.append(record)
    fail_on_exit(record)

    gate = {"passed": summary.get("passed"), "total": summary.get("total"),
            "warned": summary.get("warned"),
            "boundary": (report.get("boundary") or {}).get("checked"),
            "baseline": (report.get("baseline") or {}).get("mode"),
            "finalGates": (read_json(out_dir / "成品闸门报告.json", {}) or {}).get("ok")}
    return gate, ["独立验证结果.json", "Office导出清单.json", "产物指纹.json", "成品闸门报告.json"], report


# ---------------------------------------------------------------- 编排

def run_steps(workspace, from_step, to_step, scenario, baseline_mode):
    config, paths = workspace_config(workspace)
    log, records = [], []
    started = now()
    steps = [key for key in sorted(STEP_NAMES) if from_step <= key <= to_step]
    for key in steps:
        print("\n=== {} {} ===".format(STEP_DIRS[key], STEP_NAMES[key]))
        try:
            if key == "01":
                gate, copied, _ = step_01(workspace, paths, scenario, log)
            elif key == "02":
                gate, copied, _ = step_02(workspace, paths, log)
            elif key == "03":
                gate, copied, _ = step_03(workspace, paths, log)
            elif key == "04":
                gate, copied, _ = step_04(workspace, paths, log)
            elif key == "05":
                gate, copied, _ = step_05(workspace, paths, log)
            elif key == "06":
                gate, copied, _ = step_06(workspace, paths, log)
            elif key == "07":
                gate, copied, _ = step_07(workspace, paths, log)
            else:
                gate, copied, _ = step_08(workspace, paths, log, baseline_mode)
        except StepFailure as error:
            # 人工闸门（detail.exit=3）：不是失败，是「等用户/agent 介入」，
            # 对外返回 3，与 开工.py 的缺料/待确认语义一致。
            gate_exit = int((error.detail or {}).get("exit") or 1)
            records.append({"step": key, "name": STEP_NAMES[key],
                            "status": "gate" if gate_exit == 3 else "fail",
                            "error": str(error), "detail": error.detail,
                            "commands": log[-3:] if log else []})
            write_logs(workspace, config, scenario, started, records, log, failed=key)
            if gate_exit == 3:
                print("[总控] 第 {} 步待人工介入：{}".format(key, error))
                print("  办完后重跑：python Skills/00-Cardo投资决策总控/scripts/pipeline.py "
                      "run --workspace . --from {}".format(key))
                return 3
            print("[总控] 第 {} 步失败：{}".format(key, error))
            for line in (error.detail or {}).get("stderr") or []:
                print("  " + line)
            return 1
        records.append({"step": key, "name": STEP_NAMES[key], "status": "pass", "gate": gate, "outputs": copied})
        print("[总控] {} 通过：{}".format(key, compact(gate)))
    write_logs(workspace, config, scenario, started, records, log, failed=None)
    print("\n[总控] 01→{} 全部通过（scenario={}）".format(to_step, scenario))
    return 0


def compact(gate):
    if isinstance(gate, dict):
        keys = list(gate)[:6]
        return ", ".join("{}={}".format(key, gate[key] if not isinstance(gate[key], (list, dict))
                                        else (len(gate[key]) if isinstance(gate[key], list) else "obj"))
                         for key in keys)
    return str(gate)


def write_logs(workspace, config, scenario, started, records, commands, failed):
    out_dir = workspace / "过程文件/00-总控"
    out_dir.mkdir(parents=True, exist_ok=True)
    payload = {
        "schema": SCHEMA_LOG,
        "projectName": config.get("projectName"),
        "scenario": scenario,
        "startedAt": started,
        "finishedAt": now(),
        "failedStep": failed,
        "steps": records,
        "commands": commands,
    }
    write_json(out_dir / "运行日志.json", payload)

    lines = ["# Cardo投资决策 运行报告", "",
             "- 项目：{}".format(config.get("projectName") or "—"),
             "- 场景：{}｜开始 {}｜结束 {}".format(scenario, started, payload["finishedAt"]),
             "- 结果：{}".format("失败于第 {} 步".format(failed) if failed else "全部通过"), "",
             "| 步 | 技能 | 结果 | 闸门 / 原因 |", "|---|---|---|---|"]
    for record in records:
        if record["status"] == "pass":
            detail = compact(record.get("gate"))
        else:
            detail = record.get("error") or ""
        lines.append("| {} | {} | {} | {} |".format(record["step"], record["name"],
                                                     "通过" if record["status"] == "pass" else "失败", detail))
    lines += ["", "> 详细命令、退出码与输出见 `运行日志.json`。"]
    (out_dir / "运行报告.md").write_text("\n".join(lines), encoding="utf-8")


def command_status(workspace):
    config, paths = workspace_config(workspace)
    print("工作区：{}".format(workspace))
    print("项目：{}".format(config.get("projectName") or "—"))
    staging = workspace / paths["stagingProject"]
    for key in sorted(STEP_NAMES):
        directory = step_dir(workspace, key)
        present = [name for name in ARTIFACTS.get(key, []) if (directory / name).exists()]
        mark = "✓" if present else "—"
        print("  {} {} {:<12} {}".format(key, STEP_NAMES[key], mark,
                                        "  ".join(present) if present else "未产出"))
    deliverables = sorted((workspace / "成果文件").glob("*")) if (workspace / "成果文件").is_dir() else []
    print("成果：{}".format("、".join(item.name for item in deliverables) if deliverables else "（空）"))
    print("引擎项目：{}".format("已装配" if staging.is_dir() else "未装配"))
    log = read_json(workspace / "过程文件/00-总控/运行日志.json", {}) or {}
    if log.get("finishedAt"):
        print("最近一次：{}，{}".format(log["finishedAt"], "失败于 {} 步".format(log["failedStep"])
                                        if log.get("failedStep") else "全部通过"))
    return 0


SCAFFOLD_DIRS = [
    "输入文件",
    "过程文件/00-总控", "过程文件/01-上游取料", "过程文件/02-匡算", "过程文件/03-财务评价",
    "过程文件/04-三方案比选", "过程文件/05-章节成文", "过程文件/06-渲染docx", "过程文件/07-回读自检",
    "过程文件/08-验证与边界", "成果文件", "记忆与规则",
]


def command_scaffold(workspace, target, project_name, baseline):
    source_skills = workspace / "Skills"
    if not source_skills.is_dir():
        print("[scaffold] 当前工作区没有 Skills/，无法复制技能链")
        return 2
    target = Path(target).resolve()
    if target.exists() and any(target.iterdir()):
        print("[scaffold] 目标目录非空：{}".format(target))
        return 2
    target.mkdir(parents=True, exist_ok=True)
    shutil.copytree(source_skills, target / "Skills",
                    ignore=shutil.ignore_patterns("__pycache__", "*.pyc"))
    for name in SCAFFOLD_DIRS:
        (target / name).mkdir(parents=True, exist_ok=True)
    for name in ("README.md",):
        if (workspace / name).is_file():
            shutil.copy2(workspace / name, target / name)
    config, paths = workspace_config(workspace)
    new_config = dict(config)
    new_config["projectName"] = project_name
    write_json(target / "工作区参数.json", new_config)
    # 基线按包命名空间隔离：只继承本包写的基线，别包的（或无 package 的旧文件）一律不复制
    base_src = workspace / "记忆与规则/验证基线-CardoInvest.json"
    if baseline == "inherit" and base_src.is_file():
        try:
            saved = read_json(base_src) or {}
        except Exception:  # noqa: BLE001
            saved = {}
        if saved.get("package") == "CardoInvest":
            dest = target / "记忆与规则/验证基线-CardoInvest.json"
            dest.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(base_src, dest)
        else:
            print("  [跳过] 旧基线不是本包的（package={}），不继承".format(saved.get("package")))
    print("[scaffold] 已生成独立工作区：{}".format(target))
    print("  1) 往 输入文件/ 放 Office 输入（现状报告/AHP/低空/槽位/财务参数/模板）")
    print("  2) python Skills/00-Cardo投资决策总控/scripts/pipeline.py run --workspace .")
    return 0


def main(argv=None):
    parser = argparse.ArgumentParser(description="Cardo投资决策总控（01→08 编排 + 闸门）")
    sub = parser.add_subparsers(dest="command")
    for name in ("run", "status"):
        item = sub.add_parser(name)
        item.add_argument("--workspace", default=".")
        if name == "run":
            item.add_argument("--from", dest="from_step", default="01", choices=sorted(STEP_NAMES))
            item.add_argument("--to", dest="to_step", default="08", choices=sorted(STEP_NAMES))
            item.add_argument("--scenario", choices=("full", "gaps"), default=None)
            item.add_argument("--baseline", choices=("auto", "write", "check", "none"), default="auto")
    scaffold = sub.add_parser("scaffold")
    scaffold.add_argument("--target", required=True)
    scaffold.add_argument("--project-name", default="某某片区改造项目")
    scaffold.add_argument("--baseline", choices=("none", "inherit"), default="none")
    args = parser.parse_args(argv)

    if not args.command:
        parser.print_help()
        return 2
    workspace = Path(args.workspace).resolve() if hasattr(args, "workspace") else Path(".").resolve()
    if args.command == "status":
        return command_status(workspace)
    if args.command == "scaffold":
        return command_scaffold(workspace, args.target, args.project_name, args.baseline)

    config, _ = workspace_config(workspace)
    scenario = args.scenario or config.get("scenarioDefault") or "full"
    if args.from_step != "01" and not (workspace / "过程文件/00-总控/引擎项目").is_dir():
        print("[总控] --from {} 需要先跑 01（引擎项目未装配）".format(args.from_step))
        return 2
    # 开工第一步：技能包完整性自检。
    # 技能包最常见的故障不是逻辑错，而是**文件没拷全/依赖没装/资产缺失**——
    # 这些在第 6、7 步才暴露，返工成本极高。故开工就查，缺什么说哪个、从哪补。
    doctor = Path(__file__).resolve().parent / "doctor.py"
    if doctor.is_file():
        record = run_cmd(workspace, [str(doctor)], "00 技能包完整性自检")
        if record.get("exit") not in (0, None):
            print("[总控] 技能包完整性自检未通过 —— 请按上面的「怎么补」逐条处理后重跑")
            return 3
    return run_steps(workspace, args.from_step, args.to_step, scenario, args.baseline)


if __name__ == "__main__":
    sys.exit(main())
