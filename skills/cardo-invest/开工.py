#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Cardo 投资决策 · 一条命令开工。

依次做四件事：
  依赖自检（缺就在包根建 .venv，按 阿里云 → 清华 → USTC → 官方源 回退安装）→ 建目录与释放模板
  → 校验《财务参数.xlsx》与 工作区参数.json → 调 01→08 流水线。

退出码：0 全过；3 缺料/参数不全；1 流水线硬失败；2 环境或结构问题。
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import traceback

# 运行期不写 __pycache__：Skills/ 是只读交付物，写进 pyc 会污染交付库
# （实测交付库里混进过 19 个 pyc）。这里在导入任何本地模块之前就关掉。
sys.dont_write_bytecode = True

from datetime import datetime
from pathlib import Path

SKILL = Path(__file__).resolve().parent      # 技能包：只读，运行期不写入
RUNTIME_PY = SKILL / "runtime" / "python" / ("python.exe" if os.name == "nt" else "python")
ROOT = SKILL.parent
CACHE_BASE = Path(os.environ.get("CARDO_CACHE_HOME") or os.environ.get("LOCALAPPDATA") or str(Path.home()))
def venv_dir() -> Path:
    """派生仓库：虚拟环境放用户缓存，不写安装目录。"""
    return CACHE_BASE / "cardo" / "venv-invest"
                          # 工作区：技能包所在的上一级目录
# 工作区是否已成功解析（resolve_workspace 会抛 SystemExit；未解析时不能拿默认 ROOT 去写看板状态）
WORKSPACE_RESOLVED = False

# 看板状态注入标记（开工.py 每次运行刷新这一段；看板读不到工作区状态）
STATUS_BEGIN = "<!-- CARDO-STATUS:BEGIN -->"
STATUS_END = "<!-- CARDO-STATUS:END -->"

TITLE = "Cardo 投资决策"
PY_MIN = (3, 9)

TEMPLATE_DST = "提资数据/模板/建议书模板.docx"
DIRS = ["输入文件", "过程文件", "成果文件", "记忆与规则", "提资数据/模板", "输入文件-模板"]
ASSETS = [
    {
        "src": "assets/templates/2 城市更新老旧街区改造项目投资决策建议书（模板）-补充数据来源（包含三类项目）.docx",
        "dst": TEMPLATE_DST,
        "sha256": "F8F97D8BCD317007D2D76E1871F319797FE7B8E297EE7A48365BB5DFAFCE7B10",
        "label": "建议书模板",
    },
]
INPUT_TEMPLATES = [
    {"src": "assets/inputs/财务参数-空白.xlsx", "dst": "输入文件-模板/财务参数-空白.xlsx"},
    {"src": "assets/inputs/槽位-空白.xlsx", "dst": "输入文件-模板/槽位-空白.xlsx"},
]
# 运行期依赖。缺一个就有一类材料/能力读不进来：
#   docx/openpyxl —— Word/Excel 材料与成果表
#   yaml          —— 规则与契约文件（缺它 05 步读不到章节契约）
#   fitz          —— PDF 材料解析
#   numpy/cv2/PIL/rapidocr_onnxruntime —— 低空影像与图片 OCR
IMPORTS = ["docx", "openpyxl", "yaml", "fitz", "PIL"]
RUNTIME_MODULES = ["docx", "openpyxl", "yaml", "fitz",
                   "numpy", "cv2", "PIL", "rapidocr_onnxruntime"]

# 技能包完整性自检：每步的**必需脚本**与**必需规则文件**。
# 为什么要自检：技能包被复制不完整时（实测只拷了 开工.py + assets + runtime，
# 或 runtime 只拷了一部分——site-packages 从 49 个包变成 12 个），
# 流水线会一路跑到很后面才失败，报错还不指向根因。
# 故在开工第一步就查：缺哪个目录/脚本/规则，从哪补。
SKILL_STEPS = [
    ("00-Cardo投资决策总控", ["pipeline.py"], []),
    ("01-上游取料", ["collect_inputs.py", "extract_upstream.py", "prepare_inputs.py"],
     ["财务口径.yaml"]),
    ("02-匡算", ["calc_invest.py"], ["财务口径.yaml"]),
    ("03-财务评价", ["calc_finance.py", "calc_metrics.py"], ["财务口径.yaml"]),
    ("04-三方案比选", ["gen_schemes.py"], ["方案参数表.yaml", "财务口径.yaml"]),
    ("05-章节成文", ["build_write_tasks.py", "assemble_chapters.py", "check_chapter.py",
                     "check_banned.py", "check_keyq.py", "check_crossnum.py",
                     "build_evidence_register.py"],
     ["写作规范.yaml", "章节契约.yaml", "财务口径.yaml"]),
    ("06-渲染docx", ["build_invest_docx.py", "build_evidence_xlsx.py"], ["表格契约.yaml"]),
    ("07-回读自检", ["check_invest.py", "check_evidence.py"],
     ["章节契约.yaml", "表格契约.yaml", "财务口径.yaml"]),
    ("08-验证与边界", ["independent_verify.py", "export_office.py"], ["验证口径.yaml"]),
]
# 05 的十个章节副本（每章一份规则，节标题与必答要点从中读取）
CHAPTER_SKILLS = [
    "05-01-项目概述", "05-02-项目背景与建设必要性", "05-03-现状调研分析",
    "05-04-市场分析与功能定位", "05-05-项目建设范围、改造内容与建设方案",
    "05-06-组织架构与运营模式", "05-07-投资估算与资金筹措方案",
    "05-08-财务评价", "05-09-基于AHP层次分析法的投资决策指标体系量化分析",
    "05-10-附件清单",
]

PARAMS = "工作区参数.json"
SCHEMA = "cardo-invest/workspace/v1"
DEFAULTS = {
    "caliberVersion": 1,
    "scenarioDefault": "full",
    "paths": {
        "inputRoot": "输入文件",
        "upstreamJson": "过程文件/00-总控/上游机读件",
        "financialParams": "输入文件/财务参数.xlsx",
        "template": TEMPLATE_DST,
        "stagingProject": "过程文件/00-总控/引擎项目",
    },
}
REQUIRED = ["projectName"]

SHEETS = ["项目", "工程费", "其他费", "预备费", "建设期利息", "收入", "成本", "资金结构", "可研财务指标"]

PIPELINE = "Skills/00-Cardo投资决策总控/scripts/pipeline.py"
DELIVERABLES = ["成果文件/项目投资决策建议书.docx"]
GUIDES = {
    "01": "财务参数表结构或数值不可用：读 过程文件/00-总控/运行报告.md，指出缺哪张表/哪一列，让用户补 输入文件/财务参数.xlsx，再 --from 01。",
    "03": "核心指标（全部投资 FIRR/FNPV）算不出：通常是建设期利息或资金结构缺项；按报告补参数后 --from 03。",
    "06": "模板缺失或被改：删掉 提资数据/模板/建议书模板.docx 后重跑本命令（自动恢复标准模板），再 --from 06。",
}

# ---------------------------------------------------------------- 收料引导
# 每项：key / 名称 / 是否必需 / 落位（相对工作区）/ 识别模式 / 怎么准备
INTAKE_ITEMS = [
    {"key": "params", "name": "财务参数.xlsx（9 张表）", "required": True,
     "dst": "输入文件/财务参数.xlsx", "patterns": ["*财务参数*.xlsx"],
     "how": "复制 输入文件-模板/财务参数-空白.xlsx 按表填写；工程费至少 1 行（名称 + 金额万元）",
     "why": "唯一硬门槛；没有就不开工"},
    {"key": "slots", "name": "槽位.xlsx（功能定位/业态/改造思路/运营模式/组织架构/时序）", "required": False,
     "dst": "输入文件/槽位.xlsx", "patterns": ["*槽位*.xlsx"],
     "how": "复制 输入文件-模板/槽位-空白.xlsx 填写",
     "why": "第 4/5/6 章的定位、思路、组织与运营正文来源；缺则写「资料缺失」"},
    {"key": "status", "name": "现状分析报告.docx（CardoStatus 成果）", "required": False,
     "dst": "输入文件/（任意层级）", "patterns": ["*现状分析*.docx", "*现状分析*.md"],
     "how": "如果你手上已有 CardoStatus 跑出来的《现状分析报告》，放进 输入文件/ 即可；没有就跳过",
     "why": "第 1—6 章的事实来源（范围/面积/工程量/问题/人口）"},
    {"key": "ahp", "name": "AHP 权重表 / 方案评分结论（CardoAHP 成果）", "required": False,
     "dst": "输入文件/（任意层级）", "patterns": ["*AHP*权重*.xlsx", "*AHP*评分*.xlsx", "*权重*.xlsx", "*评分*.xlsx"],
     "how": "如果你有 CardoAHP 跑出来的权重表和评分结论，放进 输入文件/ 就行；没有跳过",
     "why": "第 9 章权重与评分；缺则用临时默认权重并标注"},
    {"key": "lowair", "name": "低空现状指标.json（CardoOSGB 成果）", "required": False,
     "dst": "输入文件/（任意层级）", "patterns": ["*现状指标*.json", "*现状指标*.xlsx"],
     "how": "从 CardoOSGB 的 过程文件/06-2.5D网格指标/ 复制《现状指标.json》放进 输入文件/；只给 xlsx 不行（留档但不解析）",
     "why": "第 3 章现状指标表；缺 json 则该表会缺，xlsx 无法替代"},
    {"key": "others", "name": "其他材料（有什么放什么）", "required": False, "catch_all": True,
     "dst": "输入文件/（任意层级）", "patterns": ["*"],
     "how": "统计年鉴、上级规划、新闻剪报、访谈录音转写、现场台账……凡是可能用得上的都能放，文件夹和 zip 也行",
     "why": "01 会自动归类；归不了的会问你怎么归。这些材料会作为事实来源和附件参考"},
]


def intake_scan():
    """逐项探测落位情况；返回 [(item, [命中文件])]。

    catch_all 项（「其他材料」）只统计没被任何具体项认领的文件，避免把已知材料重复计数。
    """
    import fnmatch

    base = ROOT / "输入文件"
    pool = [p for p in base.rglob("*")
            if p.is_file() and not p.name.startswith("~$") and "空白" not in p.name] if base.is_dir() else []
    claimed, rows = set(), []

    for item in INTAKE_ITEMS:
        if item.get("catch_all"):
            continue
        hits = [p for p in pool if any(fnmatch.fnmatch(p.name.lower(), pat.lower())
                                       for pat in item["patterns"])]
        claimed.update(hits)
        rows.append((item, sorted(set(hits))))

    for item in INTAKE_ITEMS:
        if item.get("catch_all"):
            rows.append((item, [p for p in pool if p not in claimed]))

    order = {id(item): index for index, item in enumerate(INTAKE_ITEMS)}
    return sorted(rows, key=lambda pair: order[id(pair[0])])


def write_intake_list(rows) -> Path:
    """生成人读的《放料清单.md》：要什么、放哪（绝对路径）、怎么命名、现在到没到。"""
    target = ROOT / "放料清单.md"
    ready = sum(1 for item, hits in rows if hits)
    lines = ["# 放料清单 · {}".format(SKILL.name), "",
             "> 本清单由 开工.py 自动生成，每次运行都会刷新。工作区：`{}`".format(norm(ROOT)),
             "> 放好文件后不用改本文件，重跑一次开工.py 即可。", "",
             "## 1. 现在到哪一步", "",
             "| 项 | 是否必需 | 状态 | 落位（绝对路径） |", "|---|---|---|---|"]
    for item, hits in rows:
        if hits:
            where = "、".join(norm(path.relative_to(ROOT)) if ROOT in path.parents else norm(path)
                              for path in hits[:2])
            status = "已到位（{}）".format(where)
        else:
            status = "**缺**" if item["required"] else "未提供（可选）"
        lines.append("| {} | {} | {} | `{}` |".format(
            item["name"], "必需" if item["required"] else "可选", status,
            norm(ROOT / item["dst"])))
    lines += ["", "已到位 {} / {} 项。".format(ready, len(rows)), "",
              "## 2. 每一项怎么准备", ""]
    for index, item in enumerate(INTAKE_ITEMS, start=1):
        lines += ["### {}. {}".format(index, item["name"]),
                  "- 放到：`{}`".format(norm(ROOT / item["dst"])),
                  "- 怎么来：{}".format(item["how"]),
                  "- 干什么用：{}".format(item["why"]),
                  "- 缺失后果：{}".format("不开工，提示补料" if item["required"]
                                          else "照常开工，正文写「资料缺失」并进缺项表"), ""]
    lines += ["## 3. 下一步", "",
              "```text",
              "python {}\\开工.py --intake   # 一项一项问，确认落位".format(SKILL.name),
              "python {}\\开工.py            # 直接校验并开工".format(SKILL.name),
              "```", ""]
    target.write_text("\n".join(lines), encoding="utf-8")
    return target


def run_intake(rows) -> int:
    """一项一问：给绝对路径、给准备方法、确认是否已放好。返回 0 继续，3 待补料。"""
    if not sys.stdin or not sys.stdin.isatty():
        say("  当前不是交互终端，只输出清单；请在终端里重跑 --intake，或直接按清单放料。")
        return 3
    say("")
    say("  ---- 收料引导（一项一项来；回车=已放好，输入 s=暂时没有，q=退出）----")
    for item, hits in rows:
        if hits:
            say("  [已到位] {}：{}".format(item["name"], norm(hits[0].relative_to(ROOT))))
            continue
        say("")
        say("  需要：{}（{}）".format(item["name"], "必需" if item["required"] else "可选"))
        say("  放到：{}".format(norm(ROOT / item["dst"])))
        say("  怎么来：{}".format(item["how"]))
        say("  用途：{}".format(item["why"]))
        while True:
            try:
                answer = input("  放好了吗？[回车=好了 / s=没有 / q=退出] ").strip().lower()
            except EOFError:
                # stdin 被重定向（如 < /dev/null）而 isatty() 判断漏了这种情形：
                # 没有输入可读，视为「暂时没有」，按该项的缺料口径处理，不要崩栈。
                answer = "s"
            if answer in ("", "y", "yes"):
                break
            if answer in ("s", "n", "no", "没有"):
                if item["required"]:
                    say("  ! 这一项是硬门槛，缺了不能开工。补好后重跑。")
                    return 3
                say("  · 记为缺口：正文会写「资料缺失」，不影响开工。")
                break
            if answer in ("q", "quit", "exit"):
                say("  已退出收料引导；清单已保存到 放料清单.md")
                return 3
            say("  只认 回车 / s / q。")
        found = []
        for pattern in item["patterns"]:
            found.extend(path for path in (ROOT / "输入文件").rglob(pattern)
                         if path.is_file() and not path.name.startswith("~$") and "空白" not in path.name)
        if not found:
            say("  ! 没在 输入文件/ 下找到匹配 {} 的文件{}".format(
                "、".join(item["patterns"]), "（必需项，先放料）" if item["required"] else ""))
            if item["required"]:
                return 3
    say("")
    say("  收料完成，继续校验与开工。")
    return 0


MIRRORS = [
    ("aliyun", "https://mirrors.aliyun.com/pypi/simple/", "mirrors.aliyun.com"),
    ("tsinghua", "https://pypi.tuna.tsinghua.edu.cn/simple", "pypi.tuna.tsinghua.edu.cn"),
    ("ustc", "https://mirrors.ustc.edu.cn/pypi/simple", "mirrors.ustc.edu.cn"),
    ("pypi", "https://pypi.org/simple", "pypi.org"),
]


# ---------------------------------------------------------------- 小工具

def say(msg: str = "") -> None:
    print(msg, flush=True)


def norm(p: Path) -> str:
    return str(p).replace("\\", "/")


def read_json(path: Path, default=None):
    """读 JSON。**坏文件要说一声**，不能静默返回默认值。

    早先 except 直接返回 default：工作区参数损坏 → 配置读成空 dict，
    流水线照跑，只是产出全错。静默降级在这里代价很大——参数文件是入口。
    """
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except FileNotFoundError:
        return default
    except Exception as exc:
        say(f"  警告：{norm(path)} 读取失败（{exc}），按默认值处理。"
            f"该文件若本应存在，请检查是否损坏。")
        return default


def write_json(path: Path, data) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
        f.write("\n")


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def missing_mods(py: str, mods: list[str]) -> list[str]:
    code = ("import importlib.util,json,sys\n"
            "mods=json.loads(sys.argv[1])\n"
            "print(json.dumps([m for m in mods if importlib.util.find_spec(m) is None]))")
    try:
        r = subprocess.run([py, "-c", code, json.dumps(mods)], capture_output=True,
                           text=True, timeout=90)
        return json.loads(r.stdout.strip().splitlines()[-1]) if r.returncode == 0 else list(mods)
    except Exception:
        return list(mods)


def venv_python() -> Path:
    name = "python.exe" if os.name == "nt" else "python"
    sub = "Scripts" if os.name == "nt" else "bin"
    return venv_dir() / sub / name


DELIVERY_MARKER = ".cardo_delivery_marker"


def _is_delivery_root(ws: Path) -> bool:
    """判断工作区是否为「交付库根」：交付库只读出包，不能在里面开工。

    唯一判据：**当前目录本身**存在标记文件 .cardo_delivery_marker。
    为什么不看「根下有 ≥2 个 Cardo* 目录」：用户经常把几个技能包并列放在
    同一个项目目录（如 测试-全流程贯通\）里使用，那是正常用法，不能拦。
    """
    try:
        return (ws / DELIVERY_MARKER).exists()
    except OSError:
        return False


def _has_multi_packages(ws: Path) -> bool:
    """并列多个 Cardo* 目录只作提示，不作拒绝条件。"""
    try:
        pkgs = [d for d in ws.iterdir() if d.is_dir() and d.name.startswith("Cardo")]
        return len(pkgs) >= 2
    except OSError:
        return False


ROUND_FILE = "记忆与规则/轮次.json"


def check_round(root: Path, round_arg: str) -> str:
    """开工自检：输出根必须等于本次轮次目录，不相等就拒绝启动。

    为什么必须查：工作区默认取「技能包所在的上一级目录」，本身没有轮次概念。
    2026-11-14 实测踩过：在 测试建议书\\R2 里直接重跑，产物全部写回 R2，
    把 R2 的送审原件覆盖掉，R2 已无法回溯。轮次目录必须显式声明、显式校验，
    不能靠"这次记得传 --workspace"。
    """
    record_path = root / ROUND_FILE
    record = {}
    if record_path.is_file():
        try:
            record = json.loads(record_path.read_text(encoding="utf-8")) or {}
        except (OSError, ValueError):
            record = {}
    declared = (round_arg or "").strip() or str(record.get("round") or "")
    if not declared:
        # 没声明轮次：只在「干净的工作区」放行。已有成果/过程文件说明这是一个
        # 跑过的轮次目录，此时不声明就开工＝覆盖上一轮，必须显式给 --round。
        dirty = [d for d in ("成果文件", "过程文件")
                 if (root / d).is_dir() and any((root / d).rglob("*.*"))]
        if dirty:
            say(f"[错误] 输出根 {root} 没有轮次声明，但已有 {'、'.join(dirty)}。")
            say("        本次开工会覆盖该轮产物。请显式声明本轮轮次目录，例如：")
            say("          python 开工.py --round R3 --workspace <本轮轮次目录>")
            raise SystemExit(2)
        declared = root.name
    if root.name != declared:
        say(f"[错误] 轮次自检不通过：本次轮次 {declared}，输出根却是 {root}")
        say("        输出根必须等于本次轮次目录；请把 --workspace 指到该轮次目录后再跑。")
        raise SystemExit(2)
    rec_ws = str(record.get("workspace") or "")
    if rec_ws and Path(rec_ws).resolve() != root.resolve():
        say(f"[错误] 轮次记录属于另一个目录：{rec_ws}")
        say(f"        本目录是 {root}。复制工作区后请先删除 {ROUND_FILE} 再开工。")
        raise SystemExit(2)
    write_json(record_path, {"round": declared, "workspace": str(root),
                             "startedAt": datetime.now().strftime("%Y-%m-%d %H:%M:%S")})
    return declared


def resolve_workspace(value: str) -> Path:
    """工作区 = 技能包所在的上一级目录（可 --workspace 覆盖）；技能包只读。"""
    ws = Path(value).resolve() if value else SKILL.parent
    if _is_delivery_root(ws):
        say(f"[错误] 这是交付库根目录，不能在这里开工：{ws}")
        say("        交付库是只读的发布产物，开工会在库里生成 过程文件/成果文件。")
        say("        正确做法：把技能包复制到项目目录后再运行，例如")
        say("          <项目目录> 下放 CardoInvest\\，再在该目录运行开工.py")
        raise SystemExit(2)
    if ws == SKILL or SKILL in ws.parents:
        say(f"[错误] 工作区不能是技能包目录或其子目录：{ws}")
        raise SystemExit(2)
    if not ws.is_dir():
        say(f"[错误] 工作区不存在：{ws}")
        raise SystemExit(2)
    if _has_multi_packages(ws):
        say("[提示] 当前目录并列了多个 Cardo* 技能包；未发现 .cardo_delivery_marker，按普通工作区继续。")
    if len(ws.parts) <= 1:
        # 只有当工作区本身就是盘符根（如 D://）时才拦；判断看自身层级，
        # 不能用 ws.parent.name —— 那样 D://任意文件夹 的父目录也是 D://，会被误拦。
        say("[错误] 技能包不能直接放在盘符根目录；请放进一个项目文件夹后再跑。")
        raise SystemExit(2)
    return ws


# ---------------------------------------------------------------- 环境

def make_dirs() -> list[str]:
    made = []
    for d in DIRS:
        p = ROOT / d
        if not p.is_dir():
            p.mkdir(parents=True, exist_ok=True)
            made.append(d)
    return made


def release_assets() -> list[str]:
    problems = []
    for a in ASSETS:
        src, dst = SKILL / a["src"], ROOT / a["dst"]
        if not src.is_file():
            problems.append(f"包内缺资产：{a['src']}（包被裁剪，需重新解压原包）")
            continue
        if dst.is_file():
            if a.get("sha256") and sha256_file(dst).upper() != a["sha256"].upper():
                problems.append(f"模板被改动：{a['dst']}（不要替换标准模板；恢复方法：删掉该文件后重跑本命令）")
            continue
        dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src, dst)
        say(f"  释放模板：{a['dst']}")
    for a in INPUT_TEMPLATES:
        src, dst = SKILL / a["src"], ROOT / a["dst"]
        if src.is_file() and not dst.is_file():
            dst.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(src, dst)
            say(f"  释放输入模板：{a['dst']}")
    return problems


def install(py: str) -> bool:
    req = SKILL / "requirements.txt"
    if not req.is_file():
        return True
    for name, url, host in MIRRORS:
        say(f"  安装依赖（{name}）…")
        r = subprocess.run([py, "-m", "pip", "install", "-r", str(req), "-i", url,
                            "--trusted-host", host, "--disable-pip-version-check", "--no-input"],
                           cwd=str(ROOT), capture_output=True, text=True, timeout=900)
        if r.returncode == 0:
            return True
        say(f"  {name} 失败，换下一个源。")
    return False


def system_pythons() -> list[str]:
    """找**真正的**系统 Python 候选（排除便携解释器本身）。

    为什么需要：便携解释器是内嵌的精简版，**没有 venv 模块**（实测）。
    若用它 `python -m venv` 必然失败，报错会指向"创建 venv 失败"，
    掩盖真正的问题（便携环境缺哪个包）。建 venv 必须用系统 Python。
    """
    out: list[str] = []
    self_path = str(RUNTIME_PY.resolve()) if RUNTIME_PY.is_file() else ""
    seen = set()

    def add(p: str) -> None:
        try:
            rp = str(Path(p).resolve())
        except OSError:
            return
        if rp in seen or (self_path and rp == self_path):
            return
        seen.add(rp)
        out.append(p)

    if self_path and str(Path(sys.executable).resolve()) != self_path:
        add(sys.executable)
    if os.name == "nt":
        for name in ("python.exe", "python3.exe"):
            found = shutil.which(name)
            if found:
                add(found)
        for base in (os.environ.get("LOCALAPPDATA", ""),
                     os.environ.get("ProgramFiles", ""),
                     r"C://Python312", r"C://Python311", r"C://Python310"):
            if not base:
                continue
            for minor in range(9, 14):
                cand = Path(base) / "Programs" / "Python" / f"Python3{minor}" / "python.exe"
                if cand.is_file():
                    add(str(cand))
    else:
        for name in ("python3", "python"):
            found = shutil.which(name)
            if found:
                add(found)
    return out


def check_skill_package() -> list[str]:
    """技能包完整性自检：返回问题清单（空 = 完整）。只读检查，不改任何文件。"""
    problems: list[str] = []
    skills_root = SKILL / "Skills"
    if not skills_root.is_dir():
        return ["Skills/ 目录整体缺失（技能包的核心）。"
                "从交付包完整复制 CardoInvest/Skills/ 到 {}".format(norm(skills_root))]
    for step, scripts, rules in SKILL_STEPS:
        step_dir = skills_root / step
        if not step_dir.is_dir():
            problems.append("Skills/{}/ 目录缺失。".format(step))
            continue
        scripts_dir = step_dir / "scripts"
        if not scripts_dir.is_dir():
            problems.append("Skills/{}/scripts/ 目录缺失。".format(step))
        else:
            for s in scripts:
                if not (scripts_dir / s).is_file():
                    problems.append("Skills/{}/scripts/{} 缺失。".format(step, s))
        for r in rules:
            if not (step_dir / "规则" / r).is_file():
                problems.append("Skills/{}/规则/{} 缺失。".format(step, r))
    # 05 的十个章节副本（每章一份章节契约：节标题与必答要点从中读取）
    for name in CHAPTER_SKILLS:
        if not (skills_root / name / "规则" / "章节契约.yaml").is_file():
            problems.append("Skills/{}/规则/章节契约.yaml 缺失。".format(name))
    return problems


def check_runtime() -> list[str]:
    """便携环境完整性自检：返回问题清单（空 = 完整）。只读检查。"""
    problems: list[str] = []
    if not RUNTIME_PY.is_file():
        # 派生仓库不带 runtime：有系统 Python 时按 requirements.txt 建虚拟环境即可
        if Path(sys.executable).is_file():
            return problems
        problems.append("runtime/python/ 缺失（便携 Python 解释器），且未找到系统 Python。"
                        "从交付包完整复制 CardoInvest/runtime/ 到 {}".format(norm(SKILL / "runtime")))
        return problems
    miss = missing_mods(str(RUNTIME_PY), RUNTIME_MODULES)
    if miss:
        site = RUNTIME_PY.parent / "Lib" / "site-packages"
        problems.append(
            "runtime 的 site-packages 不完整，缺 {}（共 {} 个包）。"
            "多半是复制时漏了：从交付包**重新完整复制** CardoInvest/runtime/ "
            "覆盖到 {}".format("、".join(miss), len(RUNTIME_MODULES), norm(SKILL / "runtime")))
        if not (site / "yaml").is_dir():
            problems.append("特别提示：缺 PyYAML 时 05 步会读不到章节契约，"
                            "十章成文直接失败。")
    return problems


def report_skill_problems(problems: list[str], title: str) -> None:
    """把完整性问题说清楚：缺什么、从哪补、怎么补。"""
    say("[自检] {}不完整，缺 {} 项：".format(title, len(problems)))
    for p in problems[:12]:
        say("  · {}".format(p))
    if len(problems) > 12:
        say("  · …另有 {} 项".format(len(problems) - 12))
    say("")
    say("  补法：从完整的交付包重新复制整个目录覆盖过去（推荐）：")
    say("     xcopy /E /I /Y <交付包>\\CardoInvest\\{} <工作区>\\CardoInvest\\{}".format(
        "Skills" if "技能" in title or "Skills" in problems[0] else "runtime",
        "Skills" if "技能" in title or "Skills" in problems[0] else "runtime"))
    say("  补齐后重跑本命令即可；不需要重装依赖、不需要重建工作区。")


def report_manual_install(missing: list[str]) -> None:
    """缺依赖时给出明确的手工补法（离线环境可用）。

    便携环境的 site-packages 路径是固定的，拷 wheel 进去即可，不需要建 venv、
    不需要联网——这是包内便携环境的本意。
    """
    site = RUNTIME_PY.parent / "Lib" / "site-packages"
    if os.name != "nt":
        site = RUNTIME_PY.parent / "lib" / "python3" / "site-packages"
    say("")
    say("  缺这些包：{}".format("、".join(missing)))
    say("  离线补法（任选其一）：")
    say("  A) 拷 wheel 到便携环境（推荐，无需联网）：")
    say("     把对应的 .whl 放进 {} 后重跑本命令".format(norm(site)))
    say("  B) 用系统 Python：")
    say("     python -m pip install {}（装到系统 Python 即可，流水线会自动改用它）".format(
        " ".join(missing)))
    say("  C) 有网络时：直接重跑本命令，会自动安装。")


def ensure_env(no_install: bool) -> tuple[str | None, bool]:
    """选一个能用的解释器。返回 (解释器, 是否刚装过)；None 表示环境不可用。

    选解释器的顺序（2026-11-01 调整）：
      1) **包内便携环境**：拷包即用，不联网、不建 venv（首选）。
      2) **系统 Python**：便携环境缺依赖时优先用它——
         便携解释器是内嵌的精简版，**没有 venv 模块**（实测），
         拿它 `python -m venv` 必然失败（exit 2，且报错指向"创建 venv 失败"，
         掩盖真正的问题：便携环境缺哪个包）。
      3) **联网安装**：系统 Python 也没有依赖时，才建 venv 并装（需联网）。
      4) 都不行 → 报清楚缺哪个包、怎么补（拷 wheel 到 runtime/Lib/site-packages）。
    """
    # 1) 包内便携环境
    if RUNTIME_PY.is_file():
        miss = missing_mods(str(RUNTIME_PY), IMPORTS)
        if not miss:
            return str(RUNTIME_PY), False
        say(f"  便携环境缺依赖：{'、'.join(miss)}")
    # 2) 系统 Python 已有依赖 → 直接用
    sys_miss = list(IMPORTS)
    for cand in system_pythons():
        cmiss = missing_mods(cand, IMPORTS)
        if not cmiss:
            say(f"  改用系统 Python（依赖齐全）：{cand}")
            return cand, False
        sys_miss = cmiss
    # 3) 已有 .cardo/venv-invest 且依赖齐全
    py = venv_python()
    if py.is_file() and not missing_mods(str(py), IMPORTS):
        return str(py), False
    if no_install:
        say(f"  缺依赖：{'、'.join(sys_miss)}（已指定 --no-install，不安装）")
        report_manual_install(sys_miss)
        return None, False
    # 4) 联网安装：建 venv **必须用系统 Python**——便携解释器没有 venv 模块。
    if not py.is_file():
        say("  建虚拟环境 .cardo/venv-invest …")
        venv_dir().parent.mkdir(parents=True, exist_ok=True)
        base = next((c for c in system_pythons()
                     if missing_mods(c, ["venv"]) == []), None)
        if base is None:
            say("  找不到带 venv 模块的系统 Python，无法自动建虚拟环境。")
            report_manual_install(sys_miss)
            return None, False
        r = subprocess.run([base, "-m", "venv", str(venv_dir())],
                           cwd=str(ROOT), capture_output=True, text=True, timeout=600)
        if r.returncode != 0:
            say("  创建虚拟环境失败：")
            say((r.stderr or r.stdout or "")[-800:])
            report_manual_install(sys_miss)
            return None, False
    if not install(str(py)):
        say("  依赖安装失败。离线环境可手动拷 wheel 后执行：")
        say("  python -m pip install --no-index --find-links <目录> -r requirements.txt")
        report_manual_install(sys_miss)
        return None, False
    left = missing_mods(str(py), IMPORTS)
    if left:
        say(f"  安装后仍缺：{'、'.join(left)}")
        report_manual_install(left)
        return None, False
    return str(py), True


# ---------------------------------------------------------------- 参数与输入

def load_params() -> dict:
    data = read_json(ROOT / PARAMS, {}) or {}
    data.setdefault("schema", SCHEMA)
    for k, v in DEFAULTS.items():
        data.setdefault(k, v)
    if isinstance(data.get("paths"), dict):
        for k, v in DEFAULTS["paths"].items():
            data["paths"].setdefault(k, v)
    return data


def find_financial() -> list[Path]:
    root = ROOT / "输入文件"
    if not root.is_dir():
        return []
    hits = []
    for p in root.rglob("*财务参数*.xlsx"):
        if not p.is_file() or p.name.startswith("~$") or "空白" in p.name:
            continue
        hits.append(p)
    return sorted(hits)


def check_financial(p: Path) -> tuple[bool, list[str]]:
    """返回 (是否可用, 说明行)。"""
    try:
        import openpyxl
    except Exception:
        # 缺 openpyxl → **判不可用**。早先返回 (True, "跳过检查")，
        # 等于「没检查就说通过」——财务参数表结构对不对根本没人知道。
        return False, ["未装 openpyxl，无法校验财务参数表结构；"
                       "请装 openpyxl 或把财务参数表先补齐"]
    try:
        wb = openpyxl.load_workbook(str(p), read_only=True, data_only=True)
    except Exception as e:
        return False, [f"无法读取 xlsx：{e}"]
    miss = [s for s in SHEETS if s not in wb.sheetnames]
    if miss:
        return False, ["缺工作表：" + "、".join(miss)]
    rows = []
    ws = wb["工程费"]
    header_seen = False
    filled = 0
    for row in ws.iter_rows(values_only=True):
        vals = [("" if v is None else str(v).strip()) for v in row]
        if not header_seen:
            if any("名称" in v for v in vals) and any("金额" in v for v in vals):
                header_seen = True
            continue
        if any(vals):
            text = [v for v in vals if v]
            if len(text) >= 2 and not all(v in ("合计", "小计") for v in vals):
                filled += 1
    if filled < 1:
        return False, ["工程费表未填：至少需要 1 行（名称 + 金额(万元)）"]
    rows.append(f"9 张表齐全；工程费已填 {filled} 行")
    return True, rows


def project_name_from(p: Path) -> str:
    try:
        import openpyxl
        wb = openpyxl.load_workbook(str(p), read_only=True, data_only=True)
        ws = wb["项目"]
        for row in ws.iter_rows(values_only=True):
            cells = [("" if v is None else str(v).strip()) for v in row]
            for i, v in enumerate(cells):
                if "项目名称" in v:
                    for nxt in cells[i + 1:]:
                        if nxt:
                            return nxt
    except Exception:
        return ""
    return ""


def precheck(params: dict) -> tuple[list[str], list[str]]:
    blockers, notes = [], []
    tf = params.get("template_file") or params.get("paths", {}).get("template")
    if tf and not (ROOT / tf).is_file() and (ROOT / TEMPLATE_DST).is_file():
        # 参数里记的是旧路径（模板改名/挪位）：回落到标准位置，不让老参数卡住开工
        params.setdefault("paths", {})["template"] = TEMPLATE_DST
        params["template_file"] = TEMPLATE_DST
        tf = TEMPLATE_DST
    if tf and not (ROOT / tf).is_file():
        blockers.append(f"模板不在位：{tf}（删掉后重跑本命令可恢复标准模板）")

    hits = find_financial()
    if not hits:
        blockers.append("缺必需输入：财务参数.xlsx（用 输入文件-模板/财务参数-空白.xlsx 填好后另存为 输入文件/财务参数.xlsx）")
    elif len(hits) > 1:
        names = "、".join(norm(p.relative_to(ROOT / "输入文件")) for p in hits)
        blockers.append(f"财务参数有多份（{names}）：只保留本次要用的那一份，或改名为 输入文件/财务参数.xlsx")
    else:
        ok, msgs = check_financial(hits[0])
        for m in msgs:
            (notes if ok else blockers).append(f"{norm(hits[0].relative_to(ROOT))}：{m}")
        if ok and not params.get("projectName"):
            name = project_name_from(hits[0])
            if name:
                params["projectName"] = name
                notes.append(f"已从「项目」表登记项目名：{name}")

    lowair_json = list((ROOT / "输入文件").rglob("*现状指标*.json"))
    lowair_xlsx = list((ROOT / "输入文件").rglob("*现状指标*.xlsx"))
    lowair_ok = lowair_json or (ROOT / "过程文件/00-总控/上游机读件/现状指标.json").is_file()
    optional = {
        "现状分析报告": list((ROOT / "输入文件").rglob("*现状分析*.docx")),
        "AHP 权重/评分": list((ROOT / "输入文件").rglob("*权重*.xlsx")) + list((ROOT / "输入文件").rglob("*评分*.xlsx")),
        "低空现状指标（json）": lowair_json,
        "可研": list((ROOT / "输入文件").rglob("*可研*.docx")),
        "槽位表": [p for p in (ROOT / "输入文件").rglob("*槽位*.xlsx") if "空白" not in p.name],
    }
    found = [k for k, v in optional.items() if v]
    notes.append("可选件：" + ("、".join(found) if found else "未提供（缺项会登记在缺项表，不编造）"))
    if lowair_xlsx and not lowair_ok:
        # xlsx 只做人读留档，本包不解析；说清"不在输入范围"，不要留一句"会缺"吓用户。
        notes.append(
            "低空现状指标：已检测到《现状指标.xlsx》"
            f"（{norm(lowair_xlsx[0].relative_to(ROOT))}）；该文件不在本包输入范围"
            "（本包只解析 JSON），留作人读存档，不影响出件。"
        )
    if not params.get("projectName"):
        blockers.append('缺参数 projectName：用 --project "项目全称" 指定')
    return blockers, notes


# ---------------------------------------------------------------- 联网补料

RESEARCH_DIR = "过程文件/00-总控/网络补料"
RESEARCH_ROWS = [
    ("工程费单价（元/㎡）", "是（公开造价 / 中标价）",
     "{城市} 老旧小区改造 单位造价 元/平方米｜{项目} 中标公告 工程量清单",
     "分项、含税口径、年份、发布机构、链接"),
    ("其他费费率", "是（政策文件）",
     "建设工程其他费用 费率 财建〔2016〕504号｜{城市} 前期工作费 费率",
     "文号、费率、计费基数、链接"),
    ("预备费 / 建设期利息", "是（政策 + 市场）",
     "基本预备费 费率 政府投资项目｜LPR 贷款市场报价利率 {年份}",
     "费率或利率、取价日期、来源"),
    ("收入 / 成本 / 税费", "部分（市场行情）",
     "{项目} 租金 水平｜{城市} 物业服务费 标准｜增值税 税率 建筑服务",
     "单价、税率、取价日期、来源"),
    ("资金结构（资本金比例）", "是（政策）",
     "城市更新 项目 资本金比例 规定 {城市}",
     "比例、文号、发布机构、链接"),
]


def check_research() -> tuple[list[str], list[str]]:
    """校验 网络补料.json：每条必须有值、口径、可核来源。"""
    p = ROOT / RESEARCH_DIR / "网络补料.json"
    if not p.is_file():
        return [], []
    data = read_json(p, {}) or {}
    items = data.get("items") or []
    blockers, notes = [], []
    for i, it in enumerate(items, 1):
        key = it.get("key") or f"补料#{i}"
        if it.get("not_found"):
            notes.append(f"{key}：标记查无公开数据（按缺口登记，不估算）")
            continue
        miss = [f for f in ("value", "unit", "source_url", "publisher", "date") if not it.get(f)]
        if miss:
            blockers.append(f"网络补料「{key}」缺字段：{'、'.join(miss)}")
            continue
        if not re.match(r"https?://", str(it["source_url"])):
            blockers.append(f"网络补料「{key}」source_url 不是 http(s) 链接")
            continue
        notes.append(f"{key}={it['value']}{it['unit']}（{it['publisher']} {it['date']}）"
                     + ("已确认" if it.get("confirmed_by_user") else "未确认，仅作参考值"))
    return blockers, notes


def research_guide(params: dict, blockers: list[str]) -> int:
    """缺料时给「联网检索补」的可执行清单；只给方向与来源要求，不生成数值。"""
    d = ROOT / RESEARCH_DIR
    d.mkdir(parents=True, exist_ok=True)
    name = params.get("projectName") or "本项目"
    lines = [
        f"# 联网补料清单 · {name}", "",
        "> 本清单只给检索方向与来源要求，不自动生成数值。检索结果逐条回填 `网络补料.json`，",
        "> 再按「落位」写入正式参数表；未核实来源不得当作已确认值，缺的照样留空登记缺口。", "",
    ]
    if blockers:
        lines += ["## 当前阻塞项", ""] + [f"- {b}" for b in blockers] + [""]
    lines += ["## 检索方向", "", "| 缺项 | 能否联网补 | 建议检索关键词 | 需记录 |", "|---|---|---|---|"]
    for what, how, kw, fields in RESEARCH_ROWS:
        lines.append(f"| {what} | {how} | {kw.format(项目=name, 城市='项目所在城市', 年份='2024-2025')} | {fields} |")
    lines += ["", "## 落位方式", "",
              "- 关键词里的「项目所在城市」先替换成实际地名再检索。",
              "- 先复制 `输入文件-模板/财务参数-空白.xlsx` 为 `输入文件/财务参数.xlsx`，逐格填「参考值」；",
              "  在「项目」表或该格备注写清 发布机构 + URL + 取价日期（如 LPR 必须写取价日）。",
              "- 拿不准的格留空：01 会登记为缺口，不编造、不按经验补。",
              "- 用户确认过的值，在 `网络补料.json` 里把 `confirmed_by_user` 置 true。", ""]
    (d / "补料清单.md").write_text("\n".join(lines), encoding="utf-8")
    js = d / "网络补料.json"
    if not js.is_file():
        write_json(js, {
            "schema": "cardo/research-supplement/v1",
            "project": name,
            "_说明": "逐条填 items；每项须有 key/value/unit/source_url/publisher/date；"
                     "查无公开数据写 not_found: true；用户确认后把 confirmed_by_user 置 true。",
            "items": [],
        })
    say(f"[补料] 清单：{RESEARCH_DIR}/补料清单.md")
    say(f"[补料] 回填：{RESEARCH_DIR}/网络补料.json（留空也可，缺口照常登记）")
    rb, rn = check_research()
    for x in rn:
        say(f"[补料] {x}")
    for x in rb:
        say(f"[补料][错误] {x}")
    return 3


# ---------------------------------------------------------------- 主流程

def run_pipeline(cmd: list[str], env: dict) -> tuple[int, str]:
    """边跑边打印，同时留全文供判定停顿/失败位置。"""
    lines = []
    p = subprocess.Popen(cmd, cwd=str(ROOT), env=env, stdout=subprocess.PIPE,
                         stderr=subprocess.STDOUT, text=True, encoding="utf-8", errors="replace")
    for line in p.stdout:
        say(line.rstrip("\n"))
        lines.append(line)
    return p.wait(), "".join(lines)


def cmd_hint(*args: str) -> str:
    """给用户看的一条可直接粘贴的命令。

    优先用包内便携 python（runtime/python）—— 没装 Python 的机器也能跑；
    包里没有 runtime 时才退回系统 python。
    """
    if RUNTIME_PY.is_file():
        try:
            exe = str(RUNTIME_PY.relative_to(ROOT))
        except ValueError:
            exe = str(RUNTIME_PY)
        if os.name == "nt":
            exe = exe.replace("/", "\\")
    else:
        exe = "python"
    parts = [exe, "{}\\开工.py".format(SKILL.name)]
    parts.extend(args)
    return " ".join(parts)


def write_board_status() -> None:
    """把工作区真实进度写进看板的 WORKSPACE_STATUS 块（顶栏据此显示真实状态）。

    看板是单文件 HTML，双击打开读不到工作区里的文件，所以状态只能由 开工.py 注入。
    只替换标记之间的那一段，其余字节原样保留；写不进去（包只读、看板被换掉等）
    就跳过 —— 看板会退回手动勾选计数，不影响主流程。
    """
    if not WORKSPACE_RESOLVED:
        return                      # 工作区没解析成功，别拿默认 ROOT 写错状态
    board = SKILL / "开工看板.html"
    if not board.is_file():
        return
    try:
        with open(board, "r", encoding="utf-8", newline="") as f:
            text = f.read()
    except OSError:
        return
    if STATUS_BEGIN not in text or STATUS_END not in text:
        return
    try:
        status = {
            "updatedAt": datetime.now().astimezone().isoformat(timespec="seconds"),
            "paramsFilled": (ROOT / "输入文件/财务参数.xlsx").is_file(),
            "weightsComputed": (ROOT / "过程文件/03-财务评价/财务测算.json").is_file(),
            "deliverables": sum(1 for pat in DELIVERABLES for _ in ROOT.glob(pat)),
        }
    except Exception:
        return
    # </script> 出现在 JSON 里会提前闭合脚本块 —— 转义 "<" 即可
    payload = json.dumps(status, ensure_ascii=False, indent=2).replace("<", "\\u003c")
    eol = "\r\n" if "\r\n" in text else "\n"
    block = (STATUS_BEGIN + eol + "<script>window.WORKSPACE_STATUS = "
             + payload + ";</script>" + eol + STATUS_END)
    head, _, rest = text.partition(STATUS_BEGIN)
    _, _, tail = rest.partition(STATUS_END)
    try:
        with open(board, "w", encoding="utf-8", newline="") as f:
            f.write(head + block + tail)
    except OSError:
        pass


CRASH_LOG = "过程文件/00-总控/异常日志.log"

# 异常类型 → (缺什么, 建议怎么处理)。只覆盖用户真会遇到的几类，
# 其余一律按「程序内部错误」处理，保留日志交给维护人。
_CRASH_HINTS = {
    "ModuleNotFoundError": ("缺少 Python 依赖（包内运行时可能不完整）",
                            "确认 runtime/ 与 Skills/ 是整体复制的；或用系统 Python 按 requirements.txt 装依赖"),
    "ImportError": ("运行环境不完整（缺模块或版本不符）", "重新完整复制 runtime/ 后重跑"),
    "FileNotFoundError": ("需要的文件不存在", "按上面给出的路径补齐后重跑"),
    "PermissionError": ("文件被占用或没有写权限",
                        "关掉正在打开该文件的 Word / Excel / 资源管理器预览窗格后重跑"),
    "UnicodeDecodeError": ("文件不是 UTF-8 编码", "用记事本另存为 UTF-8 后重跑"),
    "KeyError": ("输入表缺少必需的列或字段", "对照 放料清单.md 补齐该列后重跑"),
    "JSONDecodeError": ("JSON 文件内容损坏", "删掉该文件让程序重新生成，或恢复备份"),
}


def _last_step_hint() -> str:
    """从运行状态/运行日志里找最后一个进入的步骤，回答「哪一步失败」。"""
    for name in ("运行状态.json", "运行日志.json"):
        path = ROOT / "过程文件/00-总控" / name
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            continue
        if not isinstance(data, dict):
            continue
        if name == "运行状态.json":
            steps = data.get("steps") or {}
            for key in list(steps)[::-1]:
                st = steps.get(key) or {}
                if isinstance(st, dict) and st.get("status") in ("running", "failed"):
                    return "{} {}".format(key, st.get("name", "")).strip()
        else:
            entries = data.get("entries") or []
            for item in entries[::-1]:
                if isinstance(item, dict) and item.get("step"):
                    return "{} {}".format(item.get("step"), item.get("name", "")).strip()
    return "开工准备阶段"


def crash_report(exc: BaseException) -> int:
    """未预期异常的统一出口：人话四段 + 落日志；--debug 时保留原始堆栈。"""
    # 用 format_exception 而不是 format_exc：后者不在 except 块里调用时只会写出
    # "NoneType: None"，日志等于没记。这里始终带上 exc 自身的堆栈。
    tb = "".join(traceback.format_exception(type(exc), exc, exc.__traceback__))
    log = ROOT / CRASH_LOG
    try:
        log.parent.mkdir(parents=True, exist_ok=True)
        with log.open("a", encoding="utf-8") as f:
            f.write("\n===== {} =====\n".format(datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
            f.write("技能包：{}\n".format(SKILL.name))
            f.write("命令：{}\n".format(" ".join(sys.argv)))
            f.write(tb)
    except OSError:
        log = None
    if "--debug" in sys.argv or os.environ.get("CARDO_DEBUG"):
        raise exc
    name = type(exc).__name__
    what, action = _CRASH_HINTS.get(
        name, ("程序内部错误", "把异常日志发给维护人；急用可先 --from <上一步> 续跑"))
    say("")
    say("[异常] 程序没有正常跑完")
    say("  哪一步失败：{}".format(_last_step_hint()))
    say("  缺什么：{}（{}: {}）".format(what, name, exc))
    say("  建议怎么处理：{}".format(action))
    say("  日志在哪：{}".format(norm(log.relative_to(ROOT)) if log else "（日志写入失败）"))
    say("  要看完整堆栈：命令末尾加 --debug 重跑")
    return 1


def main(argv=None) -> int:
    """跑完（含各种提前返回）后刷新看板顶栏的真实状态。"""
    try:
        return _main(argv)
    except SystemExit:
        raise
    except KeyboardInterrupt:
        say("\n[中断] 已停止；已完成步骤的产物保留，续跑用 --from <步>。")
        return 130
    except BaseException as exc:  # noqa: BLE001
        return crash_report(exc)
    finally:
        try:
            write_board_status()
        except Exception:
            pass


def _main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=f"{TITLE} · 一条命令开工")
    ap.add_argument("--debug", action="store_true", help="出错时打印完整堆栈（默认只给人话提示）")
    ap.add_argument("--project", default="", help="项目全称（写入 projectName）")
    ap.add_argument("--from", dest="start_from", default="", help="从第几步开始（续跑用，如 01）")
    ap.add_argument("--to", dest="stop_at", default="", help="跑到第几步为止（如 08）")
    ap.add_argument("--research", action="store_true",
                    help="缺料时给「联网补料清单」：检索关键词、来源要求、落位方式（不自动生成数值）")
    ap.add_argument("--intake", action="store_true",
                    help="收料引导：一项一项问，给绝对路径与准备方法，确认后再校验开工")
    ap.add_argument("--check", action="store_true", help="只做结构与输入校验，不跑流水线")
    ap.add_argument("--no-install", action="store_true", help="缺依赖时不安装")
    ap.add_argument("--dry-run", action="store_true", help="只打印将要执行的命令")
    ap.add_argument("--workspace", default="", help="工作区根（缺省为技能包所在目录，即上一级）")
    ap.add_argument("--round", default="",
                    help="本轮轮次目录名（如 R3）：输出根必须等于它，不相等拒绝启动")
    ap.add_argument("passthrough", nargs=argparse.REMAINDER, help="其余参数原样传给流水线")
    args = ap.parse_args(argv)

    if not os.environ.get("CARDO_REEXEC"):
        say(f"== {TITLE} · 开工 ==")
    if sys.version_info < PY_MIN:
        say(f"[环境] 需要 Python 3.{PY_MIN[1]} 及以上，当前 {sys.version.split()[0]}。")
        return 2

    global ROOT, WORKSPACE_RESOLVED
    ROOT = resolve_workspace(args.workspace)
    WORKSPACE_RESOLVED = True
    round_name = check_round(ROOT, args.round)
    if not os.environ.get("CARDO_REEXEC"):
        say(f"[0/4] 工作区：{ROOT}（技能包只读：{SKILL.name}）")
        say(f"      轮次自检：{round_name} = 输出根目录名，通过")

    # 技能包完整性自检：放在依赖检查之前。
    # 技能包被复制不完整时（实测只拷了 开工.py + assets + runtime，
    # 或 runtime 只拷了一部分——site-packages 从 49 个包变成 12 个），
    # 若不先查，后面会一路跑到很后面才失败，报错还不指向根因。
    skill_problems = check_skill_package()
    if skill_problems:
        if not os.environ.get("CARDO_REEXEC"):
            report_skill_problems(skill_problems, "技能包")
        return 2
    runtime_problems = check_runtime()
    if runtime_problems:
        if not os.environ.get("CARDO_REEXEC"):
            report_skill_problems(runtime_problems, "便携运行环境")
        return 2

    if not os.environ.get("CARDO_REEXEC"):
        board = SKILL / "开工看板.html"
        if board.is_file():
            say("      开工看板：{}（双击打开：步骤、放料清单、可复制提示词）".format(norm(board)))

    py, installed = ensure_env(args.no_install)
    if py is None:
        return 2
    if py != sys.executable:
        say(f"[1/4] 依赖：改用 {py}")
        env = dict(os.environ, PYTHONIOENCODING="utf-8", PYTHONUTF8="1",
                   PYTHONDONTWRITEBYTECODE="1", CARDO_REEXEC="1")
        r = subprocess.run([py, str(Path(__file__).resolve())] + list(sys.argv[1:]),
                           cwd=str(ROOT), env=env)
        return r.returncode
    if py == str(RUNTIME_PY):
        say("[1/4] 依赖：就位（包内便携环境，无需联网）")
    else:
        say("[1/4] 依赖：就位" + ("（已装入 .venv）" if installed else ""))

    made = make_dirs()
    problems = release_assets()
    if problems:
        for p in problems:
            say(f"  ✗ {p}")
        return 2
    say(f"[2/4] 目录：{'新建 ' + '、'.join(made) if made else '已就绪'}；标准模板在位且 sha 一致")

    rows = intake_scan()
    ready = sum(1 for item, hits in rows if hits)
    list_path = write_intake_list(rows)
    say(f"  放料清单：{norm(list_path.relative_to(ROOT))}（已到位 {ready}/{len(rows)} 项）")
    if args.intake:
        code = run_intake(rows)
        rows = intake_scan()
        write_intake_list(rows)
        if code != 0:
            say(f"  按清单放料后重跑：{cmd_hint()}，或再跑 --intake 逐项确认")
            return code

    say("[3/4] 参数与输入")
    params = load_params()
    if args.project:
        params["projectName"] = args.project
    write_json(ROOT / PARAMS, params)
    blockers, notes = precheck(params)
    rb, rn = check_research()
    blockers += rb
    notes += rn
    write_json(ROOT / PARAMS, params)
    for n in notes:
        say(f"  · {n}")
    if args.research:
        research_guide(params, blockers)
    if blockers:
        for b in blockers:
            say(f"  ✗ {b}")
        say(f"  放料目录（绝对路径）：{norm(ROOT / '输入文件')}")
        say(f"  逐项确认：{cmd_hint('--intake')}")
        say(f"  补齐后重跑：{cmd_hint()}")
        if not args.research:
            say(f"  或缺料想联网补：{cmd_hint('--research')}"
                "（给检索清单与来源要求，结果须标来源并经用户确认）")
        return 3
    if not params.get("projectName"):
        say("  ✗ 缺参数 projectName：用 --project \"项目全称\" 指定")
        if not args.research:
            say(f"  或缺料想联网补：{cmd_hint('--research')}")
        return 3
    say(f"  项目：{params['projectName']}")

    if args.check:
        say("  校验通过（--check 不跑流水线）。")
        return 0

    say("[4/4] 流水线")
    cmd = [py, str(SKILL / PIPELINE), "run", "--workspace", "."]
    if args.start_from:
        cmd += ["--from", args.start_from]
    if args.stop_at:
        cmd += ["--to", args.stop_at]
    cmd += list(args.passthrough)
    say("  " + " ".join(cmd))
    if args.dry_run:
        return 0

    env = dict(os.environ, PYTHONIOENCODING="utf-8", PYTHONUTF8="1",
               PYTHONDONTWRITEBYTECODE="1", CARDO_SKILLS_ROOT=str(SKILL / "Skills"))
    code, text = run_pipeline(cmd, env)

    if code == 0:
        say("[完成] 全链通过。成果：")
        for pat in DELIVERABLES:
            for p in sorted(ROOT.glob(pat)):
                say(f"  - {norm(p.relative_to(ROOT))}")
        return 0

    failed = ""
    try:
        # 总控实际写的键是 steps / status（不是 entries / exit）——原先读错键名，
        # 失败时永远显示「步骤 ?」，续跑建议永远指 01 步。
        log = read_json(ROOT / "过程文件/00-总控/运行日志.json", {}) or {}
        entries = log.get("steps") or log.get("entries") or []
        failed = next((e.get("step") for e in reversed(entries)
                       if e.get("status") not in ("pass", None)), "") or log.get("failedStep") or ""
    except Exception:
        failed = ""
    stop = re.search(r"\[STOP\]\s*\[(\d{2})\]", text)
    if not failed and stop:
        failed = stop.group(1)

    if code == 3 or stop:
        say(f"[待确认] 停在需要人工介入的位置（步骤 {failed or '?'}）。")
        g = GUIDES.get(failed)
        if g:
            say(f"  做什么：{g}")
        say(f"  办完后续跑：{cmd_hint('--from ' + (failed or '01'))}")
        return 3

    say(f"[失败] 流水线退出码 {code}（步骤 {failed or '?'}）。")
    g = GUIDES.get(failed)
    if g:
        say(f"  做什么：{g}")
    report = ROOT / "过程文件/00-总控/运行报告.md"
    if report.is_file():
        say("  运行报告末段：")
        for line in report.read_text(encoding="utf-8", errors="replace").splitlines()[-20:]:
            say("  " + line)
    say(f"  修好后续跑：{cmd_hint('--from ' + (failed or '01'))}（不要跳过失败步骤）")
    return code if code in (1, 2) else 1


if __name__ == "__main__":
    sys.exit(main())
