"""批量倒推各项目的 财务参数.xlsx（交付物生成器）。

用途：把提资材料里每个项目的可研跑一遍 01b，产出
      <项目目录>/输入文件/财务参数.xlsx
      <项目目录>/输入文件/财务参数-空白.xlsx
这样用户打开任一项目目录，就能直接看到「这个项目的财务参数长什么样」，
也可直接改填。

设计要点：
  * 只在**提资数据的项目目录**里落文件，不污染技能包与测试工作区。
  * 每个项目独立工作区（临时目录），跑完只把 财务参数*.xlsx 拷回项目目录。
  * 单个项目失败不影响其它项目，失败原因写进汇总表。
  * 可研 PDF 不复制（体积大），只引用；财务参数里已记录来源文件名与页码。

用法：
    python Skills/01b-可研财务提取/scripts/backfill_projects.py --source <提资数据根> [--dry-run]
"""

from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from params_table import BLANK_NAME, FILLED_NAME  # noqa: E402

# 项目级可研的判定关键词（文件名里出现即视为项目级可研）
PROJECT_PDF_RULES = (
    "可行性研究报告",
    "可研报告",
    "项目建议书",
    "可行性研究",
)

# 明确排除的文件（不是项目级可研：批复、评审报告、单条市政道路可研等）
# 注意：「（附专家审查意见）」「（附专家意见）」是**附件说明**不是文件性质，
# 不能据此排除 —— 实测多份项目级可研都带这种后缀。
EXCLUDE_KEYS = ("批复", "评审报告", "市政道路工程可行性研究报告")


def _safe_dirname(name):
    """从文件名生成安全的目录名（去掉扩展名与非法字符）。"""
    stem = name[:-4] if name.lower().endswith(".pdf") else name
    for ch in '<>:"/\\|?*':
        stem = stem.replace(ch, "_")
    return stem.strip() or "项目"


def _project_root(landing):
    """项目根目录（输入/ 与 成果/ 的父目录）。"""
    parent = landing.parent
    if parent.name in ("输入", "成果"):
        return parent.parent
    return parent


def _project_input_dir(base, project_name):
    """确定项目的输入目录：<项目目录>/输入/

    规则（自下而上找最近的项目目录）：
      * 从 base 向上逐级找**含「输入」或「成果」子目录**的目录 → 用它的「输入」；
        （实测：某新区的可研埋在「成果-…/1 全套成果文件（…）/」三层深处，
          真正的项目输入文件夹是「老旧小区/1 某新区老旧安置小区/输入/」）
      * 找不到 → 以可研名建项目目录：<base>/<可研简称>/输入/
    """
    current = base
    for _ in range(6):                      # 最多向上找 6 层
        if current == current.parent:
            break
        if (current / "输入").is_dir() or (current / "成果").is_dir():
            return current / "输入"
        current = current.parent
    return base / project_name / "输入"


def _short_name(name):
    """项目简称：去掉序号前缀与文件性质后缀，保留可辨识的主体名。

    实测要处理的形态：
      '2-1.某企业饮品老旧厂区城市更新提升改造项目可行性研究报告' → '某企业饮品老旧厂区城市更新提升改造项目'
      '22  可研报告 某新区老旧安置小区更新工程项目建议书（盖章）(1)' → '某新区老旧安置小区更新工程项目建议书'
      '附件4-1 某片区一区至四区…可行性研究报告' → '某片区一区至四区等六个老旧小区'
      '3-附件-某路（…）市政道路工程项目可行性研究报告' → '某路(A路-B路)市政道路工程'
    """
    stem = name[:-4] if name.lower().endswith(".pdf") else name
    stem = re.sub(r"^\s*\d+(?:[\.\-]\d+)*[\.\-]?\s*", "", stem)   # 去序号前缀 2-1. / 22 / 3-
    stem = re.sub(r"^(附件|附图)\s*\d*(?:[\.\-]\d+)*\s*[\.\-]?\s*", "", stem)   # 去「附件4-1」「附图-」
    stem = re.sub(r"[-—]?附件[-—]?$", "", stem)
    stem = re.sub(r"[-—]?附图[-—]?$", "", stem)
    stem = re.sub(r"[（(]\s*(盖章|盖章全|附专家审查意见|附专家意见|代项目建议书)\s*[）)]", "", stem)
    stem = re.sub(r"(可行性研究报告|可研报告合并|可研报告|项目建议书|代项目建议书)$", "", stem)
    stem = re.sub(r"[（(]\d+[）)]\s*$", "", stem)                  # 去尾部 (1)
    stem = stem.replace("（", "(").replace("）", ")").strip(" ._-")
    # 去掉不成对的括号
    while "(" in stem and ")" not in stem:
        stem = stem.replace("(", "")
    while ")" in stem and "(" not in stem:
        stem = stem.replace(")", "")
    return stem.strip() or "项目"


def find_projects(source_root):
    """扫描提资数据根，返回 [{name, dir, pdfs:[Path]}]。

    粒度：**一份可研 = 一个项目**。

    落位目录（dir）= **项目的输入文件夹**，即 <项目目录>/输入/：
      * 可研所在目录自身就是项目目录（含「输入」/「成果」子目录）→ 用它的「输入」；
      * 否则以可研名建项目目录 → <可研所在目录>/<可研简称>/输入/。
    另：已存在但没有可研的项目（只有「输入」目录）也纳入，
      给它们发**空白模板**，用户可直接填。
    """
    source_root = Path(source_root)
    results = []
    seen = set()
    used_dirs = {}          # landing 目录 → 已用的项目名（防互相覆盖）
    def priority(path):
        """同一项目有多份可研时的取用优先级：正文报告 > 附件 > 附图。"""
        name = path.name
        if "附图" in name:
            return 2
        if "附件" in name:
            return 1
        return 0

    for path in sorted(source_root.rglob("*.pdf"), key=priority):
        name = path.name
        if any(key in name for key in EXCLUDE_KEYS):
            continue
        if not any(rule in name for rule in PROJECT_PDF_RULES):
            continue
        key = str(path)
        if key in seen:
            continue
        seen.add(key)
        landing = _project_input_dir(path.parent, _safe_dirname(_short_name(name)))
        # 同一项目目录下的多份可研（实测某企业片区：「3-附件-某路…可研」与
        # 「4-附图-某路…可研(1)」是同一份可研的附件与附图）→ **只取一份**，
        # 落到同一个 输入/ 目录。绝不造「（2）」这类多余项目目录。
        key = str(landing).lower()
        if key in used_dirs:
            continue
        used_dirs[key] = 1
        results.append({
            "name": name[:-4] if name.lower().endswith(".pdf") else name,
            "dir": landing,
            "pdfs": [path],
            "project_root": _project_root(landing),
        })
    # 已存在但没有可研的项目（也要给空白模板，用户可填）
    for dirpath, dirnames, filenames in os.walk(source_root):
        current = Path(dirpath)
        if current.name in ("输入", "成果"):
            continue
        # 输入/成果 目录的下级不再算项目目录（避免生成「输入/输入」嵌套）
        if any(parent.name in ("输入", "成果") for parent in current.parents
               if parent != source_root.parent):
            continue
        if not (current / "输入").is_dir():
            continue
        if any((current / "输入" / item).is_file() for item in filenames):
            continue
        if str((current / "输入").resolve()).lower() in used_dirs:
            continue          # 已有可研归位，不必再放空白模板
        if any(item["dir"] == current / "输入" for item in results):
            continue
        results.append({
            "name": current.name,
            "dir": current / "输入",
            "pdfs": [],
            "project_root": current,
        })
    return results


def run_one(project, package_root, dry_run=False):
    """对单个项目跑 01b，返回结果摘要。"""
    with tempfile.TemporaryDirectory(prefix="cardo_backfill_") as tmp:
        workspace = Path(tmp)
        input_dir = workspace / "输入文件"
        input_dir.mkdir(parents=True, exist_ok=True)
        # 只放可研本体（不复制大文件，用链接或直接指向原路径）
        for pdf in project["pdfs"]:
            target = input_dir / pdf.name
            try:
                if target.exists():
                    target.unlink()
                os_link = getattr(shutil, "copy2", None)
                del os_link
                target.symlink_to(pdf.resolve())
            except (OSError, NotImplementedError):
                shutil.copy2(pdf, target)
        # 成果目录也带进临时工作区：匡算工作簿在「成果」里，
        # 01b 要扫到它才能把工程费/资金/收入预填进财务参数.xlsx。
        outcome = project.get("project_root")
        if outcome and outcome.is_dir():
            outcome_dir = outcome / "成果"
            if outcome_dir.is_dir():
                target = workspace / "成果"
                target.mkdir(exist_ok=True)
                for item in outcome_dir.iterdir():
                    if item.is_file() and item.suffix.lower() in (".xlsx", ".xls"):
                        try:
                            shutil.copy2(item, target / item.name)
                        except OSError:
                            pass
        # 传递项目目录名（项目名的权威来源）：写在工作区根的 .项目名 文件里
        project_dir_name = None
        root = project.get("project_root")
        if root is not None and root.is_dir():
            project_dir_name = root.name
        elif project.get("dir"):
            project_dir_name = project["dir"].parent.name
        if project_dir_name:
            (workspace / ".项目名").write_text(project_dir_name, encoding="utf-8")
        script = package_root / "Skills" / "01b-可研财务提取" / "scripts" / "build_params.py"
        python = package_root / "runtime" / "python" / "python.exe"
        if not python.is_file():
            python = Path(sys.executable)
        proc = subprocess.run(
            [str(python), str(script), "--workspace", ".", "--no-timestamp"],
            cwd=str(workspace), capture_output=True, text=True, encoding="utf-8", errors="replace")
        output = (proc.stdout or "") + (proc.stderr or "")
        summary = {"project": project["name"], "exit": proc.returncode,
                   "pdfs": [p.name for p in project["pdfs"]],
                   "log": [line for line in output.splitlines() if line.strip()]}
        if dry_run:
            return summary
        # exit 2 = 没有可解析的财务源（项目没交可研）→
        # **照样要把空白模板放进输入文件夹**：用户可以自己填。
        # 这是设计要求：财务参数.xlsx 始终在输入文件夹，给不出也要给模板。
        if proc.returncode != 0:
            blank = input_dir / BLANK_NAME
            if blank.is_file():
                target_dir = project["dir"]
                target_dir.mkdir(parents=True, exist_ok=True)
                shutil.copy2(blank, target_dir / BLANK_NAME)
                summary["output_dir"] = str(target_dir)
                summary["note"] = "无可研，已放空白模板待填"
            return summary
        # 拷回项目目录
        target_dir = project["dir"]
        target_dir.mkdir(parents=True, exist_ok=True)
        for name in (FILLED_NAME, BLANK_NAME):
            src = input_dir / name
            if src.is_file():
                shutil.copy2(src, target_dir / name)
        summary["output_dir"] = str(target_dir)
        return summary


def main():
    parser = argparse.ArgumentParser(description="批量倒推各项目 财务参数.xlsx")
    parser.add_argument("--source", required=True, help="提资数据根目录")
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--no-timestamp", action="store_true")
    args = parser.parse_args()

    package_root = Path(__file__).resolve().parents[3]
    projects = find_projects(args.source)
    print("[backfill] 发现 {} 个项目级可研：".format(len(projects)))
    for item in projects:
        print("   · {} → {}".format(item["name"],
                                   "、".join(p.name for p in item["pdfs"])))

    results = []
    for item in projects:
        print("\n[backfill] 处理：{}".format(item["name"]))
        summary = run_one(item, package_root, dry_run=args.dry_run)
        results.append(summary)
        for line in summary["log"]:
            if line.startswith("[01b]"):
                print("    " + line)
        if summary["exit"] != 0:
            print("    ! 失败（exit={}）".format(summary["exit"]))

    ok = [r for r in results if r["exit"] == 0]
    print("\n[backfill] 完成：{}/{} 项目成功".format(len(ok), len(results)))
    for item in ok:
        print("   ✓ {} → {}".format(item["project"], item.get("output_dir", "(dry-run)")))

    if not args.dry_run:
        report = Path(args.source) / "财务参数倒推汇总.json"
        payload = {
            "generatedAt": None if args.no_timestamp else datetime.now().isoformat(timespec="seconds"),
            "projects": results,
        }
        report.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        print("[backfill] 汇总 → {}".format(report))
    return 0


if __name__ == "__main__":
    sys.exit(main())
