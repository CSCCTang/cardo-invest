"""技能包完整性自检（doctor）——开工第一步就跑，别等跑到一半才失败。

按 CardoStatus 实战教训固化：技能包最常见的故障不是逻辑错，而是
**文件没拷全 / 依赖没装 / 规则文件缺失 / 便携环境缺包**，
这些在第 8、9 步才暴露，返工成本极高。

检查项：
  1. 各步脚本目录与必需脚本齐全
  2. 各步规则目录与必需规则文件齐全（yaml）
  3. 便携解释器可用 + 关键第三方包可导入
  4. 技能包资产（模板/输入模板）存在
  5. 脚本可编译（语法级）

退出码：0 全部通过；1 有缺失（逐项列出「缺什么、从哪补」）。

用法：
    python <技能目录>/Skills/00-Cardo投资决策总控/scripts/doctor.py
    python <技能目录>/Skills/00-Cardo投资决策总控/scripts/doctor.py --json
"""

from __future__ import annotations

import argparse
import importlib
import json
import os
import sys
from pathlib import Path

SKILLS = Path(os.environ.get("CARDO_SKILLS_ROOT")
              or Path(__file__).resolve().parents[2]).resolve()
PACKAGE = SKILLS.parent
RUNTIME_PY = PACKAGE / "runtime" / "python" / "python.exe"

# 各步：必需脚本 + 必需规则文件（相对该步目录）
REQUIRED = {
    "01-上游取料": {
        "scripts": ["prepare_inputs.py", "collect_inputs.py", "extract_upstream.py"],
        "rules": [],
    },
    "01b-可研财务提取": {
        "scripts": ["build_params.py", "backfill_projects.py", "params_table.py",
                    "params_writer.py", "pdf_tables.py", "extract_feasibility.py"],
        "rules": [],
    },
    "02-匡算": {
        "scripts": ["calc_invest.py"],
        "rules": ["财务口径.yaml"],
    },
    "03-财务评价": {
        "scripts": ["calc_finance.py", "calc_metrics.py"],
        "rules": ["财务口径.yaml"],
    },
    "04-三方案比选": {
        "scripts": ["gen_schemes.py"],
        "rules": ["财务口径.yaml"],
    },
    "05-章节成文": {
        "scripts": ["assemble_chapters.py", "build_write_tasks.py", "check_banned.py",
                    "check_keyq.py", "check_crossnum.py"],
        "rules": ["章节契约.yaml", "写作规范.yaml"],
    },
    "06-渲染docx": {
        "scripts": ["build_invest_docx.py"],
        "rules": ["表格契约.yaml"],
    },
    "07-回读自检": {
        "scripts": ["check_invest.py"],
        "rules": ["章节契约.yaml", "表格契约.yaml"],
    },
    "08-验证与边界": {
        "scripts": [],
        "rules": [],
    },
}

# 渲染基底/院标原件
REQUIRED_ASSETS = [
    PACKAGE / "assets" / "templates" /
    "2 城市更新老旧街区改造项目投资决策建议书（模板）-补充数据来源（包含三类项目）.docx",
    PACKAGE / "assets" / "inputs" / "财务参数-空白.xlsx",
    PACKAGE / "assets" / "inputs" / "槽位-空白.xlsx",
]

# 关键第三方包（便携环境）
REQUIRED_PACKAGES = [
    ("pymupdf", "PDF 解析（可研附表提取）"),
    ("openpyxl", "xlsx 读写（财务参数、现状指标、AHP 权重表）"),
    ("docx", "Word 读写（建议书渲染、现状分析报告解析）"),
    ("yaml", "规则文件读取（章节契约/表格契约/财务口径）"),
    ("PIL", "图片处理"),
    ("numpy", "数值计算（AHP 权重、现金流）"),
]


def check_scripts():
    items = []
    for step, spec in REQUIRED.items():
        base = SKILLS / step
        if not base.is_dir():
            items.append({"ok": False, "kind": "目录", "name": step,
                          "detail": "步骤目录不存在", "fix": "从完整技能包重新拷贝 {}".format(step)})
            continue
        for name in spec["scripts"]:
            path = base / "scripts" / name
            ok = path.is_file()
            items.append({"ok": ok, "kind": "脚本", "name": "{}/scripts/{}".format(step, name),
                          "detail": "" if ok else "缺失",
                          "fix": "" if ok else "从完整技能包补回 {}/scripts/{}".format(step, name)})
        for name in spec["rules"]:
            path = base / "规则" / name
            ok = path.is_file()
            items.append({"ok": ok, "kind": "规则", "name": "{}/规则/{}".format(step, name),
                          "detail": "" if ok else "缺失",
                          "fix": "" if ok else "从完整技能包补回 {}/规则/{}".format(step, name)})
    return items


def check_assets():
    items = []
    for path in REQUIRED_ASSETS:
        ok = path.is_file()
        items.append({"ok": ok, "kind": "资产", "name": path.name,
                      "detail": "" if ok else "缺失",
                      "fix": "" if ok else "从完整技能包补回 assets/"})
    return items


def check_packages():
    items = []
    for name, why in REQUIRED_PACKAGES:
        try:
            importlib.import_module(name)
            items.append({"ok": True, "kind": "依赖", "name": name, "detail": why, "fix": ""})
        except ImportError as error:
            items.append({"ok": False, "kind": "依赖", "name": name,
                          "detail": "无法导入（{}）：{}".format(why, error),
                          "fix": "在便携环境安装：{} -m pip install {}".format(
                              RUNTIME_PY.name, name)})
    return items


def check_compile():
    items = []
    import py_compile
    import tempfile
    target = Path(tempfile.gettempdir()) / "cardo_doctor_compile.pyc"
    for path in sorted(SKILLS.rglob("*.py")):
        if "__pycache__" in str(path):
            continue
        try:
            py_compile.compile(str(path), doraise=True, cfile=str(target))
            items.append({"ok": True, "kind": "编译", "name": str(path.relative_to(SKILLS)),
                          "detail": "", "fix": ""})
        except Exception as error:
            items.append({"ok": False, "kind": "编译", "name": str(path.relative_to(SKILLS)),
                          "detail": str(error)[:120],
                          "fix": "修复语法错误后重跑"})
    return items


def main(argv=None):
    parser = argparse.ArgumentParser(description="CardoInvest 技能包完整性自检")
    parser.add_argument("--json", action="store_true", help="JSON 输出")
    parser.add_argument("--quiet", action="store_true", help="只报问题")
    args = parser.parse_args(argv)

    groups = {
        "脚本与规则": check_scripts(),
        "技能包资产": check_assets(),
        "便携环境依赖": check_packages(),
        "脚本编译": check_compile(),
    }
    total = sum(len(items) for items in groups.values())
    failed = [item for items in groups.values() for item in items if not item["ok"]]

    if args.json:
        print(json.dumps({"total": total, "failed": len(failed),
                          "groups": {k: v for k, v in groups.items()}},
                         ensure_ascii=False, indent=2))
        return 1 if failed else 0

    print("=" * 66)
    print("CardoInvest 技能包完整性自检")
    print("技能包根：{}".format(PACKAGE))
    print("=" * 66)
    for title, items in groups.items():
        bad = [item for item in items if not item["ok"]]
        flag = "OK  " if not bad else "FAIL"
        print("\n[{}] {}（{}/{} 通过）".format(flag, title, len(items) - len(bad), len(items)))
        if not args.quiet or bad:
            for item in (bad if args.quiet else items):
                mark = "✓" if item["ok"] else "✗"
                line = "  {} {}：{}".format(mark, item["kind"], item["name"])
                if item["detail"]:
                    line += "  —— {}".format(item["detail"])
                print(line)
                if item["fix"]:
                    print("      怎么补：{}".format(item["fix"]))
    print("\n" + "=" * 66)
    if failed:
        print("结论：{} / {} 项缺失（上面的「怎么补」逐条处理即可）".format(len(failed), total))
        return 1
    print("结论：全部通过（{} 项）".format(total))
    return 0


if __name__ == "__main__":
    sys.exit(main())
