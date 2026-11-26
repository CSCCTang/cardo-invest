#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""deai_gate.py -- 成品去 AI 味闸门。

回灌自 `04_工具与方法论\01_技能\08_去AI味`：判据固定为
  必删（B1—B12）清零；四项指标与总分只作报告，不单独阻断。
本脚本把同一套扫描器（同目录 deai_scan.py）用在成品 docx 上，
输出 json + md 两份留痕，退出码 0 = 必删清零，1 = 仍有必删命中。

用法：
  python deai_gate.py --docx <成品.docx> [--out-json x.json] [--out-md x.md]
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import deai_scan  # noqa: E402


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="成品去 AI 味闸门（必删清零）")
    ap.add_argument("--docx", required=True)
    ap.add_argument("--out-json", default="")
    ap.add_argument("--out-md", default="")
    ap.add_argument("--genre", choices=["doc", "deck"], default="doc")
    args = ap.parse_args(argv)

    path = Path(args.docx)
    if not path.is_file():
        print("[去AI味] 找不到成品：{}".format(path))
        return 2
    units = deai_scan.chunk_units(deai_scan.load_file(path))
    res = deai_scan.scan(units, args.genre, None)
    res["file"] = str(path)
    if args.out_json:
        out = Path(args.out_json)
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text(json.dumps(res, ensure_ascii=False, indent=2), encoding="utf-8")
    if args.out_md:
        out = Path(args.out_md)
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text(deai_scan.render_md(res), encoding="utf-8")
    print("[去AI味] {} 分 {}｜必删 {}｜慎用 {}｜具体段 {:.0%}".format(
        res["total"], res["grade"], res["bidao_count"], res["shen_count"],
        res["concrete"]["concrete_rate"]))
    return 0 if res["bidao_count"] == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
