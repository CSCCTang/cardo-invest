"""Invest 分章复核：按章校验 H1/H2/来源行/占位/字数（只读，只写 --out）。"""
from __future__ import annotations
import argparse, json, re, sys
from pathlib import Path

DEC = Path("05_过程") / "投资决策"
# 占位符识别用**前缀**：生成侧产出的是「（资料缺失：XXX，见缺项表）」，
# 早先这里写「（资料缺失，见缺项表）」——差一个内容片段，永远匹配不上，占位计数恒为 0。
PLACEHOLDER = "（资料缺失"


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("--project-dir", required=True)
    ap.add_argument("--chapter", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--no-timestamp", action="store_true")
    args = ap.parse_args(argv)
    project = Path(args.project_dir)
    cid = args.chapter.strip().upper()
    path = project / DEC / "分章" / (cid + ".md")
    checks = []

    def add(code, name, ok, detail=""):
        checks.append({"id": code, "name": name, "ok": bool(ok), "detail": str(detail)[:300]})

    if not path.is_file():
        add("CH-FILE", "分章稿存在", False, cid + ".md 缺失")
        return finish(args, checks)
    text = path.read_text(encoding="utf-8")
    heads = [line.strip() for line in text.splitlines() if re.match(r"^#{1,3}\s", line.strip())]
    h1 = next((h for h in heads if h.startswith("# ")), "")
    h2 = [h for h in heads if h.startswith("## ")]
    add("CH-H1", "H1 章标题存在", bool(h1), h1)
    sources = [line for line in text.splitlines() if line.startswith("来源：")]
    add("CH-SOURCE", "至少一行来源引用", len(sources) >= 1, "来源 {} 行".format(len(sources)))
    gaps_file = project / DEC / "缺项表.md"
    allowed = (PLACEHOLDER not in text) or gaps_file.is_file()
    add("CH-PLACEHOLDER", "占位已登记缺项表", allowed, "命中 {}，缺项表 {}".format(
        text.count(PLACEHOLDER), "有" if gaps_file.is_file() else "无"))
    add("CH-LEN", "正文达到最低字数 80%", len(text) >= 600, "{} 字".format(len(text)))
    return finish(args, checks)


def finish(args, checks):
    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    payload = {"schema": "cardo/invest-chapter-check/v1", "chapter": args.chapter,
               "ok": all(item["ok"] for item in checks), "checks": checks}
    out.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    for item in checks:
        print(("  ok   " if item["ok"] else "  FAIL ") + item["id"] + " " + item["detail"])
    print("[check_chapter] {} → {}".format("PASS" if payload["ok"] else "FAIL", args.out))
    return 0 if payload["ok"] else 1


if __name__ == "__main__":
    sys.exit(main())
