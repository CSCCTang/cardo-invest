"""编制依据库：读取、分类、联网核验与章节文本生成。

纪律：
  * 只读依据库，不改包内基线；运行期复核结果写回工作区 过程文件/05-章节成文/。
  * 正文文号一律逐字取库值；库里写“未标注/待核实/未见公开文号”的，正文不写文号，
    按《运行时联网核验规范.md》作方向性引用。
  * 联网失败不阻断出稿，逐条标“未联网复核”，不把待核实升格为已核实。

用法（被 assemble_chapters.py 调用）：
    from basis_library import build_basis_section
    lines = build_basis_section(project_dir, project_name, params)
"""
from __future__ import annotations

import json
import re
import ssl
import urllib.request
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path

CATEGORY_ORDER = ["法律法规", "政策文件", "技术标准", "编制办法", "项目资料"]
CN_NUM = "一二三四五六七八九十"
NO_NUMBER_RE = re.compile(r"未|待|无文号|未见|暂无")
ISO_DATE_RE = re.compile(r"^(\d{4})-(\d{2})-(\d{2})")


def _read_json(path):
    return json.loads(Path(path).read_text(encoding="utf-8-sig"))


def _find_workspace(project_dir):
    if not project_dir:
        return None
    probe = Path(project_dir).resolve()
    for _ in range(4):
        if (probe / "工作区参数.json").is_file() or (probe / "输入文件").is_dir():
            return probe
        probe = probe.parent
    return None


def _read_params(workspace):
    if workspace and (workspace / "工作区参数.json").is_file():
        try:
            return _read_json(workspace / "工作区参数.json")
        except (OSError, json.JSONDecodeError):
            return {}
    return {}


def _resolve_library(workspace, params):
    package_root = Path(__file__).resolve().parents[3]
    cfg = (params.get("basisLibrary") or {})
    candidates = []
    for raw in (cfg.get("path"), cfg.get("fallback"),
                "assets/依据库/依据库.json"):
        if not raw:
            continue
        path = Path(str(raw))
        candidates.append(path if path.is_absolute() else (workspace or package_root) / path)
    candidates.append(package_root / "assets" / "依据库" / "依据库.json")
    for path in candidates:
        if path.is_file():
            return path
    return None


def classify(entry):
    """把库里的“类型”映射为报告的五类；条例按法律法规呈现，不改库字段。"""
    kind = str(entry.get("类型") or "").strip()
    name = str(entry.get("名称") or "")
    if kind in ("技术标准", "编制办法", "项目资料"):
        return kind
    if kind == "法律法规" or "条例" in name:
        return "法律法规"
    return "政策文件"


def is_citable(entry):
    """有更新/废止的条目不进正文（规范 §4）。待核实条目保留，但只作方向性引用。"""
    return str(entry.get("联网核验结果") or "") not in ("有更新", "已废止")


def has_number(entry):
    return not NO_NUMBER_RE.search(str(entry.get("文号") or ""))


def cn_date(value):
    text = str(value or "").strip()
    if not text:
        return "发布日期未取得"
    match = ISO_DATE_RE.match(text)
    if match:
        return "{}年{}月{}日".format(match.group(1), int(match.group(2)), int(match.group(3)))
    return text


def _fetch(url, timeout):
    request = urllib.request.Request(url, headers={
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                      "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"})
    context = ssl.create_default_context()
    context.check_hostname = False
    context.verify_mode = ssl.CERT_NONE
    with urllib.request.urlopen(request, timeout=timeout, context=context) as response:
        raw = response.read(2_000_000)
    for encoding in ("utf-8", "gb18030"):
        try:
            return raw.decode(encoding)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", "ignore")


def _verify_one(entry, timeout):
    url = str(entry.get("原文URL") or entry.get("核验来源") or "").strip()
    now = datetime.now().astimezone().isoformat(timespec="seconds")
    result = {
        "id": entry.get("id"),
        "名称": entry.get("名称"),
        "文号": entry.get("文号"),
        "检索URL": url,
        "核验时间": now,
        "复核方式": "离线基线",
        "联网核验结果": "未能联网",
        "说明": "",
    }
    if not url:
        result["说明"] = "库中无官方原文链接，按离线基线引用"
        return result
    try:
        html = _fetch(url, timeout)
    except Exception as error:  # noqa: BLE001 - 网络异常一律降级
        result["说明"] = "联网失败：{}".format(type(error).__name__)
        return result
    compact = re.sub(r"\s+", "", html)
    name_hit = re.sub(r"\s+", "", str(entry.get("名称") or ""))[:12] in compact
    number = str(entry.get("文号") or "")
    number_hit = (not has_number(entry)) or re.sub(r"\s+", "", number) in compact
    if name_hit and number_hit:
        result["复核方式"] = "在线原文"
        result["联网核验结果"] = "一致"
    else:
        result["说明"] = "页面可访问但未同时命中名称与文号，按未能联网处理"
    return result


def verify_entries(entries, project_dir, params, online=True):
    cfg = params.get("basisLibrary") or {}
    if not online or cfg.get("onlineVerify") is False:
        now = datetime.now().astimezone().isoformat(timespec="seconds")
        return [{
            "id": entry.get("id"), "名称": entry.get("名称"), "文号": entry.get("文号"),
            "检索URL": str(entry.get("原文URL") or ""), "核验时间": now,
            "复核方式": "离线基线", "联网核验结果": "未能联网",
            "说明": "本次运行关闭联网核验",
        } for entry in entries]
    timeout = float(cfg.get("timeoutSeconds") or 6)
    results = []
    with ThreadPoolExecutor(max_workers=6) as pool:
        futures = {pool.submit(_verify_one, entry, timeout): entry for entry in entries}
        for future in as_completed(futures):
            results.append(future.result())
    order = {str(entry.get("id")): index for index, entry in enumerate(entries)}
    results.sort(key=lambda item: order.get(str(item.get("id")), 999))
    return results


def _write_trace(workspace, library_path, entries, results):
    if not workspace:
        return
    out_dir = workspace / "过程文件" / "05-章节成文"
    out_dir.mkdir(parents=True, exist_ok=True)
    summary = {
        "生成时间": datetime.now().astimezone().isoformat(timespec="seconds"),
        "依据库": str(library_path),
        "条目总数": len(results),
        "在线一致": sum(1 for item in results if item["联网核验结果"] == "一致"),
        "未联网复核": sum(1 for item in results if item["联网核验结果"] == "未能联网"),
        "有更新或废止": sum(1 for item in results if item["联网核验结果"] in ("有更新", "已废止")),
    }
    payload = {"summary": summary, "entries": results}
    (out_dir / "依据核验.json").write_text(
        json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    lines = ["# 编制依据运行期核验留痕", "",
             "- 生成时间：{}".format(summary["生成时间"]),
             "- 依据库：`{}`".format(library_path),
             "- 在线一致 {} 条，未联网复核 {} 条，有更新/废止 {} 条。".format(
                 summary["在线一致"], summary["未联网复核"], summary["有更新或废止"]),
             "", "| 条目 | 名称 | 文号 | 复核方式 | 结果 | 来源 |", "| --- | --- | --- | --- | --- | --- |"]
    for item in results:
        lines.append("| {} | {} | {} | {} | {} | {} |".format(
            item["id"], str(item["名称"]).replace("|", "｜"), str(item["文号"]).replace("|", "｜"),
            item["复核方式"], item["联网核验结果"], item["检索URL"] or "—"))
    (out_dir / "依据核验.md").write_text("\n".join(lines) + "\n", encoding="utf-8")


def load_library(project_dir):
    workspace = _find_workspace(project_dir)
    params = _read_params(workspace)
    path = _resolve_library(workspace, params)
    if path is None:
        return None, workspace, params
    data = _read_json(path)
    entries = data.get("条目") or data.get("entries") or []
    return {"path": path, "entries": entries, "meta": data}, workspace, params


def build_basis_section(project_dir, project_name, params=None):
    """返回 1.4 编制依据的段落列表（含分类序数、文号、发布机关、核验标注）。"""
    library, workspace, loaded_params = load_library(project_dir)
    params = params or loaded_params or {}
    if not library or not library["entries"]:
        return ["本节编制依据库未就绪，待按《依据库.json》补齐后引用。"], {}
    entries = [entry for entry in library["entries"] if is_citable(entry)]
    results = verify_entries(entries, project_dir, params)
    _write_trace(workspace, library["path"], entries, results)
    result_by_id = {str(item["id"]): item for item in results}

    project_entries = [{
        "id": "PRJ-00",
        "类型": "项目资料",
        "名称": "{}提资文件（现状分析报告、财务参数表、AHP 权重表与现状指标）".format(project_name),
        "文号": "无文号（项目资料）",
        "发布机关": "建设单位、项目组",
        "发布日期": "2026年",
        "核验状态": "项目资料",
        "联网核验结果": "项目资料",
    }]
    bucket = {name: [] for name in CATEGORY_ORDER}
    for entry in entries + project_entries:
        bucket[classify(entry)].append(entry)

    lines = ["本报告编制依据按法律法规、政策文件、技术标准、编制办法、项目资料五类列示，"
             "共 {} 项；文号以依据库登记值为准，未取得公开原文的只作方向性引用。".format(
                 sum(len(items) for items in bucket.values()))]
    offline = [str(item.get("id")) for item in results
               if item.get("联网核验结果") not in ("一致", "项目资料")]
    if offline:
        # 断网/核验不通过时只标一次：逐条重复会把正文写成工作底稿（留痕在依据核验.json）
        lines.append("本节依据条目生成时未联网复核，按离线基线引用；核验来源与核验时间随包归档。")
    for index, category in enumerate(CATEGORY_ORDER):
        items = bucket[category]
        if not items:
            continue
        lines.append("（{}）{}".format(CN_NUM[index], category))
        # 国家—省—市—项目，同层级按库内 id 排序
        order = {"国家": 0, "某省": 1, "某市": 2, "项目": 3}
        items.sort(key=lambda entry: (order.get(str(entry.get("层级") or ""), 4),
                                      str(entry.get("id") or "")))
        for number, entry in enumerate(items, 1):
            name = str(entry.get("名称") or "").strip()
            issuer = str(entry.get("发布机关") or "").strip()
            date = cn_date(entry.get("发布日期"))
            # 库内「待核实/未见公开」是登记状态，不是正文该出现的字眼：
            # 拿不到机关/日期就不写这两段，只留文件名＋方向性引用说明（C2 第 3 条）。
            if any(mark in issuer for mark in ("待核实", "未见", "未标注", "未公开", "未知", "空")):
                issuer = ""
            if any(mark in str(entry.get("发布日期") or "")
                   for mark in ("待", "未见", "未标注", "未公开", "未知")):
                date = ""
            if category == "项目资料":
                tail = "，".join(part for part in (issuer, date) if part)
                lines.append("{}.《{}》{}。".format(number, name, "，" + tail if tail else ""))
                continue
            if has_number(entry):
                lead = "{}.《{}》（{}）".format(number, name, entry.get("文号"))
            else:
                lead = "{}.《{}》".format(number, name)
            tail = "，".join(part for part in (issuer, date) if part)
            if tail:
                lead += "，" + tail
            verify = result_by_id.get(str(entry.get("id")), {})
            if verify.get("联网核验结果") in ("有更新", "已废止"):
                continue
            if str(entry.get("核验状态") or "") == "待核实":
                lead += "；未取得公开原文，仅作方向性引用，不引用其具体条款与指标"
            lines.append(lead + "。")
    return lines, {"entries": entries, "results": result_by_id, "library": library}


def cite(project_dir, entry_id):
    """按库内 id 取一条依据的正文引用串（供 2.1 等章节复用，保证文号逐字一致）。"""
    library, _workspace, _params = load_library(project_dir)
    if not library:
        return ""
    for entry in library["entries"]:
        if str(entry.get("id")) == str(entry_id):
            name, number = str(entry.get("名称") or ""), str(entry.get("文号") or "")
            if has_number(entry):
                return "《{}》（{}）".format(name, number)
            return "《{}》".format(name)
    return ""
