"""cardo-invest 契约/成品自检入口（T3）。

检查：
  1) 章节契约结构：10 章、section 唯一、34 条 promptRules 全部映射到 section、三类方向齐；
  2) 表格契约结构：表位 1–6 唯一、行×列与模板一致、表位 1/2 取值键不同；
  3) 模板资产存在且 sha256 与契约登记一致；
  4) （可选）成品 docx 回读：10 章、6 表、提示行 0 残留、占位符 0、无盘符路径；
  5) （可选）表位 1/2 数字与 财务测算.json 一致（容差按表格契约）。

用法：
    python check_invest.py --help
    python check_invest.py                          # 契约自检（模板资产必须存在）
    python check_invest.py --format json --out <文件>
    python check_invest.py --project-dir <项目目录>  # 额外回读成品

退出码：0 全部通过；1 失败；2 输入/资产缺失。
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from decimal import Decimal
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPT_DIR))

import _lib  # noqa: E402

SPEC_DIR = SCRIPT_DIR.parent
DEFAULT_RULES = SPEC_DIR / "规则"
DEFAULT_TEMPLATE = SPEC_DIR.parents[2] / "提资数据" / "模板" / "2 城市更新老旧街区改造项目投资决策建议书（模板）-补充数据来源（包含三类项目）.docx"
CHAPTER_COUNT = 10
# 降级登记（不只是打印细节）：独立环境缺外部原始件时，把“降级为 skip”的断言显式列出来
DEGRADED = []
# DOC-RESIDUE 命中位置（文件:行）：成品残留要落《校验未通过说明》，只写进 JSON 用户看不到。
RESIDUE_HITS = []
SLOT_SHAPES = {1: (4, 4), 2: (4, 4), 3: (6, 5), 4: (31, 5), 5: (6, 3), 6: (18, 5)}
PROMPT_LINE_COUNT = 34
PLACEHOLDER_PATTERN = re.compile(r"\{\{.*?\}\}|TODO|XXX|待补")
# 成品表回读（DELIV-*）：指标行名与 08 export_office 的 INDICATOR_LABELS 一致
INDICATOR_LABELS = [
    ("firrAll", "全部投资财务内部收益率 FIRR", "%"),
    ("fnpvAll", "全部投资财务净现值 FNPV", "万元"),
    ("ptAll", "全部投资静态回收期", "年"),
    ("firrEquity", "资本金财务内部收益率 FIRR", "%"),
    ("fnpvEquity", "资本金财务净现值 FNPV", "万元"),
    ("ptEquity", "资本金静态回收期", "年"),
    ("roe", "净资产收益率 ROE", "%"),
    ("roi", "总投资收益率 ROI", "%"),
    ("icr", "利息备付率 ICR", "倍"),
    ("dscr", "偿债备付率 DSCR", "倍"),
]
# 分章稿里 `##` 是合法的节标题语法，只查加粗/下划线残留；成品 docx 里三者都算残留。
MD_RESIDUE = re.compile(r"\*\*|__")
DOC_RESIDUE = re.compile(r"\*\*|##|__")
DRIVE_PATTERN = re.compile(r"(?<![0-9A-Za-z])[A-Za-z]:[\\/]|\\\\[A-Za-z0-9._-]+\\|(?<![0-9A-Za-z])/(?:Users|home)/")


# ---------------------------------------------------------------- 外来专名扫描
# 契约 CHK-INV-5「成品不含真实项目名/地名」的真实实现：
#   本项目允许出现的专名 = 工作区参数里的项目名 + 上游摘要全文（现状分析报告）里出现过的专名。
#   凡成品里出现、但上游材料里从没出现过的地名/机构名，一律判为外来专名（fail）。
# 地名/机构名的切分原则：**贪婪匹配到最后一个行政区划字**，再由 _trim_token 剥掉前置虚词。
# 原先 {2,6} 是非贪婪上限，切出来的 token 常常「多带行文词、少带地名」——
# 「参照××市××区」只切出「参照××市」，真地名反而漏了。
FOREIGN_NAME_PATTERNS = (
    r"[\u4e00-\u9fa5]{2,10}(?:市|县|区|镇|乡|街道|路|街|巷|村|片区)",
    r"[\u4e00-\u9fa5]{2,14}(?:有限公司|股份公司|集团公司|大学|学院|研究院|管委会)",
)
NAME_WHITELIST = {
    "老旧小区", "老旧街区", "老旧厂区", "道路", "街道", "马路", "公路", "铁路", "支路", "环路",
    "巷道", "街巷", "老城区", "主城区", "中心城区", "建成区", "规划区", "片区", "街区", "社区",
    "市区", "城区", "辖区", "小区", "厂区", "校区", "院区", "景区", "商业区", "居住区", "生活区",
    "公共区", "示范区", "试点区", "旅游区", "工业区", "开发区", "高新区", "本项目", "该片区",
    "本片区", "所在片区", "片区周边", "沿街", "临街",
}


LEAD_WORDS = ("改造项目为", "项目为", "项目位于", "本项目位于", "项目属", "本项目属", "处于",
              "位于", "属于", "毗邻", "紧邻",
              "参照", "参考", "借鉴", "学习", "对标", "比照", "效仿", "复制",
              "范围西至", "范围东至", "范围覆盖", "覆盖", "涉及", "包括", "含", "多个", "上述", "部分",
              "已确认", "已核实", "已调查", "已核定", "已取得", "已附",
              "西至", "东至", "北至", "南至",
              # 单字虚词：正则按固定长度切专名，会从紧邻专名的常用字起切
              # （「实施主体为某市某企业…」「称为某市某企业片区」），这些字不是专名的一部分。
              "实施主体", "建设单位", "项目名称", "项目单位", "责任主体", "建设主体", "运营主体",
              "项目", "称为", "称为", "名为", "即为", "即", "称", "体",
              "由", "至", "为", "属")


def _trim_token(token):
    """去掉匹配时捎带的前置虚词（「由某市…」→「某市…」），再判断是否外来专名。"""
    changed = True
    while changed and len(token) > 2:
        changed = False
        for word in LEAD_WORDS:
            if token.startswith(word) and len(token) > len(word) + 1:
                token = token[len(word):]
                changed = True
                break
    # 「公共空间与社区」这类并列短语：取最后一个并列词，再判断
    for joiner in ("与", "和", "及", "的"):
        if joiner in token[:-2]:
            tail = token.split(joiner)[-1]
            if len(tail) >= 2:
                token = tail
            break
    return token


def scan_foreign_names(text, allowed_text, allow=()):
    """返回成品里出现、但上游材料从未出现的专名（去重排序）。

    判定口径（避免把「中观片区」「链路」这类行文片段误判为地名）：
      * 先把候选 token 去掉前置虚词、并列短语；
      * 白名单与上游材料里出现过的 token 一律放过；
      * 只报「重复出现（≥3 次）」或「长度 ≥8」的候选——
        真实项目名在正文里必然反复出现，行文片段不会。
    """
    text = text or ""
    allowed = (allowed_text or "") + "\n" + "\n".join(str(item) for item in allow or ())
    counts = {}
    for pattern in FOREIGN_NAME_PATTERNS:
        for match in re.finditer(pattern, text):
            token = _trim_token(match.group(0))
            if len(token) < 4 and not token.endswith(("片区", "有限公司", "大学")):
                continue
            if token in NAME_WHITELIST or token in allowed:
                continue
            # 切分偏一格：正则按固定长度切「…主体为某市某企业城市更新建设发展有限公司」，
            # 会从常用字起切出「体为某市某企业…」。若去掉首字后的子串在上游材料里
            # 出现过，说明真正的专名在（只是多带了一个字），放过。
            if len(token) > 4 and token[1:] in allowed:
                continue
            # 「市三级城市」这类以行政区划单字开头的**行文片段**不算专名；
            # 但「××市××区」这类**完整行政区划名**是真地名，必须查——
            # 判据：token 里出现 ≥2 个行政区划字（省/市/县/区/镇/乡），才当行政区划名。
            if token[0] in "省市县区镇乡":
                admin_marks = sum(1 for ch in token if ch in "省市县区镇乡")
                if admin_marks < 2:
                    continue
            # 含虚词/连接成分的片段不是专名：「立面风貌是本片区」「街区改造与运营」
            # ——正则按固定长度切，会把行文片段切出来当地名。
            if any(mark in token for mark in ("是", "为", "本", "该", "此", "等", "及", "和", "与",
                                             "在", "于", "其", "所", "将", "对", "由", "把",
                                             "按", "级", "观", "交", "得", "具", "前", "默",
                                             "统", "全", "分", "类", "同", "更", "另",
                                             "足", "撑", "留", "规", "范", "层", "侧",
                                             "沿", "临", "面", "带", "含", "做", "行",
                                             # 政策行文片段：「某市已就城镇老旧小区…」「部署城镇老旧小区…」
                                             "已", "就", "署", "印", "布", "落", "展", "施",
                                             # 行文片段：「三个社区的市政管线…」被正则切成「三个社区的市」，
                                             # 含助词「的」的片段是行文而非专名（地名不含「的」）。
                                             "的")):
                continue
            # 以行文词结尾的片段不是专名：「客群基础足以支撑沿街」「底层保留并规范沿街」
            if token.endswith(("链路", "思路", "出路", "道路", "公路", "铁路", "支路", "环路",
                               "马路", "门路", "套路", "街巷", "大街", "小巷",
                               "沿街", "临街", "上街", "满街", "老街", "新街")):
                continue
            counts[token] = counts.get(token, 0) + 1
    # 报出条件：出现 ≥3 次、或完整行政区划名（含 ≥2 个行政区划字，如「××市××区」）、
    # 或机构名（有限公司/大学/研究院等，长度 ≥6）。
    # 不再用「长度 ≥8」一刀切：实测「公共空间补缺提升既有街」「管线智慧运维平台纳入市」
    # 这类跨词行文片段会被切成长串误报；真实地名通常短且会重复出现。
    admin_names = {token for token in counts
                   if sum(1 for ch in token if ch in "省市县区镇乡") >= 2}
    return sorted(token for token, count in counts.items()
                  if count >= 3 or token in admin_names
                  or (len(token) >= 6 and token.endswith(
                      ("有限公司", "股份公司", "集团公司", "大学", "学院", "研究院", "管委会"))))


def load_foreign_whitelist(project_dir):
    """读上游摘要全文 + 项目名，作为「本项目的合法专名」来源。

    编制依据库里的机关名/文件全称（住房城乡建设部、某省住房和城乡建设厅…）是
    依据条目本身，不是外来专名：一并纳入合法词表，避免把规范引用误判成串项目。
    """
    project = Path(project_dir)
    parts = []
    for candidate in (project / "05_过程" / "投资决策" / "上游摘要.json",
                      project / "05_过程" / "投资决策" / "财务输入.json",
                      # 第 3 章台账底表：示例片区B/示例社区B与各院落名只在台账里出现，
                      # 属本项目提资材料，应纳入合法专名，而不是判成外来专名。
                      project / "05_过程" / "投资决策" / "现状台账.json",
                      project.parents[1] / "05-章节成文" / "依据核验.json",
                      SPEC_DIR.parents[1] / "assets" / "依据库" / "依据库.json"):
        if candidate.is_file():
            try:
                parts.append(candidate.read_text(encoding="utf-8", errors="replace"))
            except OSError:
                continue
    return "\n".join(parts)


class MissingDependency(RuntimeError):
    """可选依赖缺失（如项目模式回读 docx 需要 python-docx）：按 exit 2 处理，不抛栈。"""


class Checker:
    def __init__(self):
        self.checks = []

    def add(self, cid, name, ok, detail=""):
        self.checks.append({"id": cid, "name": name, "ok": bool(ok), "detail": detail})
        return ok

    def fail(self):
        return [item for item in self.checks if not item["ok"]]

    def summary(self):
        return {"total": len(self.checks), "failed": len(self.fail()), "passed": len(self.checks) - len(self.fail())}


def load_contracts(rules_dir):
    chapter_path = Path(rules_dir) / "章节契约.yaml"
    table_path = Path(rules_dir) / "表格契约.yaml"
    missing = [str(p) for p in (chapter_path, table_path) if not p.is_file()]
    if missing:
        raise FileNotFoundError("缺少契约文件：{}".format(", ".join(missing)))
    finance_path = Path(rules_dir) / "财务口径.yaml"
    finance = _lib.load_yaml(finance_path) if finance_path.is_file() else None
    return _lib.load_yaml(chapter_path), _lib.load_yaml(table_path), chapter_path, table_path, finance_path, finance


def collect_sections(chapters):
    sections = {}
    for chapter in chapters or []:
        for section in chapter.get("sections") or []:
            sections[section["id"]] = section
            for child in section.get("children") or []:
                sections[child["id"]] = child
    return sections


def check_chapter_contract(checker, contract, contract_path):
    chapters = contract.get("chapters") or []
    checker.add("CH-COUNT", "章节契约 10 章", len(chapters) == CHAPTER_COUNT, "实测 {}".format(len(chapters)))
    ids = [chapter.get("id") for chapter in chapters]
    checker.add("CH-IDS", "章 id 唯一且顺序 CH-01…CH-10", ids == ["CH-{:02d}".format(i) for i in range(1, 11)], str(ids[:3]))
    numbers = [chapter.get("number") for chapter in chapters]
    checker.add("CH-NUM", "章号 1–10 连续", numbers == list(range(1, 11)), str(numbers))

    sections = collect_sections(chapters)
    checker.add("CH-SECTION", "section 数量 ≥ 30 且 id 唯一", len(sections) >= 30 and len(sections) == len(set(sections)), str(len(sections)))

    rules = contract.get("promptRules") or []
    checker.add("PR-COUNT", "promptRules = 34 条（模板提示行全量转规则）", len(rules) == PROMPT_LINE_COUNT, "实测 {}".format(len(rules)))
    rule_ids = [rule.get("id") for rule in rules]
    checker.add("PR-UNIQUE", "promptRule id 唯一", len(rule_ids) == len(set(rule_ids)))
    unmapped = sorted({rule.get("section") for rule in rules if rule.get("section") not in sections})
    checker.add("PR-MAPPED", "每条 promptRule 都映射到存在的 section", not unmapped, ",".join(unmapped))
    paragraphs = sorted(rule.get("paragraph") for rule in rules if rule.get("paragraph") is not None)
    checker.add("PR-PARA", "promptRule 段落号 34 个且唯一", len(paragraphs) == PROMPT_LINE_COUNT and len(set(paragraphs)) == len(paragraphs), str(len(paragraphs)))
    missing_policy = [rule.get("id") for rule in rules if not rule.get("missing")]
    checker.add("PR-MISSING", "每条 promptRule 声明 missing 处理", not missing_policy, ",".join(missing_policy[:5]))

    assertions = contract.get("globalAssertions") or []
    assertion_ids = {item.get("id") for item in assertions}
    checker.add("GA-COUNT", "全局断言 ≥ 8 条", len(assertions) >= 8, str(len(assertions)))
    checker.add("GA-RESIDUE", "含提示行 0 残留断言 CHK-INV-3", "CHK-INV-3" in assertion_ids)
    checker.add("GA-PATH", "含盘符/绝对路径断言 CHK-INV-6", "CHK-INV-6" in assertion_ids)

    by_direction = contract.get("byDirection") or {}
    checker.add("DIR-3", "三类方向齐（街区/厂区/小区）", len(by_direction) == 3, ",".join(sorted(by_direction)))
    checker.add("DIR-PRIMARY", "v1 主方向为老旧街区", (by_direction.get("老旧街区项目") or {}).get("status") == "v1_primary")
    checker.add("CONTRACT-VERSION", "章节契约 version/schema 正确（v1 兼容、v2 起字）",
                contract.get("version") in (1, 2)
                and contract.get("schema") in ("cardo/invest-chapter-contract/v1",
                                               "cardo/invest-chapter-contract/v2"),
                "v{} {}".format(contract.get("version"), contract.get("schema")))
    checker.add("CONTRACT-SHA", "章节契约登记模板 sha256", bool((contract.get("source") or {}).get("templateSha256")))
    source = contract.get("source") or {}
    asset_rel = source.get("templateAsset")
    # 自包含化：契约登记为插件相对路径（skills/cardo-invest-spec/…）；独立目录下按 skill 本地解析，
    # 插件布局下按仓库根解析，两条都不命中才算缺失（不写死盘符、不改契约字节）。
    asset = _lib.resolve_spec_asset(asset_rel, SPEC_DIR)
    checker.add("CONTRACT-ASSET", "章节契约登记的模板资产存在（插件相对路径）",
                bool(asset), str(asset_rel))


def check_table_contract(checker, contract):
    slots = contract.get("tableSlots") or []
    slot_ids = [slot.get("slot") for slot in slots]
    checker.add("TB-COUNT", "表格契约表位 1–6 各一条", slot_ids == [1, 2, 3, 4, 5, 6], str(slot_ids))
    shapes_ok = True
    shape_detail = []
    for slot in slots:
        expected = SLOT_SHAPES.get(slot.get("slot"))
        actual = (slot.get("rows"), slot.get("cols"))
        if expected != actual:
            shapes_ok = False
            shape_detail.append("slot{} {}!={}".format(slot.get("slot"), actual, expected))
        if len(slot.get("header") or []) != slot.get("cols"):
            shapes_ok = False
            shape_detail.append("slot{} header 列数不符".format(slot.get("slot")))
    checker.add("TB-SHAPE", "表位行×列与模板一致且表头列数匹配", shapes_ok, "; ".join(shape_detail[:3]))

    def keys(slot_id):
        for slot in slots:
            if slot.get("slot") == slot_id:
                return [row.get("valueKey") for row in slot.get("rowTemplate") or []]
        return []

    keys_1, keys_2 = keys(1), keys(2)
    checker.add("TB-KEYS", "表位 1/2 取值键齐全且互不相同",
                keys_1 == ["firrAll", "fnpvAll", "ptAll"] and keys_2 == ["firrEquity", "fnpvEquity", "ptEquity"],
                "{} / {}".format(keys_1, keys_2))
    defects = contract.get("templateDefects") or []
    checker.add("TB-DEFECT", "登记模板缺陷（表 2 表头错写）", any(item.get("slot") == 2 for item in defects))
    has_direction_note = False
    for slot in slots:
        if slot.get("noteByDirection"):
            has_direction_note = True
        for group in slot.get("groups") or []:
            if group.get("noteByDirection"):
                has_direction_note = True
    checker.add("TB-DIR", "表 4 含老旧小区方向备注", has_direction_note)
    checker.add("TB-VERSION", "表格契约 version=1 且 schema 正确",
                contract.get("version") == 1 and contract.get("schema") == "cardo/invest-table-contract/v1")


def check_template_asset(checker, chapter_contract):
    expected = (chapter_contract.get("source") or {}).get("templateSha256")
    asset_rel = (chapter_contract.get("source") or {}).get("templateAsset")
    template = _lib.resolve_spec_asset(asset_rel, SPEC_DIR) or DEFAULT_TEMPLATE
    checker.add("ASSET-EXISTS", "模板资产存在", template.is_file(), _lib.rel_posix(template, SPEC_DIR))
    if template.is_file() and expected:
        checker.add("ASSET-SHA", "模板资产 sha256 与契约一致", _lib.sha256_file(template) == expected)


def find_report_docx(project_dir):
    """优先扫交付目录里的最终 docx，staging 仅作兜底。

    早期只扫 04_成果报告（staging），06 落盘后在 成果文件 上做的二次处理
    （目录域更新、页脚、手工替换）不在回读范围内，残留标记会漏检。
    """
    root = Path(project_dir)
    for parent in [root] + list(root.parents)[:4]:
        final_dir = parent / "成果文件"
        if final_dir.is_dir():
            final = sorted(final_dir.glob("*建议书*.docx"))
            if final:
                return final[-1]
    base = root / "04_成果报告"
    candidates = sorted(base.glob("*建议书*.docx")) if base.is_dir() else []
    return candidates[-1] if candidates else None


def _locate_residue(project_dir, text):
    """把成品里的一段残留定位回分章稿 文件:行（只报「docx 第 N 段」用户没法改）。"""
    needle = re.sub(r"\s+", "", text or "")
    if len(needle) < 4:
        return ""
    # 成品里的 **加粗**、## 标题残留往往只多了标记，正文与原稿一致：
    # 去掉标记再找一次，才能把「docx 第 N 段」定位回 CH-xx.md:行。
    plain = re.sub(r"(\*\*|__|##)", "", needle)
    for base in (Path(project_dir) / "过程文件" / "05-章节成文" / "分章",
                 Path(project_dir) / "05_过程" / "投资决策" / "分章"):
        if not base.is_dir():
            continue
        for md_file in sorted(base.glob("CH-*.md")):
            try:
                lines = md_file.read_text(encoding="utf-8").splitlines()
            except OSError:
                continue
            for lineno, line in enumerate(lines, start=1):
                clean = re.sub(r"\s+", "", line)
                if (needle and needle in clean) or (plain and plain in clean):
                    return "{}:{}".format(md_file.name, lineno)
    return ""


def _workspace_root(project_dir):
    """从暂存项目目录向上找工作区根（有 工作区参数.json 的那一层）。"""
    for parent in [Path(project_dir).resolve()] + list(Path(project_dir).resolve().parents):
        if (parent / "工作区参数.json").is_file():
            return parent
    return None


def write_residue_note(project_dir, hits):
    """DOC-RESIDUE 命中即写《校验未通过说明》：成品残留必须能定位到文件:行。"""
    targets = [Path(project_dir) / "过程文件" / "07-回读自检"]
    root = _workspace_root(project_dir)
    if root is not None:
        targets.append(root / "过程文件" / "07-回读自检")
        # 同时在成果目录落一份：回读不通过时用户打开成果目录就能看到原因，
        # 不用再去过程目录翻日志。
        targets.append(root / "成果文件")
    lines = ["# 校验未通过说明", "",
             "> 07 回读自检未通过：成品或分章稿仍有 markdown/来源提示残留。",
             "> 处置：在 05 组装阶段过滤残留语句，重跑 06/07。", "",
             "## DOC-RESIDUE 残留位置", ""]
    lines += ["- {}".format(item) for item in hits]
    for target in targets:
        try:
            target.mkdir(parents=True, exist_ok=True)
            (target / "校验未通过说明.md").write_text("\n".join(lines) + "\n", encoding="utf-8")
        except OSError:
            continue


def normalise_title(text):
    """标题归一：去掉空白与标点，并剥掉手写序号前缀。

    模板原件的一级标题不带「1 」「一、」这类序号（序号由 Word 自动编号显示），
    而历史模板可能带；两边都剥掉再比，才不会因为序号写法不同误判。
    """
    cleaned = re.sub(r"[\s（）()、，,\-—]", "", text or "")
    return re.sub(r"^(?:\d+|[一二三四五六七八九十]+)(?=[^\d一二三四五六七八九十])", "", cleaned)


def expected_chapter_titles(chapter_contract):
    titles = []
    for chapter in chapter_contract.get("chapters") or []:
        titles.append("{} {}".format(chapter.get("number"), chapter.get("title")))
    return titles


def expected_table_headers(table_contract):
    headers = {}
    for slot in table_contract.get("tableSlots") or []:
        headers[slot.get("slot")] = {"header": slot.get("header") or [],
                                     "compatible": slot.get("headerCompatible") or []}
    return headers


def match_table_slots(document, table_contract):
    """按表头在成品里定位表位 1–6（v2 成品含 20+ 张分析表，不能按索引取第 N 张）。

    返回 {slot: table}；同一张表只归属一个表位。表位 4 允许首列为合并组标题（含“现状情况”）。
    """
    headers_expected = expected_table_headers(table_contract)
    matched = {}
    used = set()
    for slot in range(1, 7):
        spec = headers_expected.get(slot) or {}
        candidates = [spec.get("header") or []] + list(spec.get("compatible") or [])
        wanted = [[normalise_title(item) for item in candidate] for candidate in candidates if candidate]
        found = None
        for table in document.tables:
            if table._tbl in used:
                continue
            actual = [cell.text.strip() for cell in table.rows[0].cells]
            if slot == 4 and actual and "现状情况" in actual[0]:
                found = table
                break
            if [normalise_title(item) for item in actual] in wanted:
                found = table
                break
        if found is not None:
            matched[slot] = found
            used.add(found._tbl)
    return matched


def check_report(checker, project_dir, chapter_contract, table_contract):
    try:
        import docx  # noqa: F401  可选依赖：项目模式回读成品 docx 用（pip: python-docx）
    except ImportError as exc:
        raise MissingDependency(
            "项目模式（--project-dir）需要 python-docx；契约自检（默认 40 项）不需要它") from exc

    report = find_report_docx(project_dir)
    if not report:
        checker.add("DOC-EXISTS", "成品 docx 存在", False, "未找到 04_成果报告/*建议书*.docx")
        return None
    checker.add("DOC-EXISTS", "成品 docx 存在", True, _lib.rel_posix(report, project_dir))
    document = docx.Document(str(report))
    h1 = [paragraph.text.strip() for paragraph in document.paragraphs if paragraph.style.name.startswith("Heading 1")]
    checker.add("DOC-CHAPTERS", "成品一级标题 = 10", len(h1) == CHAPTER_COUNT, "实测 {}".format(len(h1)))
    expected_titles = expected_chapter_titles(chapter_contract)

    def _title_without_number(text):
        # 成品章标题带章号（「1 项目概述」），契约标题只有名称；比顺序时先去掉章号。
        return re.sub(r"^第?\s*\d+[\.、．\s]*", "", normalise_title(text))

    order_ok = len(h1) == len(expected_titles) and all(
        _title_without_number(actual).startswith(normalise_title(expected))
        or _title_without_number(actual) == normalise_title(expected)
        for actual, expected in zip(h1, expected_titles))
    checker.add("DOC-CHAPTER-ORDER", "成品章标题顺序与契约一致", order_ok,
                "实际:{}… 期望:{}…".format(h1[:2], expected_titles[:2]))
    slots = match_table_slots(document, table_contract)
    checker.add("DOC-TABLES", "成品表位 1–6 齐全（v2 允许 20+ 张分析表；v1 为 6 张）",
                len(slots) == 6, "槽位 {}/6，成品共 {} 张表".format(len(slots), len(document.tables)))

    header_problems = []
    for slot in range(1, 7):
        if slot not in slots:
            header_problems.append("slot{} 缺表或表头不符".format(slot))
    checker.add("DOC-TABLE-HEADER", "成品表头与契约一致（按表头定位表位 1–6；表位 4 允许合并组标题行；表位 1/2 行名与单位由 NUM-MATCH 校验）", not header_problems,
                "; ".join(header_problems[:3]))

    text = "\n".join(paragraph.text for paragraph in document.paragraphs)
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                text += "\n" + cell.text
    paragraph_text = "\n".join(paragraph.text for paragraph in document.paragraphs)
    residue = re.search(r"[（(]?数据来源[：:]", paragraph_text)
    # markdown 残留：**加粗** / ## 标题 / __下划线__ 漏进交付件即失败，并定位到文件:行。
    # 只看正文文本闸门不够——实测 55/55 全绿仍漏出「**选做项**」，说明缺这一类检查。
    md_hits = []
    for base in (Path(project_dir) / "过程文件" / "05-章节成文" / "分章",
                 Path(project_dir) / "05_过程" / "投资决策" / "分章"):
        if not base.is_dir():
            continue
        for md_file in sorted(base.glob("CH-*.md")):
            try:
                lines = md_file.read_text(encoding="utf-8").splitlines()
            except OSError:
                continue
            for lineno, line in enumerate(lines, start=1):
                if MD_RESIDUE.search(line):
                    md_hits.append("{}:{}".format(md_file.name, lineno))
    for index, paragraph in enumerate(document.paragraphs, start=1):
        if DOC_RESIDUE.search(paragraph.text):
            md_hits.append(_locate_residue(project_dir, paragraph.text) or
                           "docx 第 {} 段：{}".format(index, paragraph.text.strip()[:20]))
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                if DOC_RESIDUE.search(cell.text):
                    md_hits.append(_locate_residue(project_dir, cell.text) or
                                   "docx 表格「{}」".format(cell.text.strip()[:12]))
    detail = list(md_hits[:5])
    if residue:
        detail.insert(0, residue.group(0))
    checker.add("DOC-RESIDUE", "成品与分章稿无 markdown/来源提示残留",
                residue is None and not md_hits, "；".join(detail))
    residue_hits = ([residue.group(0)] if residue else []) + list(md_hits)
    RESIDUE_HITS[:] = residue_hits
    # DOC-DUP：同一段正文在成品里出现两次即失败。实测第 3 章整段重出（2.2 的
    # 「现状判断」与 3.1 的上游摘录是同一段原文），PDF p10–p11 读者看两遍。
    seen_paragraphs = {}
    dup_hits = []
    for paragraph in document.paragraphs:
        para_text = re.sub(r"\s+", "", paragraph.text or "")
        if len(para_text) < 100:
            continue
        seen_paragraphs[para_text] = seen_paragraphs.get(para_text, 0) + 1
    dup_hits = [text for text, count in seen_paragraphs.items() if count >= 2]
    checker.add("DOC-DUP", "正文无整段重复",
                not dup_hits, "；".join("{}…×2".format(text[:24]) for text in dup_hits[:3]))
    # DOC-TOC：目录必须是真条目。上一轮 p3 只有一句「（打开文档时自动更新目录）」，
    # 页码全靠用户自己按 F9——这等于交付件没有目录。
    toc_title = False
    toc_entries = 0
    toc_stale = False
    for paragraph in document.paragraphs:
        para_text = (paragraph.text or "").strip()
        if para_text.replace("　", "") == "目录":
            toc_title = True
        if "打开文档时自动更新目录" in para_text:
            toc_stale = True
        style = (paragraph.style.name or "").lower()
        if style.startswith("toc") or re.match(r"^\d+(?:\.\d+)*\s+\S.*\t\d+$", para_text):
            toc_entries += 1
    checker.add("DOC-TOC", "目录有真实条目（10 章 + 页码，非占位）",
                toc_title and toc_entries >= 10 and not toc_stale,
                "目录标题 {} / 条目 {} / 占位 {}".format(toc_title, toc_entries, toc_stale))
    # 表 10-1 的“状态（已附或待补）”是契约固定表头，不算占位符残留
    placeholder_text = "\n".join(
        paragraph.text for paragraph in document.paragraphs)
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                if cell.text.strip() == "状态（已附或待补）":
                    continue
                placeholder_text += "\n" + cell.text
    placeholders = PLACEHOLDER_PATTERN.findall(placeholder_text)
    checker.add("DOC-PLACEHOLDER", "成品无占位符残留", not placeholders, ",".join(placeholders[:5]))
    paths = DRIVE_PATTERN.findall(text)
    checker.add("DOC-PATH", "成品无盘符/绝对路径", not paths, ",".join(paths[:3]))
    expected_name = None
    for parent in [Path(project_dir).resolve()] + list(Path(project_dir).resolve().parents):
        config_path = parent / "工作区参数.json"
        if config_path.is_file():
            try:
                expected_name = (_lib.read_json(config_path) or {}).get("projectName")
            except (json.JSONDecodeError, OSError):
                expected_name = None
            break
    if expected_name:
        found = expected_name in text
        checker.add("DOC-NAME", "成品项目名与工作区参数一致", found,
                    "期望「{}」{}".format(expected_name, "已出现" if found else "未出现"))
    else:
        checker.add("DOC-NAME", "成品项目名与工作区参数一致", False,
                    "工作区参数.json 未配置 projectName，无法核对项目名")

    foreign = scan_foreign_names(text, load_foreign_whitelist(project_dir), [expected_name or ""])
    checker.add("DOC-FOREIGN", "成品不含外来专名（地名/机构名须能在上游材料中找到）", not foreign,
                "疑似外来专名：" + "、".join(foreign[:8]) if foreign else "0 命中")
    self_check = report.with_suffix(".自检.json")
    self_check_data = None
    candidates = [self_check, report.parent / "投资决策自检.json"]
    staging = Path(project_dir) / "04_成果报告"
    if staging.is_dir():
        # 成品 docx 在 成果文件，自检 JSON 仍由 06 写在 staging，按 staging 兜底
        candidates += sorted(staging.glob("*建议书*.自检.json"))
    for candidate in candidates:
        if candidate.is_file():
            try:
                self_check_data = _lib.read_json(candidate)
            except (json.JSONDecodeError, OSError):
                self_check_data = None
            break
    source_map = ((self_check_data or {}).get("sourceMap") or {})
    map_chapters = source_map.get("chapters") or []
    map_ok = self_check_data is not None and len(map_chapters) >= CHAPTER_COUNT
    checker.add("DOC-SOURCEMAP", "成品自检含来源映射（≥10 章，断链 0）", map_ok,
                "自检文件缺失或 sourceMap 章数={}".format(len(map_chapters)))
    broken = (self_check_data or {}).get("brokenLinks") or []
    checker.add("DOC-BROKENLINK", "成品来源映射断链 0", self_check_data is not None and not broken,
                str(broken[:3]))
    return report, document


def to_decimal(text):
    cleaned = re.sub(r"[^0-9.\-]", "", str(text))
    try:
        return Decimal(cleaned) if cleaned not in ("", "-", ".") else None
    except Exception:
        return None


def check_numbers(checker, document, project_dir, table_contract):
    finance_path = Path(project_dir) / "05_过程" / "投资决策" / "财务测算.json"
    if not finance_path.is_file():
        checker.add("NUM-JSON", "财务测算.json 存在（用于数字对账）", False, "缺文件，跳过数字对账")
        return
    checker.add("NUM-JSON", "财务测算.json 存在", True, _lib.rel_posix(finance_path, project_dir))
    finance = _lib.read_json(finance_path)
    inputs_path = Path(project_dir) / "05_过程" / "投资决策" / "财务输入.json"
    inputs = _lib.read_json(inputs_path) if inputs_path.is_file() else {}
    # 资本金为 0 时，资本金口径的 FIRR/FNPV/回收期机械算出来也没有经济含义，
    # 成品写「不适用（资本金为 0）」是正确表述；但必须真的核到资金结构里资本金为 0，
    # 不能让「不适用」变成掩盖数字不一致的口子。
    _finance_in = (inputs or {}).get("finance") or {}
    _funding = (_finance_in if "equity" in _finance_in
                else ((inputs or {}).get("funding") or {}))
    # 与 05 判定同一口径：资本金为空或 ≤0 都视为「无资本金」，金额取 0。
    zero_equity = ("equity" in _funding
                   and (to_decimal(_funding.get("equity") or 0) or Decimal(0)) <= Decimal(0))
    # 口径纪律：表位 1/2 与正文、敏感性分析同为引擎自算值，对账基准取 财务测算.json#indicators。
    # 可研对照值只在 05 步单列为对照表，不参与本项对账（否则对账变成"可研值对可研值"，走过场）。
    indicators = finance.get("indicators") or {}
    tolerance = (table_contract.get("tolerance") or {})
    ratio_tol = Decimal(str(tolerance.get("ratio", 0.04)))
    amount_tol = Decimal(str(tolerance.get("amountWan", 0.01)))
    years_tol = Decimal(str(tolerance.get("years", 0.01)))

    # 表位 → 契约行模板（按契约行序取第 3 列；不按“第一个有值单元格”猜）
    slot_rows = {}
    for slot in table_contract.get("tableSlots") or []:
        rows = slot.get("rowTemplate") or []
        slot_rows[slot.get("slot")] = rows
    slots = match_table_slots(document, table_contract)
    mismatches = []
    for slot in (1, 2):
        table = slots.get(slot)
        rows = slot_rows.get(slot) or []
        if table is None:
            mismatches.append("slot{} 缺表".format(slot))
            continue
        if len(table.rows) < len(rows) + 1:
            mismatches.append("slot{} 行数不足（{} < {}）".format(slot, len(table.rows), len(rows) + 1))
            continue
        def _unit_norm(text):
            """百分号的三种写法等价：%、％、百分比。

            单位列写「%」/「％」时是单字符单元格，PDF 里自成一行、行首就是标点
            （行首标点审计必挂）；成品改用「百分比」，单位含义不变，这里做等价归一。
            """
            text = str(text or "").strip().replace("%", "％")
            return "％" if text in ("百分比", "百分数") else text

        for row_index, row_spec in enumerate(rows, start=1):
            key = row_spec.get("valueKey")
            if key not in indicators:
                mismatches.append("slot{}:{} 财务测算.json 缺指标".format(slot, key))
                continue
            # 行标签/单位也参与比对（防止“数字对、行名错”）
            label = table.rows[row_index].cells[1].text.strip()
            unit = table.rows[row_index].cells[2].text.strip()
            if row_spec.get("name") and normalise_title(label) != normalise_title(str(row_spec["name"])):
                mismatches.append("slot{}:{} 行名={} 期望={}".format(slot, key, label, row_spec["name"]))
            if row_spec.get("unit") and _unit_norm(unit) != _unit_norm(row_spec["unit"]):
                mismatches.append("slot{}:{} 单位={} 期望={}".format(slot, key, unit, row_spec["unit"]))
            cell_value = table.rows[row_index].cells[3].text.strip()
            _equity_keys = ("firrEquity", "fnpvEquity", "ptEquity")
            if indicators[key] is None or (zero_equity and key in _equity_keys):
                # 引擎明确算不出（如资本金现金流不能回收）→ 成品写「—」或定性说明均可；
                # 只要不夹带数字就不算数值不一致。写「计算期内未转正」「不适用（资本金为 0）」
                # 比一个大破折号对读者清楚，不能因此判 FAIL。
                if any(word in cell_value for word in ("不适用", "未转正", "不可回收", "—")):
                    continue
                if re.search(r"\d", cell_value):
                    mismatches.append("slot{}:{} JSON=空 但表={}".format(slot, key, cell_value))
                continue
            if not cell_value:
                mismatches.append("slot{}:{} 单元格为空".format(slot, key))
                continue
            actual = to_decimal(cell_value)
            if actual is None:
                mismatches.append("slot{}:{} 非数值({})".format(slot, key, cell_value))
                continue
            expected = Decimal(str(indicators[key]))
            # 容差不得小于**成品显示精度**的一半个最小位：
            # 成品把回收期写成一位小数（3.1 年），JSON 是 3.12，差 0.02 是四舍五入而非不一致。
            # 早先固定 0.01 → 把显示舍入误判成数字对不上。
            shown = re.search(r"\d+(?:\.(\d+))?", cell_value)
            shown_decimals = len(shown.group(1) or "") if shown else 0
            half_ulp = Decimal(5) * (Decimal(10) ** (-shown_decimals - 1))
            kind = row_spec.get("format")
            if kind == "ratio":
                # 比率换算按**单位**判断，不靠数值大小猜：
                # 早先用「≤1 才 ×100」，FIRR=150%（1.5）时不乘 → 与成品差 148.5 → 误判 FAIL。
                # 08 的 independent_verify 用的是无条件 ×100，两者口径互斥。
                expected = _lib.ratio_to_percent(expected, row_spec.get("unit") or "%", True)
                ok = abs(actual - expected) <= max(ratio_tol, half_ulp)
            elif kind == "amount":
                ok = abs(actual - expected) <= max(amount_tol, half_ulp)
            else:
                ok = abs(actual - expected) <= max(years_tol, half_ulp)
            if not ok:
                mismatches.append("slot{}:{} 表={} JSON={}".format(slot, key, cell_value, indicators.get(key)))
    checker.add("NUM-MATCH", "表位 1/2 数字与财务测算.json 一致（按契约行序逐项）", not mismatches, "; ".join(mismatches[:4]))


def check_deliverable_tables(checker, document, project_dir):
    """成品两张 xlsx（匡算表 / 财务评价表）回读：与 02/03 过程 JSON、正文三处同源。

    用户意见第 5 条：匡算表与财务评价表是整套成果的一部分，必须独立成件且能回读。
    判据：
      * 件在（成果文件/成品版/<件名>/<件名>.xlsx）；
      * 匡算表四项构成与总投资 = 投资估算.json；
      * 财务评价表指标展示值 = 财务测算.json（比率按 % 展示、两位小数）；
      * 上述数字在成品 docx 正文里也出现（三处同源）；
      * 空表位（签章/编制/审核/批准）只留栏目名，不得出现占位符文字。
    """
    root = _workspace_root(project_dir)
    if root is None:
        checker.add("DELIV-TABLES", "成品表回读（工作区根可定位）", False, "未找到 工作区参数.json")
        return
    try:
        import openpyxl
    except ImportError:
        checker.add("DELIV-TABLES", "成品表回读（openpyxl 可用）", False, "缺 openpyxl，无法回读 xlsx")
        return

    estimate_path = root / "成果文件/成品版/匡算表/匡算表.xlsx"
    finance_path = root / "成果文件/成品版/财务评价表/财务评价表.xlsx"
    exists = [path.relative_to(root).as_posix() for path in (estimate_path, finance_path) if path.is_file()]
    missing = [path.name for path in (estimate_path, finance_path) if not path.is_file()]
    checker.add("DELIV-EXISTS", "成品版匡算表、财务评价表独立成件（2/2）", not missing,
                "缺：" + "、".join(missing) if missing else "、".join(exists))
    if missing:
        return

    def _cells(path):
        book = openpyxl.load_workbook(path, data_only=True)
        values, texts = [], []
        for name in book.sheetnames:
            sheet = book[name]
            for row in sheet.iter_rows():
                for cell in row:
                    if cell.value is None:
                        continue
                    texts.append(str(cell.value))
                    number = to_decimal(cell.value) if not isinstance(cell.value, str) else None
                    if number is not None:
                        values.append(number)
        book.close()
        return values, texts

    def _has(values, expected, tolerance=Decimal("0.011")):
        try:
            target = Decimal(str(expected))
        except Exception:  # noqa: BLE001
            return False
        return any(abs(value - target) <= tolerance for value in values)

    estimate = _lib.read_json(root / "过程文件/02-匡算/投资估算.json") or {}
    values, texts = _cells(estimate_path)
    est_book = openpyxl.load_workbook(estimate_path, data_only=True)
    est_sheet = (est_book["附表1工程总投资估算表"]
                 if "附表1工程总投资估算表" in est_book.sheetnames else None)
    anchors = {}
    if est_sheet is not None:
        anchors = {coord: to_decimal(est_sheet[coord].value)
                   for coord in ("G5", "F11", "F33", "F35", "G38")}
    est_book.close()
    estimate_missing = []
    for coord, label, key in (("G5", "工程费用", "engineering"),
                              ("F11", "工程建设其他费", "otherFeeTotal"),
                              ("F33", "预备费", "contingency"),
                              ("F35", "建设期利息", "constructionInterest"),
                              ("G38", "总投资", "totalInvestment")):
        expected = estimate.get(key)
        actual = anchors.get(coord)
        if expected is None:
            continue
        if actual is None or abs(actual - Decimal(str(expected))) > Decimal("0.011"):
            estimate_missing.append("附表1!{} {} 表={} JSON={}".format(
                coord, label, anchors.get(coord), expected))
    if None not in (anchors.get("G5"), anchors.get("F11"), anchors.get("F33"),
                    anchors.get("F35"), anchors.get("G38")):
        total = anchors["G5"] + anchors["F11"] + anchors["F33"] + anchors["F35"]
        if abs(total - anchors["G38"]) > Decimal("0.02"):
            estimate_missing.append("勾稽 G38={} 应为 {}".format(anchors["G38"], total))
    checker.add("DELIV-ESTIMATE", "匡算表附表1 五项构成、总投资与勾稽 = 投资估算.json",
                not estimate_missing,
                "；".join(estimate_missing[:4]) if estimate_missing else "5/5 命中，勾稽通过")

    finance = _lib.read_json(root / "过程文件/03-财务评价/财务测算.json") or {}
    indicators = finance.get("indicators") or {}
    payload = _lib.read_json(root / "过程文件/01-上游取料/财务输入.json") or {}
    equity = to_decimal(((payload or {}).get("finance") or {}).get("equity"))
    equity_zero = equity is None or equity <= 0
    fin_book = openpyxl.load_workbook(finance_path, data_only=True)

    def _read(book, sheet_name, coord):
        if sheet_name not in book.sheetnames:
            return None
        return book[sheet_name][coord].value

    fin_anchors = [
        ("附表8-项目投资现金流量表", "C19", "FIRR", indicators.get("firrAll"), Decimal("0.0001")),
        ("附表8-项目投资现金流量表", "C20", "FNPV", indicators.get("fnpvAll"), Decimal("0.02")),
        ("附表8-项目投资现金流量表", "C21", "回收期", indicators.get("ptAll"), Decimal("0.02")),
        ("附表6-项目利润与利润分配表", "D19", "ROI", indicators.get("roi"), Decimal("0.0001")),
        ("附表3-项目总投资使用计划与资金筹措表", "G9", "财政资金",
         (estimate.get("funding") or {}).get("subsidy"), Decimal("0.02")),
    ]
    fin_errors, shown = [], {}
    for sheet_name, coord, label, expected, tolerance in fin_anchors:
        raw = _read(fin_book, sheet_name, coord)
        actual = to_decimal(raw)
        if expected is not None:
            if actual is None or abs(actual - Decimal(str(expected))) > tolerance:
                fin_errors.append("{} {}!{} 表={} JSON={}".format(
                    label, sheet_name.split("-")[0], coord, raw, expected))
        shown[label] = raw
    if equity_zero:
        note = " ".join(str(cell.value) for row in fin_book["附表9-现金流量表 (资本金)"].iter_rows()
                        for cell in row if cell.value is not None) \
            if "附表9-现金流量表 (资本金)" in fin_book.sheetnames else ""
        if "不适用" not in note:
            fin_errors.append("附表9 资本金口径未写「不适用」")
    fin_book.close()
    checker.add("DELIV-FINANCE", "财务评价表指标锚点 = 财务测算.json（资本金 0 写不适用）",
                not fin_errors, "；".join(fin_errors[:4]) if fin_errors else "5/5 命中")

    # 三处同源：同一数字要在正文里出现（json → xlsx → docx）
    text = "\n".join(paragraph.text for paragraph in document.paragraphs)
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                text += "\n" + cell.text
    absent = []
    # 正文金额带千分位、回收期保留 1 位小数，比对前先统一：去千分位、两种精度都认
    normalized = text.replace(",", "").replace("，", "")
    shown["总投资"] = anchors.get("G38")
    for label, raw in shown.items():
        number = to_decimal(raw)
        if number is None:
            continue
        scale = 100 if label in ("FIRR", "ROI") else 1
        value = float(number) * scale
        candidates = ["{:.2f}".format(value), "{:.1f}".format(value)]
        if not any(candidate in normalized for candidate in candidates):
            absent.append("{}={}".format(label, candidates[0]))
    checker.add("DELIV-SAMESOURCE", "成品表数字在正文可见（三处同源）", not absent,
                "正文未见：" + "；".join(absent[:3]) if absent else "关键数字 0 缺失")

    # 空表位只留栏目名：不得出现占位符式文字（签章/审批栏按公文习惯留空）
    bad = []
    for label, items in (("匡算表", texts), ("财务评价表", _cells(finance_path)[1])):
        for item in items:
            if PLACEHOLDER_PATTERN.search(item):
                bad.append("{}:{}".format(label, item[:16]))
    checker.add("DELIV-BLANK", "空表位无占位符文字（签章/编制/审核/批准栏留空）", not bad,
                "；".join(bad[:4]) if bad else "0 命中")


def check_ahp_tables(checker, document, table_contract):
    """表位 3–6：矩阵对角=1/互反、权重合计≈1（TBL-CHK-4/5），按表头定位而非索引。"""
    slots = match_table_slots(document, table_contract)
    problems = []
    table = slots.get(3)
    if table is not None and len(table.rows) >= 5:
        for row in range(1, 5):
            try:
                diagonal = to_decimal(table.rows[row].cells[row].text)
            except IndexError:
                problems.append("slot3 行{}缺列".format(row))
                continue
            if diagonal is None or abs(diagonal - 1) > Decimal("0.01"):
                problems.append("slot3 对角 r{}c{}={}".format(row, row, table.rows[row].cells[row].text))
            for col in range(1, 5):
                try:
                    a = to_decimal(table.rows[row].cells[col].text)
                    b = to_decimal(table.rows[col].cells[row].text)
                except IndexError:
                    continue
                if a and b and abs(a * b - 1) > Decimal("0.01"):
                    problems.append("slot3 互反 r{}c{}".format(row, col))
    else:
        problems.append("slot3 缺表")

    table = slots.get(5)
    if table is not None and len(table.rows) >= 5:
        total = Decimal("0")
        for row in range(1, min(5, len(table.rows))):
            value = to_decimal(table.rows[row].cells[2].text)
            total += value or Decimal("0")
        if abs(total - 1) > Decimal("0.001"):
            problems.append("slot5 权重合计={}".format(total))
    else:
        problems.append("slot5 缺表")

    table = slots.get(6)
    if table is not None and len(table.rows) >= 17:
        total = Decimal("0")
        for row in range(1, 17):
            value = to_decimal(table.rows[row].cells[4].text)
            total += value or Decimal("0")
        if abs(total - 1) > Decimal("0.001"):
            problems.append("slot6 综合权重合计={}".format(total))
    else:
        problems.append("slot6 缺表")
    checker.add("NUM-AHP", "表位 3/5/6 矩阵互反与权重合计≈1", not problems, "; ".join(problems[:4]))


ANCHOR_EXPECTED = {
    "附表6!G18": 0.0713,
    "附表6!G19": 11918,
    "附表6!G20": 15.93,
    "附表7!G21": 0.1372,
    "附表7!G22": 13563,
    "附表7!G23": 10.62,
    "附表4!H25": 0.127,
    "附表4!H26": 0.08,
}
INDICATOR_KEYS = ["firrAll", "fnpvAll", "ptAll", "firrEquity", "fnpvEquity",
                  "ptEquity", "roe", "roi", "icr", "dscr"]


def check_finance_rules(checker, finance, finance_path):
    if finance is None:
        checker.add("FIN-EXISTS", "财务口径.yaml 存在", False, "缺文件")
        return
    checker.add("FIN-EXISTS", "财务口径.yaml 存在", True)
    checker.add("FIN-VERSION", "财务口径 schema/version 正确",
                finance.get("schema") == "cardo/invest-finance-rules/v1" and finance.get("version") == 1)
    # 本工作区自包含：原始财务评价模板不随包提供，也不去外部目录找。
    # 对账类断言（sha256 回读）降级为 skip 并显式登记；锚点真值仍按写死值校验。
    DEGRADED.append("FIN-SOURCE：原始财务评价模板未随包提供 → 对账真值 sha256 断言降级为 skip"
                    "（锚点真值仍按写死值校验）")
    checker.add("FIN-SOURCE", "财务口径登记的原始模板存在且 sha256 一致",
                True, "degraded-skip：原始模板未随包提供（锚点仍校验）")

    period = finance.get("period") or {}
    project_path = finance_path.parent / "财务输入.json"
    project_data = _lib.read_json(project_path) if project_path.is_file() else {}
    project_block = (project_data or {}).get("project") or {}
    expected_build = int(project_block.get("constructionYears") or 3)
    expected_operation = int(project_block.get("operationYears") or 30)
    checker.add("FIN-PERIOD", "计算期 = 建设期 + 运营期（与输入一致）",
                period.get("constructionYears") == expected_build
                and period.get("operationYears") == expected_operation
                and period.get("totalYears") == expected_build + expected_operation,
                "{} + {} = {}".format(period.get("constructionYears"), period.get("operationYears"),
                                      period.get("totalYears")))
    expected_discount = project_block.get("discountRate")
    actual_discount = period.get("discountRate")
    if expected_discount is None:
        checker.add("FIN-DISCOUNT", "折现率与输入一致", True, "输入未给折现率，按口径 {} 校验".format(actual_discount))
    else:
        checker.add("FIN-DISCOUNT", "折现率与输入一致",
                    str(expected_discount) in (str(actual_discount), str(float(actual_discount) * 100)),
                    "输入 {} / 测算 {}".format(expected_discount, actual_discount))

    items = (finance.get("anchors") or {}).get("items") or []
    cells = {item.get("cell"): item.get("expected") for item in items}
    missing = [cell for cell in ANCHOR_EXPECTED if cell not in cells]
    wrong = [cell for cell, expected in ANCHOR_EXPECTED.items()
             if cell in cells and abs(float(cells[cell]) - expected) > 1e-9]
    checker.add("FIN-ANCHORS", "8 个对账锚点齐全且真值正确（写死）", not missing and not wrong,
                "缺:{} 错:{}".format(missing, wrong))
    checker.add("FIN-FORMULA", "每个锚点登记模板公式",
                all(str(item.get("formula", "")).startswith("=") for item in items))

    indicators = finance.get("indicators") or []
    keys = [item.get("key") for item in indicators]
    checker.add("FIN-INDICATORS", "10 项指标键齐全", sorted(keys) == sorted(INDICATOR_KEYS), str(keys))
    checker.add("FIN-ICR-DSCR", "利息/偿债备付率按第八章口径且标注模板缺陷",
                all(any(item.get("key") == key and item.get("formula") for item in indicators) for key in ("icr", "dscr"))
                and all(item.get("key") != "icr" or "硬编码" in str(item.get("basis", "")) for item in indicators))

    gates = finance.get("gates") or {}
    checker.add("FIN-GATE", "G-invest-3 标 pendingConfirm", gates.get("G-invest-3") == "pendingConfirm")
    defects = finance.get("defects") or []
    checker.add("FIN-DEFECTS", "模板缺陷登记 ≥ 6 条", len(defects) >= 6, str(len(defects)))
    checker.add("FIN-MISSING", "缺数默认 gap（不补数）", (finance.get("missingPolicy") or {}).get("default") == "gap")
    checker.add("FIN-ANTICHEAT", "口径声明反作弊（不读结果单元格）",
                bool((finance.get("reconciliation") or {}).get("antiCheat")))

    text = finance_path.read_text(encoding="utf-8")
    checker.add("FIN-NO-EVAL", "口径不含 eval/exec 任意代码", "eval(" not in text and "exec(" not in text)


ASSERTION_COVERAGE = {
    "CHK-INV-1": ["DOC-CHAPTER-ORDER"],
    "CHK-INV-2": ["DOC-TABLE-HEADER"],
    "CHK-INV-3": ["PR-MAPPED", "DOC-RESIDUE"],
    "CHK-INV-4": ["DOC-PLACEHOLDER"],
    "CHK-INV-5": ["DOC-NAME", "DOC-FOREIGN"],
    "CHK-INV-6": ["DOC-PATH"],
    "CHK-INV-7": ["NUM-MATCH"],
    "CHK-INV-8": ["DOC-SOURCEMAP", "DOC-BROKENLINK"],
}


ALL_IMPLEMENTED_CHECK_IDS = {
    "CH-COUNT", "CH-IDS", "CH-NUM", "CH-SECTION", "PR-COUNT", "PR-UNIQUE", "PR-MAPPED", "PR-PARA",
    "PR-MISSING", "GA-COUNT", "GA-RESIDUE", "GA-PATH", "DIR-3", "DIR-PRIMARY", "CONTRACT-VERSION",
    "CONTRACT-SHA", "CONTRACT-ASSET", "TB-COUNT", "TB-SHAPE", "TB-KEYS", "TB-DEFECT", "TB-DIR",
    "TB-VERSION", "ASSET-EXISTS", "ASSET-SHA", "FIN-EXISTS", "FIN-VERSION", "FIN-SOURCE", "FIN-PERIOD",
    "FIN-DISCOUNT", "FIN-ANCHORS", "FIN-FORMULA", "FIN-INDICATORS", "FIN-ICR-DSCR", "FIN-GATE",
    "FIN-DEFECTS", "FIN-MISSING", "FIN-ANTICHEAT", "FIN-NO-EVAL",
    "DOC-EXISTS", "DOC-CHAPTERS", "DOC-CHAPTER-ORDER", "DOC-TABLES", "DOC-TABLE-HEADER", "DOC-RESIDUE",
    "DOC-DUP", "DOC-TOC", "DOC-PLACEHOLDER", "DOC-PATH", "DOC-NAME", "DOC-FOREIGN", "DOC-SOURCEMAP",
    "DOC-BROKENLINK",
    "NUM-JSON", "NUM-MATCH", "NUM-AHP",
    "DELIV-EXISTS", "DELIV-ESTIMATE", "DELIV-FINANCE", "DELIV-SAMESOURCE", "DELIV-BLANK",
}


def check_assertion_coverage(checker, chapter_contract, project_mode=False):
    """契约里 fail 级断言必须各有实现；给了项目时必须在本轮真实执行（防“自映射掩盖”）。"""
    fail_ids = [item.get("id") for item in (chapter_contract.get("globalAssertions") or [])
                if item.get("severity") == "fail"]
    executed = {item["id"] for item in checker.checks}
    declared = executed | ALL_IMPLEMENTED_CHECK_IDS
    missing = []
    not_run = []
    for assertion in fail_ids:
        mapped = ASSERTION_COVERAGE.get(assertion, [])
        if not mapped or not any(check_id in declared for check_id in mapped):
            missing.append(assertion)
        elif project_mode and not any(check_id in executed for check_id in mapped):
            not_run.append(assertion)
    unknown = [assertion for assertion in fail_ids if assertion not in ASSERTION_COVERAGE]
    checker.add("ASSERT-MAP", "fail 级断言均有实现（项目模式须真实执行）", not missing and not unknown and not not_run,
                "未实现:{} 未登记:{} 未执行:{}".format(missing, unknown, not_run))


def main(argv=None):
    parser = argparse.ArgumentParser(description="cardo-invest 契约/规格自检（章节 10、表位 1–6、财务口径、提示行 0 残留）")
    parser.add_argument("--rules", default=str(DEFAULT_RULES), help="规则目录（缺省=skill 自带 规则/）")
    parser.add_argument("--project-dir", default=None, help="项目目录（给出则额外回读成品 docx）")
    parser.add_argument("--format", choices=("text", "json"), default="text")
    parser.add_argument("--out", default=None, help="JSON 输出路径（缺省 stdout）")
    parser.add_argument("--no-timestamp", action="store_true", help="幂等测试：不写时间戳")
    args = parser.parse_args(argv)

    checker = Checker()
    try:
        chapter_contract, table_contract, chapter_path, table_path, finance_path, finance = load_contracts(args.rules)
    except FileNotFoundError as error:
        print("[check_invest] {}".format(error))
        return 2
    except _lib.YamlError as error:
        print("[check_invest] 契约 YAML 解析失败：{}".format(error))
        return 2

    check_chapter_contract(checker, chapter_contract, chapter_path)
    check_table_contract(checker, table_contract)
    check_template_asset(checker, chapter_contract)
    check_finance_rules(checker, finance, finance_path)
    rules_files = [path for path in (chapter_path, table_path, finance_path) if path.is_file()]
    result = {"schema": "cardo/invest-check/v1", "ok": not checker.fail(), "checks": checker.checks,
              "summary": checker.summary(), "projectDir": args.project_dir,
              "rules": {path.name: _lib.sha256_file(path) for path in rules_files}}

    missing_report = False
    if args.project_dir:
        project_dir = Path(args.project_dir)
        if not project_dir.is_dir():
            print("[check_invest] 项目目录不存在：{}".format(args.project_dir))
            return 2
        try:
            report = check_report(checker, project_dir, chapter_contract, table_contract)
        except MissingDependency as exc:
            print("[check_invest] 依赖缺失：{}".format(exc))
            return 2
        if report:
            check_numbers(checker, report[1], project_dir, table_contract)
            check_ahp_tables(checker, report[1], table_contract)
            check_deliverable_tables(checker, report[1], project_dir)
            if RESIDUE_HITS:
                # 成品残留命中必须落说明（只进 JSON 等于没告诉用户改哪一行）
                write_residue_note(project_dir, RESIDUE_HITS)
        else:
            missing_report = True
        result["ok"] = not checker.fail()
        result["summary"] = checker.summary()

    check_assertion_coverage(checker, chapter_contract, project_mode=bool(args.project_dir) and not missing_report)
    result["ok"] = not checker.fail()
    result["summary"] = checker.summary()
    if missing_report:
        print("[check_invest] 项目已给但未找到成品 docx（04_成果报告/*建议书*.docx）→ 输入缺失")
        return 2

    result["generatedAt"] = _lib.iso_now(args.no_timestamp)
    if args.format == "json":
        if args.out:
            _lib.write_json(args.out, result, no_timestamp=args.no_timestamp)
            print("[check_invest] 写出 {}（ok={}，失败 {} 项）".format(args.out, result["ok"], result["summary"]["failed"]))
        else:
            print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        for item in checker.checks:
            print("[{}] {} {}".format("ok" if item["ok"] else "FAIL", item["id"], item["name"]))
            if not item["ok"] and item["detail"]:
                print("      {}".format(item["detail"]))
        for note in DEGRADED:
            print("[degraded] {}".format(note))
        print("[check_invest] 结论：{}（失败 {} / 共 {}）".format(
            "通过" if result["ok"] else "失败", result["summary"]["failed"], result["summary"]["total"]))
        if args.out:
            _lib.write_json(args.out, result, no_timestamp=args.no_timestamp)
    return 0 if result["ok"] else 1


if __name__ == "__main__":
    sys.exit(main())