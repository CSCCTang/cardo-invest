"""cardo-invest 共享底座（被各 skill 的脚本 import；单一真值，不复制）。

定位方式（不写盘符/绝对路径）：
    SCRIPT_DIR = Path(__file__).resolve().parents[2]  # -> skills/
    sys.path.insert(0, str(SCRIPT_DIR / "cardo-invest-spec" / "scripts"))
    import _lib

提供：
  - YAML 子集解析（parse_yaml/load_yaml；支持注释、嵌套映射、序列、内联标量列表、引号）
  - sha256 / 文件遍历 / 相对路径（posix）
  - Decimal 金额与比率取整（q2/q4/dec）
  - JSON 信封（envelope）与幂等写盘（no_timestamp）
  - ISO 时间
"""
from __future__ import annotations

import hashlib
import json
import os
from datetime import datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path

# ---------------------------------------------------------------- YAML 子集


class YamlError(ValueError):
    """YAML 子集解析错误（行号 + 原因）。"""


def _strip_comment(line: str) -> str:
    """去掉行内注释（# 前必须有空白或行首；引号内的 # 保留）。"""
    quote = None
    for index, char in enumerate(line):
        if quote:
            if char == quote:
                quote = None
            continue
        if char in ("'", '"'):
            quote = char
            continue
        if char == "#" and (index == 0 or line[index - 1] in " \t"):
            return line[:index]
    return line


def _parse_scalar(text: str):
    text = text.strip()
    if text == "":
        return None
    if text.startswith("[") or text.endswith("]"):
        if not (text.startswith("[") and text.endswith("]")):
            raise YamlError("内联列表情括号不配对：{}".format(text[:60]))
        inner = text[1:-1].strip()
        if not inner:
            return []
        return [_parse_scalar(part) for part in _split_inline(inner)]
    if text.startswith(("'", '"')) or text.endswith(("'", '"')):
        if len(text) < 2 or text[0] != text[-1]:
            raise YamlError("引号不配对：{}".format(text[:60]))
        return text[1:-1]
    lowered = text.lower()
    if lowered in ("true", "yes"):
        return True
    if lowered in ("false", "no"):
        return False
    if lowered in ("null", "~"):
        return None
    try:
        if text.isdigit() or (text.startswith("-") and text[1:].isdigit()):
            return int(text)
        return float(text) if ("." in text and text.replace(".", "", 1).lstrip("-").isdigit()) else text
    except ValueError:
        return text


def _split_inline(text: str):
    parts = []
    quote = None
    current = []
    for char in text:
        if quote:
            current.append(char)
            if char == quote:
                quote = None
            continue
        if char in ("'", '"'):
            quote = char
            current.append(char)
        elif char == ",":
            parts.append("".join(current).strip())
            current = []
        else:
            current.append(char)
    if current:
        parts.append("".join(current).strip())
    return parts


def _split_key(line: str):
    """把 'key: value' 拆成 (key, value)；key 内不允许冒号。"""
    if ":" not in line:
        raise YamlError("缺少冒号：{}".format(line))
    key, value = line.split(":", 1)
    return key.strip(), value.strip()


def parse_yaml(text: str):
    """解析 YAML 子集，返回 Python 对象。"""
    lines = []
    for number, raw in enumerate(text.replace("\r\n", "\n").split("\n"), start=1):
        if number == 1 and raw.startswith("\ufeff"):
            raw = raw[1:]
        stripped = _strip_comment(raw).rstrip()
        if not stripped.strip():
            continue
        indent = len(stripped) - len(stripped.lstrip(" "))
        if "\t" in stripped[:indent]:
            raise YamlError("第 {} 行：缩进不得使用 Tab".format(number))
        if len(stripped) > 10000:
            raise YamlError("第 {} 行：单行超长（>10000 字符），疑似结构错误".format(number))
        lines.append([indent, stripped.strip(), number])

    if not lines:
        return None
    value, consumed = _parse_block(lines, 0, lines[0][0])
    if consumed != len(lines):
        raise YamlError("第 {} 行：缩进无法归属".format(lines[consumed][2]))
    return value


def _parse_block(lines, start, indent):
    if start >= len(lines):
        return None, start
    if lines[start][1].startswith("- "):
        return _parse_sequence(lines, start, indent)
    return _parse_mapping(lines, start, indent)


def _parse_sequence(lines, start, indent):
    items = []
    index = start
    while index < len(lines):
        line_indent, content, number = lines[index]
        if line_indent != indent or not content.startswith("- "):
            break
        rest = content[2:].strip()
        if rest == "":
            index += 1
            if index < len(lines) and lines[index][0] > indent:
                value, index = _parse_block(lines, index, lines[index][0])
            else:
                value = None
            items.append(value)
            continue
        if ":" in rest and not rest.startswith(("'", '"', "[")):
            key, value_text = _split_key(rest)
            item = {}
            if value_text:
                item[key] = _parse_scalar(value_text)
                index += 1
            else:
                index += 1
                if index < len(lines) and lines[index][0] > indent:
                    nested, index = _parse_block(lines, index, lines[index][0])
                    item[key] = nested
                else:
                    item[key] = None
            # 继续吸收同层（缩进 > indent）的后续 key
            while index < len(lines) and lines[index][0] > indent:
                sub_indent, sub_content, sub_number = lines[index]
                if sub_content.startswith("- "):
                    break
                sub_key, sub_value = _split_key(sub_content)
                if sub_key in item:
                    raise YamlError("第 {} 行：重复 key（{}）".format(sub_number, sub_key))
                if sub_value:
                    item[sub_key] = _parse_scalar(sub_value)
                    index += 1
                else:
                    index += 1
                    if index < len(lines) and lines[index][0] > sub_indent:
                        nested, index = _parse_block(lines, index, lines[index][0])
                        item[sub_key] = nested
                    else:
                        item[sub_key] = None
            items.append(item)
            continue
        items.append(_parse_scalar(rest))
        index += 1
    return items, index


def _parse_mapping(lines, start, indent):
    mapping = {}
    index = start
    while index < len(lines):
        line_indent, content, number = lines[index]
        if line_indent != indent:
            break
        if content.startswith("- "):
            break
        key, value_text = _split_key(content)
        if key in mapping:
            raise YamlError("第 {} 行：重复 key（{}）".format(number, key))
        if value_text:
            mapping[key] = _parse_scalar(value_text)
            index += 1
        else:
            index += 1
            if index < len(lines) and lines[index][0] > indent:
                value, index = _parse_block(lines, index, lines[index][0])
                mapping[key] = value
            else:
                mapping[key] = None
    return mapping, index


def load_yaml(path):
    with open(path, "r", encoding="utf-8") as handle:
        return parse_yaml(handle.read())


# ---------------------------------------------------------------- 文件与路径


def sha256_file(path, chunk=1024 * 1024) -> str:
    digest = hashlib.sha256()
    with open(path, "rb") as handle:
        for block in iter(lambda: handle.read(chunk), b""):
            digest.update(block)
    return digest.hexdigest()


def iter_files(base, suffixes=None, skip_dirs=("__pycache__", ".git", "node_modules")):
    base = Path(base)
    for current, dirs, names in os.walk(base):
        dirs[:] = sorted(d for d in dirs if d not in skip_dirs)
        for name in sorted(names):
            path = Path(current) / name
            if suffixes and path.suffix.lower() not in suffixes:
                continue
            yield path


def rel_posix(path, base) -> str:
    try:
        return Path(path).resolve().relative_to(Path(base).resolve()).as_posix()
    except ValueError:
        return Path(path).name


def find_repo_root(start=None) -> Path:
    """向上找含 plugins/cardo-invest 的目录；找不到就返回 start 的父级。"""
    current = Path(start or __file__).resolve()
    for candidate in [current] + list(current.parents):
        if (candidate / "plugins" / "cardo-invest").is_dir():
            return candidate
    return Path.cwd()


# ---------------------------------------------------------------- 数值


def dec(value) -> Decimal:
    if isinstance(value, Decimal):
        return value
    try:
        return Decimal(str(value))
    except (InvalidOperation, TypeError, ValueError):
        raise ValueError("不是合法数值：{!r}".format(value))


def q2(value) -> Decimal:
    """金额/比率：两位小数，四舍五入。"""
    return dec(value).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def q4(value) -> Decimal:
    return dec(value).quantize(Decimal("0.0001"), rounding=ROUND_HALF_UP)


def fmt(value) -> str:
    """把 Decimal 转成不带科学计数法的字符串（写 JSON/文档用）。"""
    value = dec(value)
    text = format(value, "f")
    if "." in text:
        text = text.rstrip("0").rstrip(".")
    return text or "0"


# ---------------------------------------------------------------- 时间与 JSON


def iso_now(no_timestamp: bool = False):
    if no_timestamp:
        return None
    return datetime.now().astimezone().isoformat(timespec="seconds")


def envelope(schema, schema_version=1, producer=None, sources=None, gaps=None, no_timestamp=False):
    payload = {
        "schema": schema,
        "schemaVersion": schema_version,
        "generatedAt": iso_now(no_timestamp),
        "producer": producer or "cardo-invest-report/0.1.0",
        "source": sources or [],
        "gaps": gaps or [],
    }
    return payload


def write_json(path, payload, no_timestamp: bool = False):
    """写 JSON；no_timestamp 时把 generatedAt 置 null，保证两次跑 sha256 一致。"""
    if no_timestamp and isinstance(payload, dict) and "generatedAt" in payload:
        payload = dict(payload)
        payload["generatedAt"] = None
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8") as handle:
        json.dump(payload, handle, ensure_ascii=False, indent=2, sort_keys=True)
        handle.write("\n")


def read_json(path):
    with open(path, "r", encoding="utf-8") as handle:
        return json.load(handle)