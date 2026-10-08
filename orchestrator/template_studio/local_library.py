#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""模板库存取：index.json 注册表 + 版本目录"""
import hashlib
import json
import re
import sys
from pathlib import Path

if getattr(sys, "frozen", False):
    # PyInstaller 打包后：用户数据放在 exe 旁边的可写目录
    ROOT = Path(sys.executable).resolve().parent
else:
    ROOT = Path(__file__).resolve().parent
LIBRARY_DIR = ROOT / "library"
INDEX_FILE = LIBRARY_DIR / "index.json"
OUTPUT_DIR = ROOT / "output"


def load_index():
    if INDEX_FILE.exists():
        return json.loads(INDEX_FILE.read_text(encoding="utf-8"))
    return {"templates": []}


def save_index(idx):
    LIBRARY_DIR.mkdir(exist_ok=True)
    INDEX_FILE.write_text(
        json.dumps(idx, ensure_ascii=False, indent=2), encoding="utf-8"
    )


def file_hash(text):
    return hashlib.sha256(text.encode("utf-8")).hexdigest()[:12]


def find_entry(idx, category, name):
    for e in idx["templates"]:
        if e["category"] == category and e["name"] == name:
            return e
    return None


def version_dir(entry, version):
    return LIBRARY_DIR / entry["category"] / entry["name"] / version


def active_version(entry):
    return entry.get("active") or entry["versions"][-1]["version"]


def bump_version(versions):
    last = versions[-1]["version"]
    m = re.match(r"v(\d+)\.(\d+)", last)
    return f"v{m.group(1)}.{int(m.group(2)) + 1}" if m else last + ".1"


def load_schema(entry, version=None):
    v = version or active_version(entry)
    d = version_dir(entry, v)
    schema = json.loads((d / "schema.json").read_text(encoding="utf-8"))
    schema["_dir"] = d
    schema["_version"] = v
    schema["_category"] = entry["category"]
    return schema


def categories(idx):
    return sorted({e["category"] for e in idx["templates"]})


def entries_of(idx, category):
    return [e for e in idx["templates"] if e["category"] == category]


# ---------------- 客户分组（client）：与 category 正交，仅作筛选维度 ----------------
# 旧条目与官方模板均无 client 字段 -> 归入「未分组」，不报错、不迁移。


def client_of(entry):
    """条目所属客户名；缺失/空串/空白一律归一为「未分组」（空串）。"""
    c = (entry or {}).get("client") or ""
    return str(c).strip()


NO_CLIENT = ""  # 未分组


def clients_of(idx):
    """库里出现过的客户清单（去重、排序；不含未分组）。"""
    return sorted({client_of(e) for e in idx.get("templates", []) if client_of(e)})


def set_client(entry, client):
    """就地写入 client；空值直接删键，避免留下 "" 影响 find/save 往返。"""
    c = str(client or "").strip()
    if c:
        entry["client"] = c
    else:
        entry.pop("client", None)
    return entry

