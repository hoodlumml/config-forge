#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
模板导入器 v3：支持 .xlsx（用户主格式）和 .txt
- xlsx: A列=命令，B~E列合并为注释/提示；自动提取嵌入图片
- 全量样式捕获：字体/底色/边框/对齐/行高/列宽（主题色自动换算 RGB）
- 文本哈希 + 样式指纹双重对比：样式-only 变化也可检测并单独更新
- 自动识别 stack_config / stack_picture 专用段
- 导入后生成"需人工检查"报告（含解析完整性校验）
"""
import hashlib
import json
import re
from datetime import datetime
from difflib import unified_diff
from pathlib import Path

import engine
from engine import SECTION_RE, slugify

STACK_MEMBER_RE = re.compile(r"if\s*\{\{\s*stack\s*member\s*\}\}\s*==\s*(\d+)", re.I)


# ---------------- 颜色解析（RGB / 主题色 / 索引色） ----------------

def _theme_colors(wb):
    """从工作簿主题解析调色板，返回 [lt1, dk1, lt2, dk2, accent1..6] 的 RRGGBB"""
    try:
        from openpyxl.xml.functions import fromstring
        theme = wb.loaded_theme
        if not theme:
            return []
        root = fromstring(theme)
        ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
        scheme = root.find(".//a:clrScheme", ns)
        if scheme is None:
            return []
        raw = {}
        for child in scheme:
            name = child.tag.split("}")[-1]
            srgb = child.find("a:srgbClr", ns)
            sysc = child.find("a:sysClr", ns)
            if srgb is not None:
                raw[name] = srgb.get("val")
            elif sysc is not None:
                raw[name] = sysc.get("lastClr")
        order = ["lt1", "dk1", "lt2", "dk2",
                 "accent1", "accent2", "accent3", "accent4", "accent5", "accent6"]
        return [raw[n] for n in order if raw.get(n)]
    except Exception:
        return []


def _apply_tint(rgb, tint):
    if not tint:
        return "FF" + rgb.upper()
    out = []
    for i in (0, 2, 4):
        c = int(rgb[i:i + 2], 16)
        if tint < 0:
            c = c * (1.0 + tint)
        else:
            c = c * (1.0 - tint) + 255 * tint
        out.append(max(0, min(255, round(c))))
    return "FF%02X%02X%02X" % tuple(out)


def resolve_color(color, themes):
    """把 openpyxl Color 统一成 'FFRRGGBB'；无法解析返回 None"""
    if color is None:
        return None
    try:
        if color.type == "rgb" and isinstance(color.rgb, str):
            if color.rgb not in ("00000000",):
                return color.rgb
        if color.type == "theme" and color.theme is not None:
            if 0 <= color.theme < len(themes):
                return _apply_tint(themes[color.theme], color.tint or 0)
        if color.type == "indexed" and color.indexed is not None:
            from openpyxl.styles.colors import COLOR_INDEX
            if 0 <= color.indexed < len(COLOR_INDEX):
                return COLOR_INDEX[color.indexed]
    except Exception:
        pass
    return None


# ---------------- 样式捕获 ----------------

def cell_style(cell, themes):
    """捕获单元格全量样式：字体/底色/边框/对齐"""
    st = {}
    f = cell.font
    if f:
        if f.name:
            st["font"] = f.name
        if f.size:
            st["size"] = f.size
        if f.bold:
            st["bold"] = True
        if f.italic:
            st["italic"] = True
        if f.underline:
            st["underline"] = True
        if f.strike:
            st["strike"] = True
        c = resolve_color(f.color, themes)
        if c:
            st["color"] = c
    fill = cell.fill
    if fill and fill.patternType:
        c = resolve_color(fill.fgColor, themes)
        if c:
            st["fill"] = c
            st["fill_type"] = fill.patternType
    al = cell.alignment
    if al:
        if al.horizontal:
            st["align"] = al.horizontal
        if al.vertical:
            st["valign"] = al.vertical
        if al.indent:
            st["indent"] = al.indent
        if al.wrap_text:
            st["wrap"] = True
    b = cell.border
    if b:
        for side_name in ("left", "right", "top", "bottom"):
            side = getattr(b, side_name, None)
            if side and side.style:
                sd = {"style": side.style}
                c = resolve_color(side.color, themes)
                if c:
                    sd["color"] = c
                st["border_" + side_name] = sd
    return st


def fingerprint_styles(col_widths, line_styles):
    """样式指纹：与文本内容无关，专用于检测"仅样式变化" """
    payload = json.dumps({"w": col_widths, "s": line_styles},
                         sort_keys=True, ensure_ascii=False, default=str)
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()[:12]


def read_styles_fp(d):
    sp = Path(d) / "styles.json"
    if not sp.exists():
        return None
    st = json.loads(sp.read_text(encoding="utf-8"))
    return fingerprint_styles(st.get("col_widths", {}), st.get("line_styles", []))


# ---------------- 读取模板文件 ----------------

def read_txt_sections(path):
    rows = [(t, c, {}, {}, None) for t, c in engine.split_lines(Path(path).read_text(encoding="utf-8"))]
    return _split_sections(rows, {}, {})


def read_xlsx_sections(path):
    from openpyxl import load_workbook
    wb = load_workbook(path)
    ws = wb.active
    themes = _theme_colors(wb)
    rows = []
    for idx, r in enumerate(ws.iter_rows(), 1):
        a = r[0]
        text = str(a.value).rstrip() if a.value is not None else ""
        comment = " | ".join(
            str(c.value).strip() for c in r[1:]
            if c.value is not None and str(c.value).strip()
        )
        style_b = cell_style(r[1], themes) if len(r) > 1 and r[1].value is not None else {}
        height = None
        dim = ws.row_dimensions.get(idx)
        if dim and dim.height:
            height = dim.height
        rows.append((text, comment, cell_style(a, themes), style_b, height))
    images = {}
    for img in ws._images:
        images[img.anchor._from.row + 1] = (img._data(), img.width, img.height)
    col_widths = {}
    for col in ("A", "B"):
        dim = ws.column_dimensions.get(col)
        if dim and dim.width:
            col_widths[col] = dim.width
    return _split_sections(rows, images, col_widths)


def _split_sections(rows, images, col_widths):
    """按 ! --- 段名 --- 切分；rows: [(text, comment, styleA, styleB, height)]"""
    sections = []
    title, body = None, []
    start_row = 1

    def flush(end_row):
        if title is None and not any(t.strip() for t, _, _, _, _ in body):
            return
        sections.append({"title": title or "Untitled", "lines": body,
                         "start_row": start_row, "end_row": end_row})

    for idx, row in enumerate(rows, 1):
        m = SECTION_RE.match(row[0].strip())
        if m:
            flush(idx - 1)
            title, body = m.group(1), []
            start_row = idx + 1
        else:
            body.append(row)
    flush(len(rows))

    for sec in sections:
        kept = [r for r in sec["lines"] if r[0].strip()]
        sec["lines"] = [(t, c) for t, c, _, _, _ in kept]
        sec["styles"] = [{"a": sa, "b": sb, "h": h} for _, _, sa, sb, h in kept]
        sec["col_widths"] = col_widths
        t = sec["title"].lower()
        if "stack" in t and ("picture" in t or "image" in t or "photo" in t):
            sec["type"] = "stack_picture"
            member_of_row = {}
            current_n = None
            for i, (text, _) in enumerate(sec["lines"]):
                m = STACK_MEMBER_RE.search(text)
                if m:
                    current_n = m.group(1)
                member_of_row[sec["start_row"] + i] = current_n
            sec["images"] = {}
            for row_no, blob in images.items():
                if sec["start_row"] <= row_no <= sec["end_row"]:
                    n = member_of_row.get(row_no)
                    if n:
                        sec["images"][n] = blob
        elif "stack" in t and "config" in t:
            sec["type"] = "stack_config"
        else:
            sec["type"] = "normal"
    return sections


def section_to_template_text(lines):
    out = []
    for text, comment in lines:
        out.append(f"{text} ## {comment}" if comment else text)
    return "\n".join(out) + "\n"


def build_schema(name, label, sec_type, images=None):
    return {
        "name": name,
        "label": label,
        "section": label,
        "type": sec_type,
        "stack_related": sec_type in ("stack_config", "stack_picture"),
        "images": images or {},
    }


def detect_stack_related(sec):
    if sec["type"] in ("stack_config", "stack_picture"):
        return True
    text = "\n".join(t for t, _ in sec["lines"])
    return bool(re.search(r"\{\{\s*stack\s*member\s*\}\}", text, re.I))


# ---------------- 入库 ----------------

def _write_styles(d, sec):
    (Path(d) / "styles.json").write_text(json.dumps({
        "col_widths": sec.get("col_widths", {}),
        "line_styles": sec.get("styles", []),
    }, ensure_ascii=False), encoding="utf-8")


def _write_version(d, sec, body_text, name, label, out=print):
    d = Path(d)
    d.mkdir(parents=True, exist_ok=True)
    (d / "template.txt").write_text(body_text, encoding="utf-8")
    image_map = {}
    if sec["type"] == "stack_picture" and sec.get("images"):
        img_dir = d / "images"
        img_dir.mkdir(exist_ok=True)
        for n, (blob, w, h) in sec["images"].items():
            fname = f"stack_{n}.png"
            (img_dir / fname).write_bytes(blob)
            image_map[n] = f"images/{fname}"
    schema = build_schema(name, label, sec["type"], image_map)
    schema["stack_related"] = detect_stack_related(sec)
    (d / "schema.json").write_text(
        json.dumps(schema, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    if sec.get("styles"):
        _write_styles(d, sec)
    if image_map:
        out(f"    已提取 {len(image_map)} 张图片到 {d / 'images'}")
    return schema


def import_sections(sections, io_hooks, library):
    """逐段导入，返回 [(dir, schema)] 本次实际写入的版本/样式"""
    ask, yes_no, out = io_hooks["ask"], io_hooks["yes_no"], io_hooks["print"]
    idx = library.load_index()
    category = ask("模板分类 (如 switch/access, 可输入新分类)")
    # 客户（与分类正交，仅作筛选维度）：整模板级、不按段拆分；留空 = 未分组。
    # 问句放在分类之后、逐段问询之前——ask() 靠关键词识别，"客户"不与既有词冲突。
    client = str(ask("客户名称 (如 华为/华三, 留空=未分组)", default="") or "").strip()
    written = []

    for sec in sections:
        body_text = section_to_template_text(sec["lines"])
        new_fp = fingerprint_styles(sec.get("col_widths", {}), sec.get("styles", []))
        out(f"\n=== 段落: {sec['title']} ({sec['type']}) ===")
        for t, _ in sec["lines"][:8]:
            out("    " + t)
        if len(sec["lines"]) > 8:
            out(f"    ... 共 {len(sec['lines'])} 行")
        if sec["type"] == "stack_picture":
            out(f"    图片: member {sorted(sec['images'], key=int)} 共 {len(sec['images'])} 张")
        if not yes_no("导入该段落吗", True):
            continue

        name = ask("模板名(英文标识)", default=slugify(sec["title"]))
        label = ask("显示名称", default=sec["title"])
        cat = ask("分类", default=category)
        entry = library.find_entry(idx, cat, name)
        # 已有条目时允许改挂客户（留空表示移到未分组）；模板名禁止含 "@"——
        # 它是 app.py 内部「唯一名」的客户分隔符，冲突会导致引擎按错误模板名查找
        if client and "@" in name:
            raise ValueError(f"模板名不能包含 '@'（客户分隔符）: {name}")

        if entry:
            cur = library.active_version(entry)
            d_cur = library.version_dir(entry, cur)
            old_text = (d_cur / "template.txt").read_text(encoding="utf-8")
            if library.file_hash(old_text) == library.file_hash(body_text):
                # 文本相同：检查样式指纹
                old_fp = read_styles_fp(d_cur)
                if old_fp == new_fp:
                    out(f"[无变化] {cat}/{name}，跳过")
                    continue
                # 样式-only 变化：静默自动更新（解析预览已提示，不逐段询问）
                _write_styles(d_cur, sec)
                written.append((d_cur, library.load_schema(entry, cur)))
                out(f"[样式已更新] {cat}/{name} {cur}（文本无变化，不产生新版本）")
                continue
            out(f"[有变化] {cat}/{name}（当前 {cur}），差异:")
            for l in unified_diff(old_text.splitlines(), body_text.splitlines(), lineterm="", n=1):
                out("    " + l)
            choice = ask("[a] 存为新版本  [b] 覆盖当前版本  [c] 跳过", default="a")
            if choice == "c":
                continue
            if choice == "b":
                schema = _write_version(d_cur, sec, body_text, name, label, out)
                for v in entry["versions"]:
                    if v["version"] == cur:
                        v["hash"] = library.file_hash(body_text)
                        v["date"] = f"{datetime.now():%Y-%m-%d}"
                written.append((d_cur, schema))
                out(f"已覆盖 {cat}/{name} {cur}")
            else:
                ver = ask("新版本号", default=library.bump_version(entry["versions"]))
                d = library.version_dir(entry, ver)
                schema = _write_version(d, sec, body_text, name, label, out)
                entry["versions"].append({
                    "version": ver, "date": f"{datetime.now():%Y-%m-%d}",
                    "hash": library.file_hash(body_text),
                })
                entry["active"] = ver
                # 已有条目的客户归属保持不变（新问到的 client 不覆盖，避免改内容时丢掉归属）
                written.append((d, schema))
                out(f"已入库 {cat}/{name} 新版本 {ver}（旧版本保留）")
        else:
            ver = ask("版本号", default="v1.0")
            d = library.LIBRARY_DIR / cat / name / ver
            schema = _write_version(d, sec, body_text, name, label, out)
            idx["templates"].append({
                "category": cat, "name": name, "label": label, "active": ver,
                "versions": [{
                    "version": ver, "date": f"{datetime.now():%Y-%m-%d}",
                    "hash": library.file_hash(body_text),
                }],
            })
            # 客户字段（与分类正交）；留空则不写键，避免留下空串
            if client:
                idx["templates"][-1]["client"] = client
            written.append((d, schema))
            out(f"[新增] 已入库 {cat}/{name} {ver}")

    library.save_index(idx)
    out("\n导入完成。查看模板库: python main.py list")
    return written


# ---------------- 段落状态检测（解析预览用） ----------------

def section_status(sec, cat, library, idx=None):
    """对比模板库，返回 (状态, entry)：
    新增 / 内容有变化 / 样式有变化 / 无变化"""
    idx = idx or library.load_index()
    name = slugify(sec["title"])
    entry = library.find_entry(idx, cat, name)
    if not entry:
        return "新增", None
    cur = library.active_version(entry)
    d = library.version_dir(entry, cur)
    old_text = (d / "template.txt").read_text(encoding="utf-8")
    if library.file_hash(old_text) != library.file_hash(section_to_template_text(sec["lines"])):
        return "内容有变化", entry
    new_fp = fingerprint_styles(sec.get("col_widths", {}), sec.get("styles", []))
    if read_styles_fp(d) != new_fp:
        return "样式有变化", entry
    return "无变化", entry


# ---------------- 导入后检查报告 ----------------

def section_warnings(d, schema):
    """生成该模板"需要人工确认/修改"的检查项"""
    d = Path(d)
    warns = []
    t = schema.get("type", "normal")
    if t == "stack_picture":
        have = set(schema.get("images", {}))
        missing = [str(n) for n in range(2, 10) if str(n) not in have]
        if missing:
            warns.append(
                f"[{schema['label']}] 缺少成员数 {', '.join(missing)} 的连线图。"
                f"如需要，请把图片放入 {d / 'images'} 并在 schema.json 的 images 中登记")
    elif t == "stack_config":
        warns.append(
            f"[{schema['label']}] 堆叠参数化段：priority 默认 16-成员号、"
            f"renumber 为手动逐条添加，请确认符合你的习惯")
    else:
        nodes, prompts = engine.load_template(d, schema)
        raw_lines = engine.split_lines((d / "template.txt").read_text(encoding="utf-8"))
        lost = engine.verify_parse(raw_lines, nodes)
        if lost:
            warns.append(
                f"[{schema['label']}] ⚠ 严重：解析后丢失 {len(lost)} 行"
                f"（首行: {lost[0]!r}），这是工具 bug，请报告并暂停使用该模板")
        if not prompts["vars"] and not prompts["repeats"] and not prompts["choices"]:
            warns.append(
                f"[{schema['label']}] 未检测到任何 {{{{变量}}}}，是纯固定文本段落，请确认是否符合预期")
        all_vars = prompts["vars"] + [v for r in prompts["repeats"] for v in r["vars"]]
        for v in all_vars:
            warns.append(
                f"[{schema['label']}] 变量 '{v['display']}' 猜测类型为 {v['type']}"
                f"（如不对请修改 template.txt 中变量名或告知调整规则）")
        if not (d / "styles.json").exists():
            warns.append(f"[{schema['label']}] 无样式信息（txt 导入），输出将使用默认字体")
    return warns
