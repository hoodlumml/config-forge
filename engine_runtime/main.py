#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
配置生成器 v3.1 —— 模板库 + 专用堆叠段 + 图形界面（纯离线）
输出 Excel 严格沿用模板字体样式（styles.json）。

用法:
  python main.py                CLI 生成向导
  python main.py gui            图形界面（选择类操作推荐）
  python main.py import <文件>  导入模板（.xlsx 或 .txt），导入后打印检查报告
  python main.py list           列出模板库
"""
import sys
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

import engine
import library

INVALID_SHEET_CHARS = set('[]:*?/\\')
MAX_IMG_WIDTH = 640
ROW_HEIGHT_PX = 20


# ---------------- 输入与校验 ----------------

def ask(prompt, default=None, required=False, validator=None, hint=None):
    if hint:
        print(f"  提示: {hint}")
    while True:
        suffix = f" [{default}]" if default not in (None, "") else ""
        try:
            raw = input(f"{prompt}{suffix}: ").strip()
        except EOFError:
            print("\n输入结束，退出。")
            sys.exit(1)
        if not raw:
            if default not in (None, ""):
                return default
            if not required:
                return ""   # 规则：任何值都可留空
            print("  ! 该字段必填")
            continue
        if validator:
            ok, msg = validator(raw)
            if not ok:
                print(f"  ! {msg}")
                continue
        return raw


def yes_no(prompt, default=True):
    d = "y" if default else "n"
    return ask(f"{prompt} (y/n)", default=d).lower().startswith("y")


def int_validator(raw):
    try:
        int(raw)
        return True, ""
    except ValueError:
        return False, "请输入整数"


# ---------------- 模板填写（CLI） ----------------

def fill_normal(schema, prompts, stack_member, prev=None):
    data = {"vars": {}, "repeats": {}, "choices": {}}
    prev = prev or {}

    for var in prompts["vars"]:
        prev_val = prev.get("vars", {}).get(var["name"])
        data["vars"][var["name"]] = ask(
            var["display"], default=prev_val, hint=var.get("hint") or None)

    for rep in prompts["repeats"]:
        prev_entries = prev.get("repeats", {}).get(rep["id"], [])
        entries = []
        if prev_entries and yes_no(f"沿用上一台的 {len(prev_entries)} 条 [{rep['label']}] 吗", True):
            entries = [dict(e) for e in prev_entries]
        else:
            print(f"\n  [{rep['label']}]")
            idx = 0
            while True:
                idx += 1
                if rep["member_bound"] and stack_member and idx > int(stack_member):
                    print(f"  ⚠ 已超过堆叠成员数 {stack_member}，请确认是否有误（不阻止继续）")
                print(f"  第 {idx} 条（输入 d 删除上一条）:")
                prev_entry = prev_entries[idx - 1] if idx <= len(prev_entries) else None
                entry = {}
                for var in rep["vars"]:
                    dflt = (prev_entry or {}).get(var["name"])
                    entry[var["name"]] = ask(var["display"], default=dflt,
                                             hint=var.get("hint") or None)
                entries.append(entry)
                if not yes_no("再添加一条吗", False):
                    break
        data["repeats"][rep["id"]] = entries

    for ch in prompts["choices"]:
        print(f"\n  [{ch['label']}]")
        for i, opt in enumerate(ch["options"], 1):
            print(f"    {i}. {opt}")
        prev_choice = prev.get("choices", {}).get(ch["id"])
        default = None
        if prev_choice in ch["options"]:
            default = str(ch["options"].index(prev_choice) + 1)
        while True:
            raw = ask("请选择编号", default=default)
            if raw == "":
                chosen = None
                break
            if raw.isdigit() and 1 <= int(raw) <= len(ch["options"]):
                chosen = ch["options"][int(raw) - 1]
                break
            print("  ! 无效编号")
        data["choices"][ch["id"]] = chosen
        if chosen:
            for var in ch["option_vars"].get(chosen, []):
                prev_val = prev.get("vars", {}).get(var["name"])
                data["vars"][var["name"]] = ask(
                    var["display"], default=prev_val, hint=var.get("hint") or None)
    return data


def fill_stack_config(schema, stack_member, prev=None):
    print(f"\n  [{schema['label']}] 参数化堆叠配置")
    n = stack_member
    if not n:
        n = ask("堆叠成员数", required=True, validator=int_validator)
    n = int(n)
    members = []
    prev_members = {m["number"]: m for m in (prev or {}).get("members", [])}
    for i in range(1, n + 1):
        default = (prev_members.get(i) or {}).get("priority", 16 - i)
        p = ask(f"Member {i} priority", default=default, validator=int_validator)
        members.append({"number": i, "priority": p})

    renumber = []
    prev_rn = (prev or {}).get("renumber", [])
    if prev_rn and yes_no(f"沿用上一台的 {len(prev_rn)} 条 renumber 配置吗", True):
        renumber = [dict(r) for r in prev_rn]
    elif yes_no("需要 renumber 配置吗", False):
        while True:
            cur = ask("  当前成员号 (switch X)", validator=int_validator)
            new = ask("  renumber 为", validator=int_validator)
            renumber.append({"cur": cur, "new": new})
            if not yes_no("再添加一条 renumber 吗", False):
                break
    return {"members": members, "renumber": renumber}


def fill_template(schema, stack_member, prev=None):
    print(f"\n--- {schema['label']} ({schema['_version']}) ---")
    t = schema.get("type", "normal")
    if t == "stack_config":
        return fill_stack_config(schema, stack_member, prev)
    if t == "stack_picture":
        n = stack_member or "?"
        img = schema.get("images", {}).get(str(n))
        print(f"  堆叠连线图: member={n} -> {img or '无对应图片'}（生成时自动嵌入）")
        return {}
    _, prompts = engine.load_template(schema["_dir"], schema)
    return fill_normal(schema, prompts, stack_member, prev)


# ---------------- 渲染与输出 ----------------

def render_device(device):
    """items: ("section", title) / ("row", cmd, comment, kind, style) / ("image", path, caption)"""
    items = []
    g = {"hostname": device["hostname"]}
    if device.get("stack_member"):
        g["stack_member"] = device["stack_member"]
    for schema, data in device["templates"]:
        items.append(("section", schema.get("section", schema["label"])))
        t = schema.get("type", "normal")
        if t == "stack_config":
            for cmd, comment, kind, style in engine.render_stack_config(data):
                items.append(("row", cmd, comment, kind, style))
        elif t == "stack_picture":
            n = str(device.get("stack_member") or "")
            img_rel = schema.get("images", {}).get(n)
            if img_rel:
                img = schema["_dir"] / img_rel
                if img.exists():
                    items.append(("image", img, f"Stack cabling diagram (member={n})"))
        else:
            nodes, _ = engine.load_template(schema["_dir"], schema)
            for cmd, comment, kind, style in engine.render_nodes(nodes, data, g):
                items.append(("row", cmd, comment, kind, style))
    return items


def sanitize_sheet_name(name):
    cleaned = "".join("_" if c in INVALID_SHEET_CHARS else c for c in name)
    return cleaned[:31] or "device"


def add_image(ws, img_path, start_row):
    img = XLImage(str(img_path))
    if img.width > MAX_IMG_WIDTH:
        ratio = MAX_IMG_WIDTH / img.width
        img.width = MAX_IMG_WIDTH
        img.height = int(img.height * ratio)
    img.anchor = f"A{start_row}"
    ws.add_image(img)
    return img.height // ROW_HEIGHT_PX + 2


def _apply_style(cell, st):
    """把捕获的模板样式（字体/底色/边框/对齐）完整应用到单元格"""
    if not st:
        return
    fkw = {}
    if st.get("font"):
        fkw["name"] = st["font"]
    if st.get("size"):
        fkw["size"] = st["size"]
    if st.get("bold"):
        fkw["bold"] = True
    if st.get("italic"):
        fkw["italic"] = True
    if st.get("underline"):
        fkw["underline"] = "single"
    if st.get("strike"):
        fkw["strike"] = True
    if st.get("color"):
        fkw["color"] = st["color"]
    if fkw:
        cell.font = Font(**fkw)
    if st.get("fill"):
        cell.fill = PatternFill(patternType=st.get("fill_type", "solid"),
                                fgColor=st["fill"])
    akw = {}
    if st.get("align"):
        akw["horizontal"] = st["align"]
    if st.get("valign"):
        akw["vertical"] = st["valign"]
    if st.get("indent"):
        akw["indent"] = st["indent"]
    if st.get("wrap"):
        akw["wrap_text"] = True
    if akw:
        cell.alignment = Alignment(**akw)
    bkw = {}
    for side_name in ("left", "right", "top", "bottom"):
        sd = st.get("border_" + side_name)
        if sd:
            bkw[side_name] = Side(style=sd["style"], color=sd.get("color"))
    if bkw:
        cell.border = Border(**bkw)


def write_xlsx(devices):
    library.OUTPUT_DIR.mkdir(exist_ok=True)
    wb = Workbook()
    wb.remove(wb.active)
    header_font = Font(bold=True)
    note_fallback = Font(italic=True, color="808080")
    # 列宽取第一个带样式信息的模板
    col_widths = {}
    for dev in devices:
        for schema, _ in dev["templates"]:
            st = engine.load_styles(schema["_dir"])
            if st and st.get("col_widths"):
                col_widths = st["col_widths"]
                break
        if col_widths:
            break

    for dev in devices:
        ws = wb.create_sheet(sanitize_sheet_name(dev["hostname"]))
        ws.append(["Command", "Comment"])
        for cell in ws[1]:
            cell.font = header_font
        for item in render_device(dev):
            if item[0] == "section":
                ws.append([f"! --- {item[1]} ---", ""])
                for cell in ws[ws.max_row]:
                    cell.font = header_font
            elif item[0] == "row":
                _, cmd, comment, kind, style = item
                ws.append([cmd, comment])
                row = ws[ws.max_row]
                if style:
                    _apply_style(row[0], style.get("a"))
                    _apply_style(row[1], style.get("b") or style.get("a"))
                    if style.get("h"):
                        ws.row_dimensions[ws.max_row].height = style["h"]
                elif kind == "note":
                    for cell in row:
                        cell.font = note_fallback
            elif item[0] == "image":
                ws.append([item[2], ""])
                for cell in ws[ws.max_row]:
                    cell.font = header_font
                rows = add_image(ws, item[1], ws.max_row + 1)
                for _ in range(rows):
                    ws.append([])
        ws.column_dimensions["A"].width = col_widths.get("A", 55)
        ws.column_dimensions["B"].width = col_widths.get("B", 45)
        ws.freeze_panes = "A2"
    out = library.OUTPUT_DIR / f"config_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    wb.save(out)
    return out


# ---------------- 生成向导 ----------------

def pick_category(idx, prev_cat=None):
    cats = library.categories(idx)
    print("\n设备分类:")
    for i, c in enumerate(cats, 1):
        print(f"  {i}. {c}")
    default = str(cats.index(prev_cat) + 1) if prev_cat in cats else None
    while True:
        raw = ask("请选择分类", default=default, required=True)
        if raw.isdigit() and 1 <= int(raw) <= len(cats):
            return cats[int(raw) - 1]
        if raw in cats:
            return raw
        print("  ! 无效选择")


def pick_templates(idx, category, prev_names=None):
    entries = library.entries_of(idx, category)
    print(f"\n可用模板 ({category}):")
    for i, e in enumerate(entries, 1):
        print(f"  {i}. {e['label']}  ({library.active_version(e)})")
    default = ",".join(
        str(i) for i, e in enumerate(entries, 1)
        if prev_names and e["name"] in prev_names
    ) or None
    sel = ask("请选择模板(可多选,逗号分隔; 编号@版本 可指定旧版本)",
              default=default, required=True)
    picked = []
    for token in sel.replace("，", ",").split(","):
        token = token.strip()
        if not token:
            continue
        ver = None
        if "@" in token:
            token, ver = token.split("@", 1)
        if token.isdigit() and 1 <= int(token) <= len(entries):
            e = entries[int(token) - 1]
            known = [v["version"] for v in e["versions"]]
            if ver and ver not in known:
                print(f"  ! {e['name']} 没有版本 {ver}，使用当前版本")
                ver = None
            picked.append((e, ver))
    return picked


def generate_wizard():
    idx = library.load_index()
    if not idx["templates"]:
        print("模板库为空，请先导入: python main.py import <模板文件>")
        sys.exit(1)
    print("=" * 50)
    print("配置生成器 (CLI)")
    print("=" * 50)

    devices = []
    while True:
        prev = devices[-1] if devices else None
        prev_map = prev_cat = None
        if prev and yes_no(f"\n基于上一台设备 {prev['hostname']} 复制再修改吗", True):
            prev_map = {s["name"]: d for s, d in prev["templates"]}
            prev_cat = prev["category"]

        hostname = ask("\n设备 hostname", required=True,
                       default=prev["hostname"] if prev else None)
        category = pick_category(idx, prev_cat)
        picked = pick_templates(idx, category, set(prev_map) if prev_map else None)
        if not picked:
            print("  ! 未选择有效模板，请重新选择")
            continue

        schemas = [library.load_schema(e, ver) for e, ver in picked]
        stack_member = None
        if any(s.get("stack_related") for s in schemas):
            prev_sm = prev.get("stack_member") if prev else None
            stack_member = ask("\n堆叠成员数 stack member (2-9, 非堆叠留空)",
                               default=prev_sm, validator=int_validator)

        templates = []
        for s in schemas:
            data = fill_template(s, stack_member,
                                 prev_map.get(s["name"]) if prev_map else None)
            templates.append((s, data))
        devices.append({"hostname": hostname, "category": category,
                        "stack_member": stack_member, "templates": templates})

        if not yes_no("\n再添加下一台设备吗", False):
            break

    out = write_xlsx(devices)
    print(f"\n✅ 已生成: {out}")
    print(f"   共 {len(devices)} 台设备: {', '.join(d['hostname'] for d in devices)}")


# ---------------- 命令 ----------------

def cmd_list():
    idx = library.load_index()
    if not idx["templates"]:
        print("模板库为空")
        return
    print(f"{'分类':<16} {'模板名':<24} {'显示名称':<26} {'版本':<12} 当前")
    for e in sorted(idx["templates"], key=lambda x: (x["category"], x["name"])):
        vers = ",".join(v["version"] for v in e["versions"])
        print(f"{e['category']:<16} {e['name']:<24} {e['label']:<26} {vers:<12} {library.active_version(e)}")


def cmd_import(path):
    import importer
    p = Path(path)
    if not p.exists():
        print(f"文件不存在: {path}")
        sys.exit(1)
    if p.suffix.lower() in (".xlsx", ".xlsm"):
        sections = importer.read_xlsx_sections(p)
    else:
        sections = importer.read_txt_sections(p)
    print(f"解析到 {len(sections)} 个段落: {', '.join(s['title'] for s in sections)}")
    hooks = {"ask": ask, "yes_no": yes_no, "print": print}
    written = importer.import_sections(sections, hooks, library)
    # 导入后检查报告：提示需要人工确认/修改的部分
    all_warns = []
    for d, schema in written:
        all_warns.extend(importer.section_warnings(d, schema))
    if all_warns:
        print("\n" + "=" * 50)
        print("导入后检查报告（请人工确认以下项目）:")
        print("=" * 50)
        for w in all_warns:
            print("  - " + w)
    else:
        print("\n检查通过，无需人工处理。")


def main():
    args = sys.argv[1:]
    if args and args[0] == "import":
        if len(args) < 2:
            print("用法: python main.py import <模板文件>")
            sys.exit(1)
        cmd_import(args[1])
        return
    if args and args[0] == "list":
        cmd_list()
        return
    if args and args[0] == "gui":
        import gui
        gui.run()
        return
    generate_wizard()


if __name__ == "__main__":
    main()
