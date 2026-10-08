#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
模板引擎 v3.1：解析模板结构（@repeat/@choice/stack 专用段）、渲染配置行。
行可携带模板源样式（styles.json），渲染时透传给输出。

模板纯文本约定：
  ! --- Section Name ---          功能段落分隔
  ! @repeat 提示语                 可重复块开始（逐条问答）
  ! @end                          块结束
  ! @choice 提示语                 选择块开始
  ! @option 选项名                 一个选项（正文到下一个 @option 或 @end 为止）
  {{任意变量名}}                   占位符（内部自动转 slug）
  命令 ## 注释                     注释进 Excel B 列，同时作为输入提示
"""
import json
import re

VAR_RE = re.compile(r"\{\{\s*(.+?)\s*\}\}")
SECTION_RE = re.compile(r"^!\s*-{2,}\s*(.+?)\s*-{2,}\s*$")


# ---------------- 基础工具 ----------------

def slugify(name):
    s = re.sub(r"[^a-zA-Z0-9]+", "_", name).strip("_").lower()
    return s or "var"


def placeholders(text):
    return VAR_RE.findall(text)


def substitute(text, values):
    """替换 {{占位符}}；留空/缺失 -> 空字符串，但保留整行（规则 B）"""
    def rep(m):
        v = values.get(slugify(m.group(1)))
        return "" if v is None else str(v)
    return VAR_RE.sub(rep, text).rstrip()


def guess_type(display):
    n = display.lower()
    if n.endswith("_id") or "number" in n or "count" in n or "priority" in n:
        return "int"
    if any(k in n for k in ("ip ", "ip a", "gateway", "server", "address")):
        return "ip"
    return "str"


# ---------------- 模板解析 ----------------

def split_lines(text):
    """文本 -> [(命令, 注释)]，空行跳过"""
    out = []
    for raw in text.splitlines():
        if not raw.strip():
            continue
        if "##" in raw:
            t, c = raw.split("##", 1)
            out.append((t.rstrip(), c.strip()))
        else:
            out.append((raw.rstrip(), ""))
    return out


def _unpack(line):
    """行元组兼容 (text, comment) 与 (text, comment, style)"""
    text, comment = line[0], line[1]
    style = line[2] if len(line) > 2 else None
    return text, comment, style


def parse_nodes(lines):
    """[(text, comment, style?)] -> 结构化节点列表"""
    nodes = []
    i = 0
    ri = ci = 0
    while i < len(lines):
        s = lines[i][0].strip()
        if s.startswith("! @repeat"):
            sub, j = _parse_until_end(lines, i + 1)
            nodes.append({
                "type": "repeat", "id": f"r{ri}",
                "label": s[len("! @repeat"):].strip(), "body": sub,
            })
            ri += 1
            i = j
        elif s.startswith("! @choice"):
            label = s[len("! @choice"):].strip()
            options = []
            i += 1
            while i < len(lines):
                s2 = lines[i][0].strip()
                if s2.startswith("! @option"):
                    sub, j = _parse_until_end(lines, i + 1, stop_at_option=True)
                    options.append({"name": s2[len("! @option"):].strip(), "body": sub})
                    i = j
                elif s2 == "! @end":
                    i += 1
                    break
                else:
                    # 选择块内不应出现选项外的行；遇到说明块已结束，直接退出，
                    # 绝不能继续跳过——否则会把 @end 之后的正常配置行全部吞掉
                    break
            nodes.append({"type": "choice", "id": f"c{ci}", "label": label, "options": options})
            ci += 1
        elif s == "! @end":
            i += 1
        else:
            text, comment, style = _unpack(lines[i])
            nodes.append({"type": "line", "text": text, "comment": comment, "style": style})
            i += 1
    return nodes


def _parse_until_end(lines, i, stop_at_option=False):
    """收集到 ! @end（消费）；stop_at_option 时遇 ! @option 也停（不消费）"""
    sub = []
    while i < len(lines):
        s = lines[i][0].strip()
        if s == "! @end":
            return sub, i + 1
        if stop_at_option and s.startswith("! @option"):
            return sub, i
        sub.append(lines[i])
        i += 1
    return sub, i


def vars_in_lines(lines):
    """从行列表收集变量元信息 {slug: {...}}"""
    out = {}
    for line in lines:
        text, comment, _ = _unpack(line)
        for ph in placeholders(text):
            slug = slugify(ph)
            if slug not in out:
                display = ph.strip()
                out[slug] = {
                    "name": slug,
                    "display": display,
                    "hint": comment or "",
                    "type": guess_type(display),
                }
    return out


def build_prompts(nodes):
    """从节点树生成提问结构（CLI 和 GUI 共用）"""
    prompts = {"vars": [], "repeats": [], "choices": []}
    top = {}
    for node in nodes:
        if node["type"] == "line":
            top.update(vars_in_lines([(node["text"], node["comment"])]))
        elif node["type"] == "repeat":
            vs = list(vars_in_lines(node["body"]).values())
            member_bound = any("member" in v["display"].lower() for v in vs)
            prompts["repeats"].append({
                "id": node["id"], "label": node["label"],
                "vars": vs, "member_bound": member_bound,
            })
        elif node["type"] == "choice":
            option_vars = {}
            for opt in node["options"]:
                option_vars[opt["name"]] = list(vars_in_lines(opt["body"]).values())
            prompts["choices"].append({
                "id": node["id"], "label": node["label"],
                "options": [o["name"] for o in node["options"]],
                "option_vars": option_vars,
            })
    prompts["vars"] = list(top.values())
    return prompts


# ---------------- 渲染 ----------------

def render_nodes(nodes, data, dev_globals):
    """返回 [(cmd, comment, kind, style)]，kind: cmd / note"""
    rows = []
    g = {k: v for k, v in dev_globals.items() if v is not None}
    top_vars = data.get("vars", {})
    for node in nodes:
        t = node["type"]
        if t == "line":
            cmd = substitute(node["text"], {**g, **top_vars})
            rows.append((cmd, node["comment"], _kind(cmd), node.get("style")))
        elif t == "repeat":
            for entry in data.get("repeats", {}).get(node["id"], []):
                merged = {**g, **top_vars, **entry}
                for line in node["body"]:
                    text, comment, style = _unpack(line)
                    cmd = substitute(text, merged)
                    rows.append((cmd, comment, _kind(cmd), style))
        elif t == "choice":
            chosen = data.get("choices", {}).get(node["id"])
            for opt in node["options"]:
                if opt["name"] == chosen:
                    for line in opt["body"]:
                        text, comment, style = _unpack(line)
                        cmd = substitute(text, {**g, **top_vars})
                        rows.append((cmd, comment, _kind(cmd), style))
    return rows


def _kind(cmd):
    return "note" if cmd.strip().startswith("!") else "cmd"


def verify_parse(lines, nodes):
    """完整性校验：模板里的每一行命令都必须出现在解析结果中（防吞行）。
    返回丢失的行列表（空 = 完整）。标记行（! @...）和段头（! ---）不参与校验。"""
    parsed = set()

    def collect(ns):
        for n in ns:
            if n["type"] == "line":
                parsed.add(n["text"].strip())
            elif n["type"] == "repeat":
                for l in n["body"]:
                    parsed.add(_unpack(l)[0].strip())
            elif n["type"] == "choice":
                for o in n["options"]:
                    for l in o["body"]:
                        parsed.add(_unpack(l)[0].strip())
    collect(nodes)

    missing = []
    for l in lines:
        t = _unpack(l)[0].strip()
        if t and not t.startswith("! @") and not SECTION_RE.match(t) and t not in parsed:
            missing.append(t)
    return missing


def render_stack_config(data):
    """参数化堆叠段：members=[{number, priority}], renumber=[{cur, new}]"""
    rows = []
    for m in data.get("members", []):
        p = m.get("priority")
        rows.append((f"switch {m['number']} priority {p if p not in (None, '') else ''}",
                     "", "cmd", None))
        rows.append(("!", "", "note", None))
    for r in data.get("renumber", []):
        cur, new = r.get("cur", ""), r.get("new", "")
        rows.append((f"switch {cur} renumber {new}", "", "cmd", None))
        rows.append(("reload", "", "cmd", None))
        rows.append(("!", "", "note", None))
    return rows


# ---------------- 模板加载 ----------------

def load_styles(version_dir):
    sp = version_dir / "styles.json"
    if sp.exists():
        return json.loads(sp.read_text(encoding="utf-8"))
    return None


def load_template(version_dir, schema):
    """加载模板文本并解析，返回 (nodes, prompts)；专用段返回 (None, None)。
    若有 styles.json，行样式按位置对齐附加到节点。"""
    tpl_path = version_dir / "template.txt"
    if schema.get("type") in ("stack_config", "stack_picture"):
        return None, None
    lines = split_lines(tpl_path.read_text(encoding="utf-8"))
    styles = load_styles(version_dir)
    if styles:
        line_styles = styles.get("line_styles", [])
        lines = [
            (t, c, line_styles[i] if i < len(line_styles) else None)
            for i, (t, c) in enumerate(lines)
        ]
    nodes = parse_nodes(lines)
    return nodes, build_prompts(nodes)
