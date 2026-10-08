#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CONFIG FORGE 引擎运行器 —— 在包内 engine_runtime 目录执行（使用 runtime/ 便携 Python）。
由 config_forge skill 通过子进程调用，传入 spec.json，输出 {output: xlsx路径}。

只负责：把 spec 转成 main.write_xlsx 需要的 devices 结构并渲染写盘。
不依赖 route-D 程序的任何包；openpyxl 等由 CG 环境提供。
"""
import json
import re
import sys
from pathlib import Path

# 自包含包：引擎文件与本脚本同目录，直接从自身目录加载（不再依赖外部 config-generator）
CFG = Path(__file__).resolve().parent
sys.path.insert(0, str(CFG))

import library
import engine  # noqa: F401  (确保引擎可用)
import main as cfg_main


# ---------------- 思科 interface range 语法修正 ----------------
# 规则：同前缀的接口范围，结束口只写端口号。
#   interface range g3/0/21 - g3/0/24  ->  interface range g3/0/21 - 24
#   default int range g3/0/21 - g3/0/24 ->  default int range g3/0/21 - 24
# 这样无论 LLM / 手填把结束口写成 g3/0/24 还是 24，渲染结果都正确。
_RANGE_RE = re.compile(
    r"((?:interface|int|default\s+int)\s+range\s+)([^\s]+)(\s+-\s+)([^\s]+)"
)


def _shorten_end(start_if, end_if):
    if "/" in start_if:
        prefix = start_if[: start_if.rfind("/") + 1]  # 含末尾 '/'
        if end_if.startswith(prefix):
            return end_if[len(prefix):]
    return end_if


def normalize_range_lines(path):
    """遍历所有工作表 A 列，把接口范围结束口按思科语法缩短。"""
    import openpyxl
    wb = openpyxl.load_workbook(path)
    changed = False
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            cell = row[0]
            val = cell.value
            if not isinstance(val, str):
                continue
            m = _RANGE_RE.search(val)
            if not m:
                continue
            new_end = _shorten_end(m.group(2), m.group(4))
            if new_end != m.group(4):
                cell.value = val[: m.start(4)] + new_end + val[m.end(4):]
                changed = True
    if changed:
        wb.save(path)


def _find_by_name(idx, name):
    for e in idx["templates"]:
        if e["name"] == name:
            return e
    return None


def _split_repeat_item(item):
    """从 repeats 的一个 list 元素里析出 (节点id, 变量dict)。解析失败返回 (None, {})。

    兼容 LLM 吐出的多种形态：
      {"id":"r0","vars":{...}}          -> ("r0", {vars})
      {"r0":{"member_...":...}}         -> ("r0", {member_...})   （单键 dict，键即节点id）
      {"id":"r0","member_...":...}      -> ("r0", {member_...})   （变量直接挂在 item 上）
    """
    if not isinstance(item, dict):
        return None, {}
    if "id" in item:
        rid = item["id"]
        v = item.get("vars", item)
        return rid, (v if isinstance(v, dict) else {})
    if len(item) == 1:
        k, v = next(iter(item.items()))
        return k, (v if isinstance(v, dict) else {})
    return None, {}


def _normalize_template_data(data):
    """把 per-template data 归一成引擎期望的形态。

    引擎 render_nodes 期望：
      data["repeats"] = { 节点id: [变量dict, ...] }
      data["choices"] = { 节点id: 选中的option名 }
    但 LLM/填表流可能吐出多种非规范形态，这里尽量兼容：
      - data 套了一层多余键名（如按 section 命名：
        {"used_interface_configuration":{"repeats":[...]}}）→ 顶层无标准键时解开这一层。
      - repeats 是 list，每个 item 可能是 {"id":..,"vars":..} / {"r0":{..}} / {"id":..,..}。
      - choices 类似：list 里 {"id":"c0","value":"xxx"} 等。
    空 list / 缺失 / 套壳 都归一成 {}，避免引擎对 list 调 .get 抛 AttributeError。
    """
    if not isinstance(data, dict):
        return {"vars": {}, "repeats": {}, "choices": {}}
    # 解开一层多余键名（LLM 按 section 套壳的情况）
    _KNOWN = ("vars", "repeats", "choices")
    if not any(k in data for k in _KNOWN):
        sub = [k for k, v in data.items() if isinstance(v, dict)]
        if len(sub) == 1:
            data = data[sub[0]]

    data.setdefault("vars", {})
    if not isinstance(data["vars"], dict):
        data["vars"] = {}

    raw_repeats = data.get("repeats")
    if isinstance(raw_repeats, list):
        norm = {}
        for item in raw_repeats:
            rid, vars_ = _split_repeat_item(item)
            if rid is not None:
                norm.setdefault(rid, []).append(vars_)
        data["repeats"] = norm
    elif not isinstance(raw_repeats, dict):
        data["repeats"] = {}

    raw_choices = data.get("choices")
    if isinstance(raw_choices, list):
        norm = {}
        for item in raw_choices:
            if not isinstance(item, dict):
                continue
            cid = item.get("id")
            val = None
            if cid is None and len(item) == 1:
                k, v = next(iter(item.items()))
                cid, val = k, v
            if cid is None:
                continue
            if val is None:
                for key in ("value", "choice", "name", "option"):
                    if key in item and item[key] not in (None, ""):
                        val = item[key]
                        break
            if val is not None:
                norm[cid] = val
        data["choices"] = norm
    elif not isinstance(raw_choices, dict):
        data["choices"] = {}

    return data


def build_devices(spec, idx):
    devices = []
    for d in spec.get("devices", []):
        category = d.get("category")
        templates = []
        for t in d.get("templates", []):
            entry = (
                library.find_entry(idx, category, t["name"])
                if category
                else _find_by_name(idx, t["name"])
            )
            if entry is None:
                raise ValueError(f"模板未找到: {t['name']} (category={category})")
            schema = library.load_schema(entry, t.get("version"))
            # data 来源兼容：优先"模板内 data"（LLM 有时把数据塞进 templates[i].data），
            # 否则取"设备级 data[模板名]"，再否则整体设备级 data，最后回退空壳。
            tpl_data = t.get("data")
            dev_data = d.get("data")
            if isinstance(tpl_data, dict) and tpl_data:
                data = tpl_data
            elif isinstance(dev_data, dict):
                data = dev_data.get(t["name"]) or dev_data
            else:
                data = {"vars": {}, "repeats": {}, "choices": {}}
            # 归一化：LLM/填表流可能吐非规范形态（list / 套壳键名 / 模板内 data），引擎要求 dict 形态
            data = _normalize_template_data(data)
            templates.append((schema, data))
        devices.append(
            {
                "hostname": d["hostname"],
                "category": category,
                "stack_member": d.get("stack_member"),
                "templates": templates,
            }
        )
    return devices


def main_run(spec_path):
    spec = json.loads(Path(spec_path).read_text(encoding="utf-8"))
    idx = library.load_index()
    if not idx.get("templates"):
        raise RuntimeError("模板库为空，请先在 config-generator 内 import 模板")
    devices = build_devices(spec, idx)
    out = cfg_main.write_xlsx(devices)
    # 思科 interface range 语法修正：结束口只保留端口号
    normalize_range_lines(out)
    print(json.dumps({"output": str(out)}, ensure_ascii=False))


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("usage: _forge_runner.py <spec.json>")
        sys.exit(2)
    main_run(sys.argv[1])
