"""跨平台本地 Web 服务（仅标准库，任意 OS 的 Python 3 直接跑）。

启动：  python app.py
演示：  浏览器打开 http://localhost:8000

本服务集成两层：
1. LLM 调度层（llm.py）：自动/手动模式，手动模式强制只调本地模型（dsh）。
2. Skill 层：CONFIG FORGE 配置生成（skills/config_forge.py），通过子进程调用
   包内 engine_runtime 的引擎（使用 runtime/ 便携 Python），确定性产出 Excel，不污染本程序依赖。
"""
import json
import os
import re
import shutil
import tempfile
import time
import sys
import urllib.parse
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

import llm as llm_module
from skills import config_forge

BASE_DIR = Path(__file__).resolve().parent
CONFIG_PATH = BASE_DIR / "config.json"
LOCAL_CONFIG_PATH = BASE_DIR / "config.local.json"
INDEX_PATH = BASE_DIR / "index.html"

# CONFIG FORGE 引擎目录（自包含包内 vendored 副本：同包的 engine_runtime/）
CG_DIR = BASE_DIR.parent / "engine_runtime"


def load_merged_config():
    """config.json 为基准；若存在 config.local.json（本机覆盖，含 API Key），叠加其 providers 覆盖。"""
    cfg = llm_module.load_config(CONFIG_PATH)
    if LOCAL_CONFIG_PATH.exists():
        try:
            local = json.loads(LOCAL_CONFIG_PATH.read_text(encoding="utf-8"))
            for name, overrides in local.get("providers", {}).items():
                if name in cfg["providers"]:
                    cfg["providers"][name].update(
                        {k: v for k, v in overrides.items() if v not in (None, "")}
                    )
        except Exception as e:
            print("读取 config.local.json 失败:", e)
    return cfg


def rebuild_orchestrator():
    global cfg, orchestrator
    cfg = load_merged_config()
    orchestrator = llm_module.build_orchestrator(cfg)


cfg = load_merged_config()
orchestrator = llm_module.build_orchestrator(cfg)


# ---------------- CONFIG FORGE 辅助（仅读库，无需 openpyxl） ----------------

def _norm_ws(text):
    """命令区文本归一：NBSP/全角空格/零宽字符 → 半角空格/删除。

    来源：Excel、网页、PDF 粘进来的中文资料常带这些字符。NBSP(\\u00a0) 在 GBK 里
    无法编码，Windows 子进程按 locale(GBK) 写 stdout 会 UnicodeEncodeError 崩溃
    （2026-10-05 现场报错）；即便传得回来，它进到 Cisco 命令区也是非法字符。
    单一实现在 template_studio/_norm.py，xlsx 探针与这里共用。
    """
    try:
        from template_studio._norm import norm_ws as _n
        return _n(text)
    except Exception:
        return (text or "").replace(chr(0x00A0), " ") \
                           .replace(chr(0x200B), "") \
                           .replace(chr(0xFEFF), "")


def _norm_deep(obj):
    """递归归一（list/tuple/dict 全覆盖），用于解析结果兜底。"""
    try:
        from template_studio._norm import norm_deep as _d
        return _d(obj)
    except Exception:
        return obj


def _cg_modules():
    """懒加载 CG 的 library / engine（二者仅依赖标准库，可在本进程直接 import）。"""
    if "library" not in sys.modules:
        sys.path.insert(0, str(CG_DIR))
        import library  # noqa: F401
        import engine  # noqa: F401
    return sys.modules["library"], sys.modules["engine"]


def _studio_client(entry):
    """取条目的 client（归一空白）；官方条目无此字段 -> ""（未分组）。"""
    return str((entry or {}).get("client") or "").strip()


def _official_index():
    library, _ = _cg_modules()
    return library.load_index()


def _studio_index():
    return _studio_library().load_index()


def list_clients():
    """客户清单：优先读 config.json 的 clients（手动维护），否则从两库聚合。"""
    try:
        cfg = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    except Exception:
        cfg = {}
    declared = cfg.get("clients")
    if isinstance(declared, list) and declared:
        return [str(x).strip() for x in declared if str(x).strip()]
    seen = set()
    for getter in (_official_index, _studio_index):
        try:
            for e in getter().get("templates", []) or []:
                c = _studio_client(e)
                if c:
                    seen.add(c)
        except Exception:
            continue
    return sorted(seen)


# ---------------- 客户上下文：唯一名编解码 ----------------
# 引擎按 (category, name) 精确查找，且引擎与 CG 不可改（CG 是外部只读依赖，portable 内是
# vendored 副本，升级会被覆盖）。所以「同一模板按客户各有一版」不能靠改引擎实现，
# 改用内部唯一名做桥：app.py 内把客户专属模板编码为「名称@客户」，交引擎前再解码回真名
# —— 引擎收到的永远是合法真名，全程零感知。
CLIENT_SEP = "@"


def _split_unique_name(unique):
    """'名称@客户' -> ('名称','客户')；无分隔符 -> (原名, '')。"""
    if not isinstance(unique, str):
        return unique, ""
    if CLIENT_SEP not in unique:
        return unique, ""
    name, _, client = unique.partition(CLIENT_SEP)
    return name.strip(), client.strip()


def _join_unique_name(name, client):
    """编码：仅客户专属模板加后缀；官方未覆盖的保持原名不动。"""
    c = (client or "").strip()
    return f"{name}{CLIENT_SEP}{c}" if c else name


def _resolve_client_templates(official, local, client):
    """按客户视角解析模板集，返回 [(条目, 是否顶替官方同名), ...]。

    - 该客户的条目顶替同名官方条目（覆盖键仍是 (category,name)，与引擎查找键一致）
    - 其他客户的条目不可见（客户之间互不干扰）
    - 该客户未覆盖的官方条目照常保留
    - client 为空（未分组/全部）：官方 + 未分组本地条目，保持旧行为
    """
    if not client:
        items, seen = [], set()
        for e in official:
            seen.add((e["category"], e["name"]))
            items.append((e, False))
        for e in local:
            if (e["category"], e["name"]) in seen:
                continue
            items.append((e, False))
        return items

    mine = [e for e in local if _studio_client(e) == client]
    mine_by_key = {(e["category"], e["name"]): e for e in mine}
    items = []
    for e in official:
        if (e["category"], e["name"]) in mine_by_key:
            continue  # 被该客户覆盖，跳过官方版
        items.append((e, False))
    for e in mine:
        items.append((e, True))  # 客户版顶替官方位置
    return items


def _tpl_item_for_client(lib, entry, override=False):
    """库条目 -> 前端/规划用 item（含 client、real_name、override、唯一名 name）。"""
    schema = lib.load_schema(entry)
    t = schema.get("type", "normal")
    client = _studio_client(entry)
    item = {
        "category": entry["category"],
        "name": _join_unique_name(entry["name"], client),
        "real_name": entry["name"],
        "client": client,
        "override": bool(override),
        "label": entry["label"],
        "version": lib.active_version(entry),
        "type": t,
    }
    if t == "normal":
        nodes, prompts = _cg_modules()[1].load_template(schema["_dir"], schema)
        if prompts:
            item["prompts"] = prompts
    return item


def list_templates(client=""):
    """可用模板集。client 非空按该客户视角解析（覆盖+隔离）；为空则兼容旧行为。"""
    client = (client or "").strip()
    official_entries = _official_index().get("templates", []) or []
    try:
        local_entries = _studio_index().get("templates", []) or []
    except Exception as e:
        print("读取本地模板库失败:", e)
        local_entries = []

    resolved = _resolve_client_templates(official_entries, local_entries, client)
    official_keys = {(e["category"], e["name"]) for e in official_entries}
    local_keys = {(e["category"], e["name"]) for e in local_entries}
    official_lib, studio_lib = _cg_modules()[0], _studio_library()

    out = []
    for entry, flagged in resolved:
        key = (entry["category"], entry["name"])
        is_local = key in local_keys
        lib = studio_lib if is_local else official_lib
        override = bool(flagged and is_local and key in official_keys)
        try:
            out.append(_tpl_item_for_client(lib, entry, override=override))
        except Exception as e:
            print(f"读取模板失败 {entry.get('name')}: {e}")
    return out


def _q_client(query):
    """从 query string 取 client 参数（缺失即空 = 全部客户）。"""
    try:
        vals = urllib.parse.parse_qs(query or "").get("client") or [""]
        return str(vals[0]).strip()
    except Exception:
        return ""


def _iter_prompt_vars(prompts):
    """遍历 prompts 中所有变量（含 top-level vars 与 @repeat 块内的变量）。"""
    out = []
    if not prompts:
        return out
    out.extend(prompts.get("vars", []) or [])
    for r in prompts.get("repeats", []) or []:
        out.extend(r.get("vars", []) or [])
    return out


# ---------------- 规划时只注入"相关模板"，避免全库塞进 prompt 撑爆上下文 ----------------
_CJK = re.compile(r"[\u4e00-\u9fff]+")
_ASCII = re.compile(r"[a-z0-9_]+")


def _template_search_text(t):
    """模板可被检索的文本：name + label + category + 各变量显示名。"""
    parts = [t.get("name", ""), t.get("label", ""), t.get("category", "")]
    for v in _iter_prompt_vars(t.get("prompts")):
        parts.append(v.get("display") or v.get("name", ""))
    return " ".join(str(p) for p in parts).lower()


# 通用停用词：几乎每个模板标签里都有，毫无区分度，检索时直接剔除
_STOPWORDS = {
    "配置", "生成", "交换机", "接口", "端口", "管理", "设备", "模板", "网络", "需求", "规划", "系统",
    "信息", "设置",
    "config", "configuration", "template", "templates", "switch", "switches", "port", "ports",
    "interface", "interfaces", "management", "device", "devices", "generate", "plan", "requirement",
    "system", "info", "setting", "settings", "basic", "core",
}


def _tokenize(text):
    """抽取检索词：ASCII 词（≥2）直接取；中文按「整串 + 二元组」允许局部重叠匹配；通用停用词剔除。"""
    text = (text or "").lower()
    toks = set()
    for m in _ASCII.findall(text):
        if len(m) >= 2:
            toks.add(m)
    for m in _CJK.findall(text):
        toks.add(m)
        for i in range(len(m) - 1):
            toks.add(m[i:i + 2])
    toks.difference_update(_STOPWORDS)
    return toks


# 模板检索索引缓存：模板集合不变时复用 haystack，避免每次规划都重算
# （_template_search_text 含字符串拼接 + 变量遍历，模板一多纯属浪费）。
# 以各模板检索文本拼接串为签名——任一模板内容变化（含本地库新增 / 改版本）都会改变签名自动失效，无需手动清缓存。
_TEMPLATE_INDEX = {}


def _get_template_hays(templates):
    sig = "\x00".join(_template_search_text(t) for t in templates)
    cached = _TEMPLATE_INDEX.get(sig)
    if cached is not None:
        return cached
    hays = [_template_search_text(t) for t in templates]
    if len(_TEMPLATE_INDEX) > 16:   # 简单限容，避免长期运行无限增长
        _TEMPLATE_INDEX.clear()
    _TEMPLATE_INDEX[sig] = hays
    return hays


def _select_relevant_templates(message, templates):
    """按需求挑相关模板，缩小规划 prompt。

    两层净化：① 剔除通用停用词；② 排除在过半模板都出现的高频词。
    再按重叠得分排名，只保留得分 ≥ 最高分一半的模板；完全匹配不到（模糊需求）回退全量。
    """
    if not templates:
        return templates
    q = _tokenize(message)
    if not q:
        return templates
    hays = _get_template_hays(templates)
    n = len(hays)
    # 排除在过半模板中都出现的高频通用词（兜底，防漏网停用词）
    q = [tok for tok in q if sum(1 for h in hays if tok in h) <= max(1, n * 0.6)]
    if not q:
        return templates
    scored = []
    for t, h in zip(templates, hays):
        score = sum(1 for tok in q if tok in h)
        scored.append((score, t))
    max_s = max((s for s, _ in scored), default=0)
    if max_s == 0:
        return templates
    keep = [t for s, t in scored if s > 0 and s >= max_s * 0.5]
    if not keep:
        return templates
    explicit = [t for t in templates if t.get("name", "").lower() in (message or "").lower()]
    keep = keep + [t for t in explicit if t not in keep]
    return keep


def _build_plan_prompt(message, templates, client=""):
    # 只把与本次需求相关的模板注入 prompt（全量仍会加载，但 prompt 不再膨胀）
    templates = _select_relevant_templates(message, templates)
    # 接口/端口类变量识别：命中这些词（含常见缩写，要求整体词边界避免误伤）即视为接口类。
    IFACE_RE = re.compile(
        r"\b(gi|ge|te|xe|eth|fe|hu|tw|po|port-channel|loopback|vlan|tunnel|interface|port|slot)\b",
        re.I,
    )

    ctx = []
    iface_vars = []  # 接口类变量的针对性强规则清单
    for t in templates:
        # 客户专属模板带 client 标记，让模型知道这个 name 只服务该客户
        ctag = {"client": t["client"]} if t.get("client") else {}
        if t["type"] == "normal":
            prompts = t.get("prompts") or {"vars": [], "repeats": [], "choices": []}
            ctx.append({
                "name": t["name"], "category": t["category"], "version": t["version"],
                "type": "normal", **ctag,
                "prompts": prompts,
            })
            for v in _iter_prompt_vars(prompts):
                dn = (v.get("display") or v.get("name", "")).lower()
                if IFACE_RE.search(dn):
                    iface_vars.append(f"{t['name']}::{v['name']}（显示名：{v.get('display','')}）")
        else:
            ctx.append({
                "name": t["name"], "category": t["category"], "version": t["version"],
                "type": t["type"], **ctag,
            })
    # 把接口类变量显式标注为字符串类型，抵消引擎 guess_type 因变量名含 "number" 误判 int 而误导模型。
    for t in ctx:
        if t["type"] == "normal":
            for v in _iter_prompt_vars(t["prompts"]):
                dn = (v.get("display") or v.get("name", "")).lower()
                if IFACE_RE.search(dn):
                    v["type"] = "str(interface)"

    contract = (
        '{"devices":[{"hostname":str,"category":str,"client":str?,"templates":[{"name":str,"version":str}],'
        '"stack_member":int?,"data":{'
        '"<normal模板名>":{"vars":{slug:值},"repeats":{rid:[{...}]},"choices":{cid:选项名}},'
        '"<stack_config模板名>":{"members":[{"number":int,"priority":int}],"renumber":[{"cur":int,"new":int}]},'
        '"<stack_picture模板名>":{}'
        '}}]}'
    )
    # 接口类变量针对性强规则：把命中变量逐条列出，要求值必须是含前缀的完整字符串。
    iface_rule = (
        "- 【接口变量清单】以下变量是接口/端口类，其值必须始终是\"含类型前缀的完整字符串\""
        "（如 \"g1/0/1\"、\"GigabitEthernet1/0/1\"、\"TenGigE0/0/1\"），"
        "即使变量名含 number/start/end 也绝不能是纯数字；"
        "禁止拆掉接口类型前缀、禁止改写、禁止补零：\n  "
        + ("\n  ".join(iface_vars) if iface_vars else "（无）")
        + "\n"
    )
    # 客户上下文：把「当前客户」与「已知客户清单」告知模型，让它能识别需求里提到的客户名
    client = str(client or "").strip()
    known = list_clients()
    if client:
        head = f"【当前客户上下文】{client}\n"
        if known:
            head += f"【已知客户清单】{'、'.join(known)}\n"
        head += (
            "客户规则：若用户需求里出现其他客户名（如「给华三的设备做配置」而当前客户是「华为」），"
            "必须在该 device 的 client 字段回填那个客户名；"
            "无法判断客户时留空（用当前客户）。\n"
        )
        # 当前客户有专属模板时，明确要求优先使用
        if any(t.get("client") == client for t in templates):
            head += "注意：带 client 字段的模板是该客户的专属版本，优先使用它们。\n"
    else:
        head = ""
    return (
        "你是一个网络配置生成规划器。根据下面的「可用模板及变量」，把用户需求翻译成符合契约的 JSON spec。\n\n"
        f"{head}"
        f"契约：\n{contract}\n\n"
        "可用模板：\n"
        f"{json.dumps(ctx, ensure_ascii=False, indent=2)}\n\n"
        "规则：\n"
        "- 只使用上面列出的模板 name；version 用给出的。\n"
        "- normal 模板：slug 必须与 prompts.vars[].name 完全一致；用户没说的变量填空字符串 \"\"；"
          "repeats 的项按 prompts.repeats[].vars 填；choices 填选项名。\n"
        "- 【data 位置铁律】data 只能放在 device 顶层、与 templates 同级，键名=模板 name"
          "（如 \"ap_management_port\"）；严禁把 data 塞进 templates 数组里的模板对象内部，"
          "也严禁在 data 下再套一层多余键名（如按 section 命名 {\"used_interface_configuration\":{...}}）。\n"
        "- 【原样保留】接口名/接口范围/IP/主机名/名称等网络标识必须原样输出：用户写 g1/0/1 就输出 "
          "\"g1/0/1\"，严禁拆掉接口类型前缀（g1/0/、Gi、Te、Eth 等）、严禁补零/改写/翻译；"
          "接口类变量的值必须是含前缀的完整字符串（如 \"g1/0/1\"），绝不能写成纯数字（如 1）。\n"
        + iface_rule +
        "- stack_config 模板：data 用 members（成员号与优先级）/ renumber（可选）结构，无需 vars。\n"
        "- stack_picture 模板：data 留空 {}，但设备需给 stack_member（成员数）以便嵌入连线图。\n"
        "- 只输出 JSON，不要任何解释或 markdown 代码块。\n\n"
        f"用户需求：{message}"
    )


def _body_vs_commands(cmd_text, body):
    """比对 AI 正文与命令原文，返回 (疑似尾部截断, 缺命令率, 缺命令样例)。

    判「截断」不靠行数（AI 归并 @repeat 会让行数本来就不齐），
    而看原文末尾 3 条命令还有没有落在正文里——被掐断时它们一条都找不到。
    """
    cmds = [ln.strip() for ln in cmd_text.splitlines()
            if ln.strip() and not ln.lstrip().startswith("!")]
    if not cmds:
        return False, 0.0, []

    def core(s):
        s = s.split("##")[0].strip()
        s = re.sub(r"\{\{[^}]*\}\}", " ", s)
        s = re.sub(r"<<[^>]*>>", " ", s)
        return re.sub(r"\s+", " ", s).strip()

    # 正文侧也用同一套 core() 归一化（去 ## 注释、去 {{}}、去 <<>>、压空白），
    # 两边口径一致才不会把「占位符写法不同」误判成缺命令。
    blines = [core(ln) for ln in body.splitlines() if ln.strip()]
    blines = [b for b in blines if b]

    def present(probe):
        return any(probe in ln or ln in probe for ln in blines)

    miss = [c for c in cmds if core(c) and not present(core(c))]
    # 判截断只看「原文末尾 3 条命令还在不在正文里」，比行数比准，也不受 @repeat 归并影响
    tail3 = [core(c) for c in cmds[-3:] if core(c)]
    tail_hit = sum(1 for p in tail3 if present(p))
    return tail_hit == 0, (len(miss) / float(len(cmds))), miss[:3]


def _unwrap_json_body(obj):
    """递归解开「body 字段又是一整段 JSON」的嵌套（模型偶发二次编码）。

    2026-10-05 用真实 Core switch 表自检时抓到：draft 返回的 body 是
    `{"body": "{\"body\": \"! --- ...\\n...\"}"}` 这种双层文本，
    而引擎校验照样 passed——只有靠巧合触发的表头软提醒才暴露。
    最多解 4 层，解不开就原样返回，交给 verify 的硬校验去拦。
    """
    for _ in range(4):
        if not isinstance(obj, dict):
            break
        v = obj.get("body")
        if not isinstance(v, str):
            break
        s = v.strip()
        # 只认「一层 JSON 外壳」的开头：截断场景连右括号都没有，不能要求 endswith("}")
        if not s.startswith("{"):
            break
        inner = _loose_body_obj(s)
        if not isinstance(inner, dict) or "body" not in inner:
            break
        obj = inner
    return obj


def _loose_body_obj(s):
    """从一段「可能截断 / 带杂文 / 换行写成字面 \\n」的 JSON 文本里尽力取出 body 对象。

    2026-10-05 真实 Core switch 表自检抓到最坏情况：模型输出超长被截断，
    JSON 连右括号都没写出来。四级宽容逐级降级，宁可糙也别把一坨 JSON 当正文。
    """
    try:
        o = json.loads(s)
        if isinstance(o, dict):
            return o
    except Exception:
        pass
    # raw_decode：容忍首尾夹着 markdown 之类的杂文
    try:
        o, _ = json.JSONDecoder().raw_decode(s)
        if isinstance(o, dict):
            return o
    except Exception:
        pass
    # 补尾部闭合：截断场景（缺 "} 或 ")
    for tail in ('"}', '"', '}'):
        try:
            o = json.loads(s + tail)
            if isinstance(o, dict) and isinstance(o.get("body"), str):
                return o
        except Exception:
            pass
    # 最后一级：手动平衡扫描找 body 值的闭合引号。
    # 不能用非贪婪正则——命令正文里常有裸双引号（如 remark "*GOS*"），
    # 非贪婪会在那里提前收尾，把整段正文几乎全丢（实测只捞回 1 行）。
    k = s.find('"body"')
    if k != -1:
        c2 = s.find(":", k)
        q0 = s.find('"', c2 + 1) if c2 != -1 else -1
        if q0 != -1:
            j = q0 + 1
            while j < len(s):
                ch = s[j]
                if ch == chr(92):          # 转义字符，跳掉它和它转义的那个
                    j += 2
                    continue
                if ch == '"':              # 找到闭合引号
                    break
                j += 1
            raw_frag = s[q0 + 1:min(j, len(s))]
            # 只有尾巴确实粘上了后面的键（截断特征）才退到最后一个完整行；
            # 否则会把正常命中的最后一行也砍掉——这个坑自己踩过一次。
            if re.search(r',\s*"(?:changes|uncovered)"\s*[:\[]?\s*$', raw_frag) or \
                    raw_frag.rstrip().endswith((",", "[")):
                frag = raw_frag.replace("\\n", "\n").replace('\\"', '"')
                cut = frag.rfind("\n")
                if cut != -1:
                    frag = frag[:cut]
            else:
                frag = raw_frag.replace("\\n", "\n").replace('\\"', '"')
            if frag.strip():
                return {"body": frag}
    return None


def _resolve_spec_names(spec):
    """交引擎前把 spec 里的唯一名解码回真名（引擎只认 (category,name)）。

    幂等：已是真名的原样透传，因此「未选客户」与旧流程行为完全一致。
    解码失败不在这里静默吞掉——真实模板缺失交给引擎报「模板未找到」，
    这样用户能看到引擎的原始报错，而不是被 app.py 改写过的假信息。
    """
    if not isinstance(spec, dict):
        return spec
    devices = spec.get("devices")
    if not isinstance(devices, list):
        return spec
    for dev in devices:
        if not isinstance(dev, dict):
            continue
        # device 级 client 仅用于结果回显与自检，不参与引擎查找
        for tpl in dev.get("templates") or []:
            if isinstance(tpl, dict) and isinstance(tpl.get("name"), str):
                real, _c = _split_unique_name(tpl["name"])
                tpl["name"] = real
    return spec


def spec_clients(spec):
    """提取 spec 里用到的客户集合（供结果区回显「本次使用客户模板」）。"""
    out = set()
    if not isinstance(spec, dict):
        return []
    for dev in spec.get("devices") or []:
        if not isinstance(dev, dict):
            continue
        c = str(dev.get("client") or "").strip()
        if c:
            out.add(c)
        for tpl in dev.get("templates") or []:
            if isinstance(tpl, dict):
                c2 = str(tpl.get("client") or "").strip()
                if c2:
                    out.add(c2)
    return sorted(out)


def _extract_json(text):
    text = text.strip()
    try:
        return json.loads(text)
    except Exception:
        pass
    s, e = text.find("{"), text.rfind("}")
    if s != -1 and e != -1 and e > s:
        try:
            return json.loads(text[s:e + 1])
        except Exception:
            pass
    # 模型输出超长被掐断时会连 `}` 都写不出来，补尾部闭合再试一次
    for tail in ('"}', '"', '}'):
        try:
            return json.loads(text + tail)
        except Exception:
            pass
    raise ValueError("无法从模型输出解析 JSON")


# ------------------------------- HTTP 服务 -------------------------------

class Handler(BaseHTTPRequestHandler):
    def _send_json(self, code, obj):
        body = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path

        if path in ("/", "/index.html"):
            try:
                with open(INDEX_PATH, "rb") as f:
                    data = f.read()
                self.send_response(200)
                self.send_header("Content-Type", "text/html; charset=utf-8")
                self.send_header("Content-Length", str(len(data)))
                self.end_headers()
                self.wfile.write(data)
            except FileNotFoundError:
                self._send_json(404, {"error": "index.html not found"})
            return

        if path == "/api/providers":
            qs = urllib.parse.parse_qs(parsed.query)
            mode = qs.get("mode", ["auto"])[0]
            self._send_json(200, {"mode": mode, "providers": orchestrator.list_providers(mode)})
            return

        if path == "/api/health":
            self._send_json(200, orchestrator.health())
            return

        if path == "/api/templates":
            try:
                self._send_json(200, {"templates": list_templates(_q_client(parsed.query))})
            except Exception as e:
                self._send_json(500, {"error": f"读取模板库失败: {e}"})
            return

        if path == "/api/template/entries":
            try:
                self._send_json(200, _template_entries(parsed.query))
            except Exception as e:
                self._send_json(500, {"error": f"读取模板库失败: {e}"})
            return

        if path == "/api/template/body":
            qs = urllib.parse.parse_qs(parsed.query)
            try:
                self._send_json(200, _template_body(
                    qs.get("source", ["local"])[0],
                    qs.get("category", [""])[0],
                    qs.get("name", [""])[0],
                    qs.get("version", [""])[0],
                ))
            except ValueError as e:
                self._send_json(400, {"error": str(e)})
            except Exception as e:
                self._send_json(500, {"error": str(e)})
            return

        if path == "/api/template/clients":
            self._send_json(200, {"clients": list_clients()})
            return

        if path == "/api/config":
            self._send_json(200, self._get_config())
            return

        if path == "/api/download":
            self._handle_download(parsed.query)
            return

        self._send_json(404, {"error": "not found"})

    def do_POST(self):
        parsed = urllib.parse.urlparse(self.path)
        length = int(self.headers.get("Content-Length", 0))
        raw = self.rfile.read(length) if length else b"{}"
        try:
            data = json.loads(raw.decode("utf-8") or "{}")
        except Exception as e:
            self._send_json(400, {"error": f"invalid json: {e}"})
            return

        if parsed.path == "/api/chat":
            self._handle_chat(data)
        elif parsed.path == "/api/skill/run":
            self._handle_skill_run(data)
        elif parsed.path == "/api/skill/plan":
            self._handle_skill_plan(data)
        elif parsed.path == "/api/skill/plan-draft":
            self._handle_skill_plan_draft(data)
        elif parsed.path == "/api/template/draft":
            self._handle_template_draft(data)
        elif parsed.path == "/api/template/verify":
            self._handle_template_verify(data)
        elif parsed.path == "/api/template/commit":
            self._handle_template_commit(data)
        elif parsed.path == "/api/template/save":
            self._handle_template_save(data)
        elif parsed.path == "/api/template/delete":
            self._handle_template_delete(data)
        elif parsed.path == "/api/template/fork":
            self._handle_template_fork(data)
        elif parsed.path == "/api/client/add":
            self._handle_client_add(data)
        elif parsed.path == "/api/config":
            self._handle_config(data)
        elif parsed.path == "/api/provider/test":
            self._handle_provider_test(data)
        else:
            self._send_json(404, {"error": "not found"})

    # ---- LLM 路由演示 ----
    def _handle_chat(self, data):
        mode = data.get("mode", "auto")
        message = data.get("message", "")
        if not message.strip():
            self._send_json(400, {"error": "message 不能为空"})
            return
        try:
            provider, content = orchestrator.complete(mode, data.get("provider"), message, json_mode=True)
        except PermissionError as e:
            self._send_json(403, {"error": str(e)})
            return
        except Exception as e:
            self._send_json(400, {"error": str(e)})
            return
        self._send_json(200, {"provider": provider.name, "kind": provider.kind, "content": content})

    # ---- CONFIG FORGE：直接执行（spec 已就绪） ----
    def _handle_skill_run(self, data):
        spec = data.get("spec")
        if not isinstance(spec, dict):
            self._send_json(400, {"error": "缺少 spec 对象"})
            return
        try:
            spec = _resolve_spec_names(spec)
            result = config_forge.run(spec)
        except Exception as e:
            self._send_json(502, {"error": f"CONFIG FORGE 执行失败: {e}"})
            return
        self._send_json(200, {"output": result.get("output"),
                              "clients": spec_clients(data.get("spec"))})

    # ---- CONFIG FORGE：一句话 → LLM 规划 → 执行 ----
    def _handle_skill_plan(self, data):
        message = data.get("message", "")
        mode = data.get("mode", "auto")
        if not message.strip():
            self._send_json(400, {"error": "message 不能为空"})
            return
        try:
            templates = list_templates(data.get("client", ""))
            prompt = _build_plan_prompt(message, templates, data.get("client", ""))
        except Exception as e:
            self._send_json(500, {"error": f"构造规划 prompt 失败: {e}"})
            return
        # 路由判定基于"用户原始需求"，而非构造后的规划 prompt。
        # 规划 prompt 是固定脚手架、恒含"配置"等词，若用其判定会误锁本地，
        # 导致自动模式永远走不了云端（演示断点）。
        try:
            provider, text = orchestrator.complete(
                mode, data.get("provider"), prompt, route_text=message, json_mode=True)
        except PermissionError as e:
            self._send_json(403, {"error": str(e)})
            return
        except Exception as e:
            self._send_json(400, {"error": f"LLM 规划失败: {e}"})
            return
        try:
            spec = _extract_json(text)
        except Exception as e:
            self._send_json(502, {"error": f"规划结果无法解析: {e}", "raw": text})
            return
        used = spec_clients(spec) or ([data.get("client", "")] if data.get("client") else [])
        try:
            run_spec = _resolve_spec_names(spec)
            result = config_forge.run(run_spec)
        except Exception as e:
            self._send_json(502, {"error": f"CONFIG FORGE 执行失败: {e}", "spec": spec})
            return
        self._send_json(200, {
            "provider": provider.name, "kind": provider.kind,
            "spec": spec, "output": result.get("output"), "clients": used,
        })

    # ---- CONFIG FORGE：一句话 → LLM 规划 → 返回草稿（不执行，等用户确认） ----
    def _handle_skill_plan_draft(self, data):
        message = data.get("message", "")
        mode = data.get("mode", "auto")
        if not message.strip():
            self._send_json(400, {"error": "message 不能为空"})
            return
        try:
            templates = list_templates(data.get("client", ""))
            prompt = _build_plan_prompt(message, templates, data.get("client", ""))
        except Exception as e:
            self._send_json(500, {"error": f"构造规划 prompt 失败: {e}"})
            return
        try:
            provider, text = orchestrator.complete(
                mode, data.get("provider"), prompt, route_text=message, json_mode=True)
        except PermissionError as e:
            self._send_json(403, {"error": str(e)})
            return
        except Exception as e:
            self._send_json(400, {"error": f"LLM 规划失败: {e}"})
            return
        try:
            spec = _extract_json(text)
        except Exception as e:
            self._send_json(502, {"error": f"规划结果无法解析: {e}", "raw": text})
            return
        # 只返回草稿，不调用 config_forge.run；交由前端弹窗让用户确认/修改后再执行。
        # 草稿里保留唯一名（前端要显示"这是华为那套"），执行时由 _resolve_spec_names 解码。
        self._send_json(200, {
            "provider": provider.name, "kind": provider.kind, "draft": spec,
            "clients": spec_clients(spec),
        })

    # ---- 模型设置（本机覆盖，含 API Key） ----
    def _get_config(self):
        p = cfg["providers"].get("cloud_openai", {})
        key = p.get("api_key", "")
        masked = ("********" + key[-4:]) if (key and not key.startswith("${")) else ""

        def loc(name):
            q = cfg["providers"].get(name, {})
            return {"base_url": q.get("base_url", ""), "model": q.get("model", "")}

        return {
            "cloud_openai": {
                "base_url": p.get("base_url", ""),
                "model": p.get("model", ""),
                "api_key_set": bool(key and not key.startswith("${")),
                "api_key_mask": masked,
            },
            "local_dsh": loc("local_dsh"),
            "local_ollama": loc("local_ollama"),
        }

    _CONFIG_PROVIDERS = ("cloud_openai", "local_dsh", "local_ollama")

    def _handle_config(self, data):
        """分块持久化：只更新本次提交的 provider 块，其他块保留。

        修复点：旧实现每次整体重写 config.local.json，云端保存会抹掉本地块（反之亦然）。
        现在先读旧文件，按 provider 名分块替换，再写回。空值不写（清空字段=回退 config.json 默认）。
        """
        local = {"providers": {}}
        if LOCAL_CONFIG_PATH.exists():
            try:
                local = json.loads(LOCAL_CONFIG_PATH.read_text(encoding="utf-8"))
            except Exception:
                local = {"providers": {}}
        local.setdefault("providers", {})
        for name in self._CONFIG_PROVIDERS:
            p = data.get("providers", {}).get(name)
            if p:
                local["providers"][name] = {
                    k: p[k] for k in ("api_key", "base_url", "model") if p.get(k)
                }
        LOCAL_CONFIG_PATH.write_text(json.dumps(local, ensure_ascii=False, indent=2), encoding="utf-8")
        rebuild_orchestrator()
        self._send_json(200, {"ok": True})

    def _handle_provider_test(self, data):
        """定点测试：直接 ping 指定 provider，绕过路由（云端/本地各自独立测）。"""
        name = (data.get("provider") or "").strip()
        p = orchestrator.providers.get(name)
        if p is None:
            self._send_json(400, {"error": f"未知 provider: {name}"})
            return
        t0 = time.monotonic()
        try:
            p.chat([{"role": "user", "content": "ping"}], max_tokens=1)
            self._send_json(200, {"ok": True, "provider": name, "kind": p.kind,
                                  "model": p.model, "ms": int((time.monotonic() - t0) * 1000)})
        except Exception as e:
            self._send_json(200, {"ok": False, "provider": name, "kind": p.kind,
                                  "model": p.model, "ms": int((time.monotonic() - t0) * 1000),
                                  "error": str(e)[:200]})

    def _handle_download(self, query):
        qs = urllib.parse.parse_qs(query)
        fname = qs.get("f", [""])[0]
        if not fname or "/" in fname or "\\" in fname or ".." in fname:
            self._send_json(400, {"error": "非法文件名"})
            return
        fp = CG_DIR / "output" / fname
        if not fp.exists():
            self._send_json(404, {"error": "文件不存在"})
            return
        data = fp.read_bytes()
        self.send_response(200)
        self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        self.send_header("Content-Disposition", f'attachment; filename="{fname}"')
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def log_message(self, *args):
        pass

    # ---- 模板工作室 ①：AI 生成模板草稿（不落盘） ----
    def _handle_template_draft(self, data):
        kind = data.get("kind") or "text"
        requirement = (data.get("requirement") or "").strip()
        category = (data.get("category") or "").strip()
        name = (data.get("name") or "").strip()
        label = (data.get("label") or "").strip()
        mode = data.get("mode", "auto")
        try:
            if kind == "xlsx":
                import base64
                import tempfile
                from pathlib import Path
                b64 = data.get("file_b64") or ""
                if not b64:
                    self._send_json(400, {"error": "缺少 xlsx 内容"})
                    return
                tmp = Path(tempfile.gettempdir()) / "tpl_studio_input.xlsx"
                tmp.write_bytes(base64.b64decode(b64))
                sections = _xlsx_sections(str(tmp))
                cmd_text = _sections_to_commands(sections)
            else:
                cmd_text = _norm_ws(data.get("content") or "").strip()
        except Exception as e:
            self._send_json(400, {"error": f"输入解析失败: {e}"})
            return
        if not cmd_text:
            self._send_json(400, {"error": "命令内容不能为空"})
            return
        if not category:
            self._send_json(400, {"error": "缺少分类（格式如 switch/access）"})
            return
        prompt = _build_template_prompt(cmd_text, requirement, category, name, label)
        try:
            provider, text = orchestrator.complete(
                mode, data.get("provider"), prompt,
                route_text=requirement or cmd_text[:200], json_mode=True)
        except Exception as e:
            self._send_json(400, {"error": f"LLM 调用失败: {e}"})
            return
        # 解包兜底（2026-10-05 用真实 Core switch 表实测抓到两坑）：
        #  ① 模型会把 JSON 再套一层，body 字段里是一段 `{"body": "...\n..."}` 转义文本；
        #  ② 输出超长被掐断，连 `}` 都没写出来，导致 _extract_json 抛异常，
        #     旧兜底把「整坨 JSON 原文」当成了模板正文（假绿，用户看到一坨花括号）。
        obj = {}
        try:
            obj = _extract_json(text)
        except Exception:
            try:
                obj = _loose_body_obj(text)
            except Exception:
                obj = {}
        if not isinstance(obj, dict) or not obj.get("body"):
            obj = _loose_body_obj(text) if text.strip().startswith("{") else {}
        obj = _unwrap_json_body(obj)
        body = obj.get("body") or ""
        changes = obj.get("changes") or []
        uncovered = obj.get("uncovered") or []
        if not body.strip():
            self._send_json(502, {"error": "未生成有效模板正文（模型输出既不是 JSON 也没抽到 body）",
                                  "raw": text[:300]})
            return
        if not body.strip():
            self._send_json(502, {"error": "未生成有效模板正文", "raw": text[:400]})
            return
        # 截断/漏命令护栏（真实 Core switch 表自检抓到）：兜底解包只能救回截断的前半段，
        # 后半段命令静默消失、引擎校验照样 passed。用「原文末尾命令还在不在正文里」判截断。
        try:
            tail_cut, gap_rate, gap_sample = _body_vs_commands(cmd_text, body)
            if tail_cut or gap_rate > 0.25:
                self._send_json(502, {
                    "error": "模型输出疑似被截断/漏命令（正文缺了约 %.0f%% 的命令，"
                             "如 %s）。长表容易触发，请点「重新生成」；若反复出现，"
                             "改用「按段落分批导入」或拆成小表。"
                             % (gap_rate * 100, (gap_sample or ["（见原文末尾）"])[0][:60]),
                })
                return
        except Exception:
            pass
        # 解包后还是 JSON 外壳 = 兜底也没救回来，直接拦住，别把一坨 JSON 推给用户
        bhead = body.strip()
        if bhead.startswith("{") and '"body"' in bhead[:120]:
            self._send_json(502, {
                "error": "模型返回的模板正文仍是嵌套 JSON 文本（自动解包后依旧如此），请点「重新生成」再来一次。",
                "raw": body[:300],
            })
            return
        self._send_json(200, {
            "provider": provider.name, "kind": provider.kind,
            "draft": {"body": body, "changes": changes, "uncovered": uncovered},
        })

    # ---- 模板工作室 ②：运行校验（硬校验=引擎，软提醒=规则扫描） ----
    def _handle_template_verify(self, data):
        import re
        body = (data.get("body") or "")
        if not body.strip():
            self._send_json(400, {"error": "模板正文不能为空"})
            return
        hard, soft = [], []
        engine = _cg_modules()[1]
        try:
            sections = _body_to_sections(body)
        except Exception as e:
            self._send_json(502, {"error": f"模板解析失败: {e}"})
            return
        # 草稿形态硬校验（2026-10-05 用真实 Core switch 表自检抓到）：模型偶发把 body
        # 二次编码成一整段 JSON 文本、或把换行全写成字面 \n，这两种引擎解析都照样过，
        # 只有靠巧合的软提醒才暴露。必须在硬校验层拦死，不能留「假绿」。
        bstrip = body.strip()
        if bstrip.startswith("{") and bstrip.endswith("}") and '"body"' in bstrip:
            hard.append({"level": "error",
                         "msg": "模板正文是一整段 JSON 原文（疑似模型又把 body 套了一层）。"
                                "请点「重新生成」再来一次；若是你手改粘进来的，把它拆成真实换行的配置行。"})
        if "\\n" in body and body.count("\n") == 0:
            hard.append({"level": "error",
                         "msg": "模板正文里的换行全被转义成了字面 \\n，引擎只会把它当一整行。"
                                "请生成真实换行的正文后再校验。"})
        if "! ---" not in body:
            hard.append({"level": "error", "msg": "未找到 `! --- 段名 ---` 段落分隔行"})
        if body.count("! @repeat") + body.count("! @choice") != body.count("! @end"):
            hard.append({"level": "error", "msg": "@repeat/@choice 与 @end 数量不配对（块缺 @end 会吞掉后续配置行）"})
        if "｛" in body or "｝" in body:
            soft.append({"level": "warn", "msg": "检测到全角大括号，占位符必须用半角 {{ }}"})
        # 非标准占位符：用户表格常见 <<x>> 写法，引擎不认，会导致变量抽不出来
        dbl = re.findall(r"<<([^>]+)>>", body)
        if dbl:
            # 注意：f-string 里要输出一对真实花括号得写两对，这里整体用 join 避开转义坑
            example = sorted(set(dbl))[0] if dbl else ""
            soft.append({
                "level": "warn",
                "msg": ("检测到非标准占位符 <<...>>（%d 处 / %d 个变量名，如 %r）。"
                        "模板语法只认半角双花括号，示例：%s。AI 包装时应把 <<x>> 转成 %s；"
                        "不转的话这些变量不会出现在变量列表里。"
                        % (len(dbl), len(set(dbl)), example, "{{" + example.strip() + "}}", "{{x}}")),
            })
        vars_found = re.findall(r"\{\{([^}]+)\}\}", body)
        # 变量名前后带空格：通常来自 `<< x >>` 写法直接替换，引擎会把空格算进变量名
        padded = sorted({v for v in vars_found if v != v.strip()})
        if padded:
            good = padded[0].strip()
            soft.append({
                "level": "warn",
                "msg": ("变量名前后带空格（%d 个，如 %r），通常是把 `<< x >>` 直接替换成 `%s` 造成的；"
                        "请改成不带首尾空格的 %s 形式。"
                        % (len(padded), padded[0], "{{" + padded[0] + "}}", "{{" + good + "}}")),
            })
        # 疑似表头行：复用 _is_header_row，保证与输入阶段「表头跳过」同一套判定，不写第二份规则
        for sec in sections:
            ls = sec.get("lines") or []
            if not ls:
                continue
            if not _is_header_row(ls[0]):
                continue
            cells = [str(c) for c in list(ls[0]) if isinstance(c, str)][:2]
            soft.append({
                "level": "warn",
                "msg": ("段落『%s』首行是疑似表头『%s』。输入阶段已自动跳过它；"
                        "如果正文里还能看到这一行，说明是手改时粘回来的，请删掉。"
                        % (sec.get("title"), " ## ".join(cells))),
            })
            break
        # 渲染冒烟：用示例值真跑一遍引擎
        sample = {v: ("g1/0/1" if ("nterface" in v or "port" in v.lower()) else "SAMPLE") for v in set(vars_found)}
        try:
            nodes = engine.parse_nodes(engine.split_lines(body))
            engine.render_nodes(nodes, sample, {"stack_member": 2})
            smoke_ok, smoke_err = True, ""
        except Exception as e:
            smoke_ok, smoke_err = False, str(e)[:300]
            hard.append({"level": "error", "msg": f"渲染冒烟失败（引擎用示例值渲染报错）: {smoke_err}"})
        self._send_json(200, {
            "hard": hard, "soft": soft,
            "variables": sorted(set(vars_found)),
            "sections": [s["title"] for s in sections],
            "smoke_ok": smoke_ok, "smoke_error": smoke_err,
            "passed": not [h for h in hard if h["level"] == "error"],
        })

    # ---- 模板工作室 ③：确认入库（仅显式调用时写库） ----
    def _handle_template_commit(self, data):
        body = (data.get("body") or "")
        category = (data.get("category") or "").strip().lstrip("/").lstrip("\\")
        # 模板名会直接拼进目录路径，净化掉路径分隔符，避免拼出双斜杠/越权路径
        name = re.sub(r"[\\/]+", "_", (data.get("name") or "")).strip("_")
        label = (data.get("label") or "").strip()
        version = (data.get("version") or "").strip()
        client = (data.get("client") or "").strip()
        if not body.strip():
            self._send_json(400, {"error": "模板正文不能为空"})
            return
        if not category:
            self._send_json(400, {"error": "缺少分类"})
            return
        try:
            sections = _body_to_sections(body)
        except Exception as e:
            self._send_json(502, {"error": f"模板解析失败: {e}"})
            return
        imp, lib = _studio_importer(), _studio_library()
        report = []

        def ask(q, default=""):
            # 分类/模板名/显示名/版本号：用户给了就用给的，没给就用 importer 的默认值（段名派生）
            if "模板分类" in q or q.strip().startswith("分类"):
                return category
            if "客户" in q:
                return client  # 客户归属：留空 = 未分组（import_sections 收到空串就不写 client 键）
            if "模板名" in q:
                if not name:
                    return default  # 按段名派生（推荐：一段一模板，互不覆盖）
                # 多段时若共用一个名字，后一段会把前一段覆盖成新版本 → 自动加段名后缀避免
                return f"{name}-{default}" if len(sections) > 1 else name
            if "显示名称" in q:
                if not label:
                    return default
                return f"{label}·{default}" if len(sections) > 1 else label
            if "版本号" in q:
                return version or default
            return default or ""

        def yes_no(q, default=True):
            return True  # 已由前端「我已检查」+ 显式提交把关

        io = {"ask": ask, "yes_no": yes_no, "print": lambda s: report.append(str(s))}
        try:
            written = imp.import_sections(sections, io, lib)
        except Exception as e:
            self._send_json(502, {"error": f"入库失败: {e}"})
            return
        self._send_json(200, {"report": "\n".join(report), "written": [str(d) for d, _ in written]})

    # ---- 模板工作室 ④：模板管理（改正文 → 存新版本 / 删模板） ----
    # 官方库（CG 自带）是只读基准，只能看；本地库（工作室自有）才能改能删。
    def _handle_template_save(self, data):
        try:
            self._send_json(200, _template_save(data))
        except ValueError as e:
            self._send_json(400, {"error": str(e)})
        except Exception as e:
            self._send_json(502, {"error": f"保存失败: {e}"})

    def _handle_template_delete(self, data):
        try:
            self._send_json(200, _template_delete(data))
        except ValueError as e:
            self._send_json(400, {"error": str(e)})
        except Exception as e:
            self._send_json(502, {"error": f"删除失败: {e}"})

    # ---- 模板工作室 ⑤：官方模板复制到某客户下（fork） ----
    def _handle_template_fork(self, data):
        try:
            self._send_json(200, _template_fork(data))
        except ValueError as e:
            self._send_json(400, {"error": str(e)})
        except Exception as e:
            self._send_json(502, {"error": f"复制失败: {e}"})

    # ---- 新建客户（写入 config.json 的 clients，手动维护口径） ----
    def _handle_client_add(self, data):
        try:
            self._send_json(200, _client_add(data))
        except ValueError as e:
            self._send_json(400, {"error": str(e)})
        except Exception as e:
            self._send_json(502, {"error": f"新建客户失败: {e}"})


# ---------------- 模板工作室（模板导入能力） ----------------

_LOCAL_LIBRARY = None


def _studio_library():
    """自有库模块：vendored 的 CG library.py，ROOT 落在 template_studio/library/ → 入库不碰 CG。"""
    global _LOCAL_LIBRARY
    if _LOCAL_LIBRARY is None:
        from template_studio import local_library as m
        _LOCAL_LIBRARY = m
    return _LOCAL_LIBRARY


def _studio_importer():
    """CG 的 importer.py 只读 import 复用（不改 CG）；经 library= 参数把入库指向我们自己的库。"""
    _cg_modules()  # 确保 engine 可 import
    from template_studio import importer as m
    return m


# ---------------- 模板工作室：模板管理（改 / 删） ----------------
# 双库口径（2026-10-06 定）：官方库 = CG 自带，只读基准（看得见，不能改不能删）；
# 本地库 = 工作室自有（template_studio/library/），可改可删。改完按新版本 v+1 存，旧版本保留。

def _library_by_source(source):
    """按 source 定位库模块：official=CG 官方（只读），local=工作室自有。"""
    if source == "official":
        return _cg_modules()[0]
    if source == "local":
        return _studio_library()
    raise ValueError("source 必须是 official 或 local")


def _safe_name(s):
    """分类/模板名会直接拼成盘路径：清掉路径分隔符与首尾杂符，防拼出双斜杠/越权路径。"""
    return re.sub(r"[\\/]+", "_", str(s or "")).strip("._ ")


def _find_by_category_name(library, category, name):
    """按（分类, 模板名）定位条目；前端传来的写法可能与库里不同，归一化后再兜底找一次。"""
    idx = library.load_index()
    entry = library.find_entry(idx, category, name)
    if entry:
        return idx, entry
    cat, nm = _safe_name(category), _safe_name(name)
    for e in idx["templates"]:
        if _safe_name(e.get("category")) == cat and _safe_name(e.get("name")) == nm:
            return idx, e
    return idx, None


def _template_item(lib, e, editable):
    """库条目 → 前端可 presentations 的清单项（含版本列表）。单条坏数据不许拖垮整个列表。"""
    cat, nm = e.get("category", ""), e.get("name", "")
    versions = [v.get("version") for v in e.get("versions", [])]
    item = {
        "category": cat, "name": nm, "label": e.get("label", ""),
        "version": lib.active_version(e) if versions else "",
        "type": "normal", "editable": editable, "versions": versions,
        "client": _studio_client(e),   # 客户归属，供前端分组显示与筛选
        "path": f"{cat}/{nm}" + (f"/{lib.active_version(e)}" if versions else ""),
    }
    try:
        schema = lib.load_schema(e)
        item["type"] = schema.get("type", "normal")
    except Exception as ex:
        item["error"] = str(ex)[:200]
    return item


def _template_entries(query=""):
    """模板清单：官方库（只读）+ 本地库（可改可删），本地库读失败也只是这一块为空。"""
    official = []
    try:
        library = _cg_modules()[0]
        for e in library.load_index()["templates"]:
            official.append(_template_item(library, e, editable=False))
    except Exception as e:
        official = [{"category": "", "name": "", "label": "", "version": "",
                     "type": "error", "editable": False, "versions": [],
                     "error": f"读取官方库失败: {e}"}]
    local = []
    try:
        lib = _studio_library()
        for e in lib.load_index()["templates"]:
            local.append(_template_item(lib, e, editable=True))
    except Exception as e:
        local = [{"category": "", "name": "", "label": "", "version": "",
                  "type": "error", "editable": False, "versions": [],
                  "error": f"读取本地库失败: {e}"}]
    # 按客户过滤（模板管理区选了客户就只看该客户 + 未分组；不选=全看）
    client = _q_client(query)
    if client:
        local = [x for x in local
                 if not x.get("client") or x.get("client") == client]
    return {"official": official, "local": local, "client": client}


def _template_body(source, category, name, version=""):
    """按 source/category/name 取模板正文；官方库返回 editable=false（只能看）。"""
    lib = _library_by_source(source)
    _idx, entry = _find_by_category_name(lib, category, name)
    if not entry:
        raise ValueError(f"{source} 库里没有模板 {category}/{name}")
    v = version or lib.active_version(entry)
    p = lib.version_dir(entry, v) / "template.txt"
    if not p.exists():
        raise FileNotFoundError(f"模板正文缺失: {p}")
    return {
        "source": source, "category": entry["category"], "name": entry["name"],
        "label": entry.get("label", ""), "version": v,
        "editable": source == "local",
        "versions": [x.get("version") for x in entry.get("versions", [])],
        "body": p.read_text(encoding="utf-8"),
    }


def _template_save(data):
    """改本地库模板正文 → 复用 importer 存为新版本（bump_version 给 v+1，旧版本保留）。"""
    source = (data.get("source") or "local").strip()
    category = (data.get("category") or "").strip()
    name = (data.get("name") or "").strip()
    body = data.get("body") or ""
    label = (data.get("label") or "").strip()
    if source != "local":
        raise ValueError("官方库（CG 自带）是只读基准，改不了；请先在上方 ①-③ 另建一个新模板")
    if not body.strip():
        raise ValueError("模板正文不能为空")
    lib = _library_by_source(source)
    _idx, entry = _find_by_category_name(lib, category, name)
    if not entry:
        raise ValueError(f"本地库里没有 {category}/{name}，无从改起；请先点①生成、③入库")
    try:
        sections = _body_to_sections(body)
    except Exception as e:
        raise RuntimeError(f"模板解析失败: {e}")
    report = []
    new_ver = lib.bump_version(entry["versions"])  # 统一口径：永远存新版本，旧版本留着可回退

    def ask(q, default=""):
        # 沿用 commit 的关键词识别：ask 的 q 来自 importer，顺序不能动（"模板分类" 优先于 "分类"）
        if "模板分类" in q or q.strip().startswith("分类"):
            return entry["category"]
        if "模板名" in q:
            return entry["name"]
        if "显示名称" in q:
            return label or entry.get("label", "")
        if "版本号" in q:
            return new_ver
        return default or ""

    io = {"ask": ask, "yes_no": lambda q, default=True: True,
          "print": lambda s: report.append(str(s))}
    written = _studio_importer().import_sections(sections, io, lib)
    return {
        "report": "\n".join(report), "category": entry["category"], "name": entry["name"],
        "label": label or entry.get("label", ""), "version": new_ver,
        "changed": bool(written), "written": [str(d) for d, _ in written],
    }


# ---- 新建客户（写入 config.json 的 clients，手动维护口径；函数体在类外的模块区） ----


def _client_add(data):
    """新建客户：把名字写进 config.json 的 clients 数组（手动维护，不碰任何模板）。"""
    name = _safe_name(data.get("client"))
    if not name:
        raise ValueError("请填写客户名")
    try:
        cfg = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    except Exception:
        cfg = {}
    clients = cfg.get("clients")
    if not isinstance(clients, list):
        clients = []
    if name in [str(c).strip() for c in clients]:
        return {"clients": clients, "added": False, "message": f"客户「{name}」已存在"}
    clients.append(name)
    cfg["clients"] = clients
    CONFIG_PATH.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")
    return {"clients": clients, "added": True, "client": name}


def _template_fork(data):
    """把官方模板整目录复制到本地库的某个客户下（客户定制入口）。

    - 官方库物理不动（只读），复制进的是本地库 template_studio/library/
    - 必须递归复制整个版本目录：模板含 schema.json + template.txt + styles.json，
      个别还带 images/ 子目录（如 stack_picture_selection），漏了会丢图
    - 复制后落本地库、与官方同名 (category,name)，并打 client 标记 → 覆盖该客户视角
    - 该客户已有同名条目时不静默覆盖：返回 needs_confirm，让前端提示"改存为新版本"
    """
    client = _safe_name(data.get("client"))
    if not client:
        raise ValueError("请指定客户（client）")
    # 注意：category 合法含 "/"（如 switch/access 是目录层级），不能像 name 那样被 _safe_name
    # 把斜杠换成下划线——那会拼出错的层级、find_entry 必然落空。这里只做去空白 + 越权净化。
    category = re.sub(r"[\\]+", "_", str(data.get("category") or "")).strip("/._ ")
    raw_name = str(data.get("name") or "")
    if CLIENT_SEP in raw_name:
        raw_name = _split_unique_name(raw_name)[0]
    name = _safe_name(raw_name)
    if not category or not name:
        raise ValueError("请指定要复制的官方模板（category + name）")

    official = _cg_modules()[0]
    oidx = official.load_index()
    oentry = official.find_entry(oidx, category, name)
    if not oentry:
        raise ValueError(f"官方库里没有 {category}/{name}")

    from_version = (data.get("from_version") or "").strip()
    if not from_version:
        from_version = official.active_version(oentry)

    local = _studio_library()
    lidx = local.load_index()
    # 客户维度必须参与查找：本地库允许「同一 (category,name) 存在多个客户版本」，
    # find_entry 只认 (category,name)，会把"华为已有"误判成"华三已有"。
    same = [e for e in lidx.get("templates", [])
            if e.get("category") == category and e.get("name") == name
            and _studio_client(e) == client]
    existed = same[0] if same else None

    src_dir = Path(official.version_dir(oentry, from_version)).resolve()
    if not src_dir.exists():
        raise ValueError(f"官方模板版本目录不存在: {from_version}")

    # 复制整个版本目录（递归，含 images/ 等子目录）
    dst_dir = Path(local.version_dir({"category": category, "name": name}, from_version)).resolve()
    root = Path(local.LIBRARY_DIR).resolve()
    if dst_dir != root and root not in dst_dir.parents:
        raise ValueError("拒绝写入越权路径")
    dst_dir.parent.mkdir(parents=True, exist_ok=True)
    if existed is None or data.get("overwrite"):
        if dst_dir.exists():
            shutil.rmtree(str(dst_dir))
        shutil.copytree(str(src_dir), str(dst_dir))

    # 校验复制完整性：文件数 + 总字节一致（漏文件/截断都能抓到）
    def _measure(d):
        files = [p for p in Path(d).rglob("*") if p.is_file()]
        return len(files), sum(p.stat().st_size for p in files)

    n_src, b_src = _measure(src_dir)
    n_dst, b_dst = _measure(dst_dir)

    if existed is not None and not data.get("overwrite"):
        # 该客户已有同名条目：不覆盖，让用户决定是「改存为新版本」还是别的
        entry = existed
        return {
            "needs_confirm": True,
            "message": f"客户「{client}」下已有 {category}/{name}"
                        f"（当前版本 {local.active_version(entry)}）。"
                        f"是否改存为新版本？官方原版不会被改动。",
            "category": category, "name": name, "client": client,
            "current_version": local.active_version(existed),
            "copied": False,
        }

    if (n_src, b_src) != (n_dst, b_dst):
        raise ValueError(
            f"复制完整性校验失败：官方 {n_src} 个文件/{b_src} 字节，"
            f"本地 {n_dst} 个文件/{b_dst} 字节；已清理"
        )

    # 写本地库条目：该客户已有同名条目 -> 沿用它；否则追加一条新条目（不同客户各自一条）
    versions = [{"version": from_version, "date": "", "hash": ""}]
    if existed is None:
        lidx.setdefault("templates", []).append({
            "category": category, "name": name, "label": oentry.get("label", name),
            "client": client, "active": from_version, "versions": versions,
        })
    else:
        existed["client"] = client
        existed["active"] = from_version
        if not any(v.get("version") == from_version for v in existed.get("versions", [])):
            existed.setdefault("versions", []).append(
                {"version": from_version, "date": "", "hash": ""})
    local.save_index(lidx)

    return {
        "needs_confirm": False, "copied": True,
        "category": category, "name": name, "client": client,
        "version": from_version, "files": n_dst,
    }


def _template_delete(data):
    """删本地库模板：清掉全部版本目录 + 从 index 摘掉条目（官方库一律拒绝）。

    带 client 时只删该客户的那一条：同一 (category,name) 可能存在多个客户版本共存，
    物理目录同名共享，所以只有在「该 (category,name) 已无任何其他客户条目」时才删物理目录，
    否则会连带删掉别的客户的模板。
    """
    source = (data.get("source") or "local").strip()
    if source != "local":
        raise ValueError("官方库（CG 自带）是只读基准，不能删除")
    client = _studio_client({"client": data.get("client")})
    category = (data.get("category") or "").strip()
    name = (data.get("name") or "").strip()
    if CLIENT_SEP in name:
        name = _split_unique_name(name)[0]
    lib = _library_by_source(source)
    idx = lib.load_index()
    if client:
        matches = [e for e in idx.get("templates", [])
                   if e.get("category") == category and e.get("name") == name
                   and _studio_client(e) == client]
        if not matches:
            raise ValueError(f"客户「{client}」下没有 {category}/{name}，可能已被删掉")
        entry = matches[0]
    else:
        idx, entry = _find_by_category_name(lib, category, name)
        if not entry:
            raise ValueError(f"本地库里没有 {category}/{name}，可能已被删掉")
    root = Path(lib.LIBRARY_DIR).resolve()
    removed, kept = [], []

    # 还有没有其他客户在用同一个 (category,name)？有就别动物理目录。
    siblings = [e for e in idx.get("templates", [])
                if e is not entry and e.get("category") == category and e.get("name") == name]

    if not siblings:
        for v in entry.get("versions", []):
            d = lib.version_dir(entry, v["version"])
            d_res = Path(d).resolve()
            # 越权护栏：只允许删 LIBRARY_DIR 之下的目录
            if d_res != root and root not in d_res.parents:
                raise ValueError(f"拒绝删除越权路径: {d}")
            if not d_res.exists():
                kept.append(f"{v['version']}（目录不存在）")
                continue
            shutil.rmtree(str(d_res))
            removed.append(f'{Path(entry["category"]) / entry["name"] / v["version"]}'.replace("\\", "/"))
    else:
        kept.append(f"物理目录与其他 {len(siblings)} 个客户版本共享，未删除（已只移除本客户的登记）")

    # 版本目录清完再收空壳：模板目录与（若空的）分类目录一起删掉，别在盘上留空壳
    def _empty_and_inside(d):
        d = Path(d)
        if not d.exists():
            return False
        if d.resolve() != root and root not in d.resolve().parents:
            return False
        return not any(d.iterdir())

    if not siblings:
        versions = entry.get("versions", [])
        base = lib.version_dir(entry, versions[0]["version"]).parent if versions else None
        if base:
            # 从深到浅逐级收空壳：版本目录删完后，模板目录可能还残留下一层空目录，
            # 只试两级会漏（曾留下 library/switch 空壳）。逐层向上直到收不动或触到根。
            cur = Path(base)
            while cur and cur.resolve() != root and root in cur.resolve().parents:
                if not _empty_and_inside(cur):
                    break
                shutil.rmtree(str(cur.resolve()))
                removed.append(str(cur.relative_to(root)).replace("\\", "/"))
                cur = cur.parent
            # 兜底：摘掉登记后再扫一遍，若该 (category,name) 已无任何登记但目录还在，补清
            still = [e for e in idx.get("templates", [])
                     if e.get("category") == category and e.get("name") == name]
            if not still and base:
                cur = Path(base)
                while cur and cur.resolve() != root and root in cur.resolve().parents:
                    if not _empty_and_inside(cur):
                        break
                    shutil.rmtree(str(cur.resolve()))
                    cur = cur.parent

    idx["templates"] = [e for e in idx["templates"] if e is not entry]
    lib.save_index(idx)
    return {
        "category": entry["category"], "name": entry["name"], "client": _studio_client(entry),
        "label": entry.get("label", ""), "removed": removed, "missing": kept,
    }


def _read_spec_doc():
    """模板编写规范（CG 官方文档，仅只读读取）"""
    p = CG_DIR / "examples" / "模板编写规范.md"
    try:
        return p.read_text(encoding="utf-8")
    except Exception:
        return "（规范文档缺失，按既有格式约定处理）"


def _is_header_row(row):
    """保守表头识别：一段的首行同时含 ` ## ` 与 ` | `（多列表头特征）才判为表头。

    用户口径（2026-10-05 定）：宁可漏判也不能误跳真命令，所以条件取严。
    """
    cells = [str(c) for c in row if isinstance(c, str)][:2]
    if not cells:
        return False
    head = " ## ".join(cells)
    # 注释行（! 开头）不可能是表头，否则会被「表头说明注释」自己误判成表头
    if any(c.lstrip().startswith("!") for c in cells):
        return False
    return " ## " in head and " | " in head


def _sections_to_commands(sections):
    """sections 摊平成 `命令 ## 注释` 文本，供 AI 包装；段落首行为表头时跳过并留一行说明注释。"""
    lines = []
    for sec in sections:
        lines.append(f"! --- {sec['title']} ---")
        rows = list(sec.get("lines") or [])
        if rows and _is_header_row(rows[0]):
            lines.append(f"! [输入解析] 已自动跳过本段表头行: {' ## '.join(str(c) for c in list(rows[0])[:2])}")
            rows = rows[1:]
        for t, c in rows:
            if str(t).strip():
                lines.append(f"{t} ## {c}" if c else t)
    return "\n".join(lines)


def _body_to_sections(body):
    """模板正文 → sections（复用 CG 解析器，只读）"""
    import tempfile
    from pathlib import Path
    tmp = Path(tempfile.gettempdir()) / "tpl_studio_e2e_body.txt"
    tmp.write_text(body, encoding="utf-8")
    return _studio_importer().read_txt_sections(str(tmp))


def _xlsx_sections(path):
    """xlsx 需 openpyxl：沿用 config_forge 的子进程套路，用包内 runtime/ 便携 Python 解析"""
    import json
    import subprocess
    venv = BASE_DIR.parent / "runtime" / "python.exe"
    if not venv.exists():
        raise RuntimeError("未找到包内 runtime/python.exe，无法解析 xlsx")
    probe = BASE_DIR / "template_studio" / "_xlsx_probe.py"
    if not probe.exists():  # 降级：无探针时按纯文本读，再归一一次兜底
        return _norm_deep(_studio_importer().read_txt_sections(path))
    proc = subprocess.run(
        [str(venv), str(probe), str(path), str(CG_DIR)],
        capture_output=True, text=True, encoding="utf-8", timeout=120,
    )
    if proc.returncode != 0:
        err = (proc.stderr or "").strip()
        # 编码类崩溃的原因往往看不懂，补一句人话定位
        hint = ""
        if "codec can't encode" in err:
            hint = "（多为 Excel/网页粘贴带入的不换行空格 NBSP 等不可见字符，已在探针侧归一）"
        raise RuntimeError(f"xlsx 解析失败: {err[:400]}{hint}")
    return json.loads(proc.stdout.strip())


def _build_template_prompt(command_text, requirement, category, name, label):
    """AI 只做语法包装的强约束 prompt（规范全文 + 不得创作命令）"""
    return (
        "你是一个【配置模板包装器】，不是配置创作者。你只做语法包装，绝不创作/改写配置命令。\n\n"
        "必须严格遵循下面的《模板编写规范》：\n"
        f"{_read_spec_doc()}\n\n"
        "硬性约束（违反即判失败）：\n"
        "- 你输出的每一行配置命令都必须逐字来自【命令原文】；"
        "严禁新增、改写、省略、合并命令，也严禁凭需求描述推测补充命令。\n"
        "- 你只能做这些语法层处理：用 `! --- 段名 ---` 分段；"
        "把命令中会变化的值抽成 `{{变量}}`（变量名要人能看懂，如 {{VLAN id}}，不要 {{v1}}）；"
        "把多行同类命令归并成 `@repeat ... @end`；把互斥方案套成 `@choice ... @option ... @end`；"
        "把命令原文里既有的双尖括号变量标记 `<<x>>` 改写成 `{{x}}`（占位符写法转换，不算改写命令，"
        "必须成对替换且不能漏）；注意原文可能写成 `<< SNMP location >>` 这种带内边距空格的形式，"
        "改写时**去掉变量名前后的空格**（得到 `{{SNMP location}}`），变量名内部保留原文的空格/连字符。\n"
        "- 【需求描述】提到但【命令原文】没有对应命令的功能，不要自己写命令，只写进返回结果的 uncovered 字段。\n"
        "- 注释沿用命令原文里的注释（` ## ` 之后的内容），不要自造注释。\n"
        "- 【body 写法铁律·违反即判失败】`body` 字段里必须是真实回车换行的配置行文本，直接把配置行填进去；"
        "严禁再把整个 JSON 套一层当作 body、严禁 body 以 `{` 开头、严禁在 body 里写 `\\n` 这类转义序列"
        "（换行就是真正的回车，一行一条命令），也严禁 body 里出现 markdown 代码块符号。\n"
        "- 输出格式：只输出一个 JSON 对象，不要解释、不要 markdown 代码块，形如：\n"
        '  {"body": "模板正文（多行）", "changes": ["做了什么"], "uncovered": ["未覆盖项"]}\n\n'
        f"【模板信息】分类：{category}｜指定模板名：{name or '（不指定则按段名自动生成）'}"
        f"｜指定显示名：{label or '（不指定则按段名自动生成）'}\n"
        f"【命令原文】\n{command_text}\n"
        f"【需求描述】\n{requirement or '（无）'}\n"
        "请输出包装后的模板正文 JSON。"
    )


def main():
    # 默认 18000：本机 Windows 动态端口范围是 1024~15000（netsh int ipv4 show dynamicport tcp），
    # 8000 落在其中，会被任意进程的出站连接随机占用（WinError 10013/10048，概率性复现）。
    base_port = int(os.environ.get("PORT", "18000"))
    server = None
    last_err = None
    for port in range(base_port, base_port + 10):
        try:
            server = ThreadingHTTPServer(("0.0.0.0", port), Handler)
            break
        except OSError as e:
            last_err = e
    if server is None:
        print(f"[ERROR] 端口 {base_port}~{base_port + 9} 全部无法绑定：{last_err}")
        print("可能原因：端口被其他程序占用，或落在 Windows 动态端口范围内被出站连接随机占用。")
        print("办法一：换段端口启动，例如  start.bat 23000")
        print("办法二（根治）：以管理员身份执行下面的命令恢复系统默认动态端口范围（1024~15000 之外不再被随机占用）：")
        print("  netsh int ipv4 set dynamicport tcp start=49152 num=16384")
        return
    port = server.server_address[1]
    if port != base_port:
        print(f"[提示] 端口 {base_port} 被占用，已自动改用 {port}。")
    print(f"LLM Orchestrator 已启动: http://localhost:{port}  (Ctrl+C 退出)")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        server.shutdown()


if __name__ == "__main__":
    main()
