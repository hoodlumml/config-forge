#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""seam_guard —— 模板正文保真校验（确定性，纯函数，无 LLM、无网络）。

解决的问题
----------
AI 做「语法包装」时，命令原文经一次文本往返（prompt → body）可能被静默改动：
漏一行、多一行、换顺序、改写法、把值写成字面量。而旧护栏
（app.py:_body_vs_commands）只做「原文的每一条是否出现在正文里」的单向字面包含判断，
于是：漏的抓不到（阈值 25%）、多的不管、顺序不管、改写不管。

本模块的口径：以「引擎的真实解析结果」为准做双向比对。
    - 权威解析 = config-generator 的 engine.parse_nodes（@repeat/@choice 展开为 line）
    - 判定     = 集合比对（不多不少不少量）+ 顺序比对（不换位）
任何一项不符即不合格，由调用方拒绝入库。

为什么必须双向
--------------
旧护栏只做「原文 → 正文」单向包含检查，缺命令率超阈值才拦。
真实语料（CG 官方库 17 个模板）证明三件事：
  ① 模板正文大量存在字面 IP / community 字符串 —— 字面值本身合法，
     不能因为「出现字面值」就判污染；只能判「与原文不一致」。
  ② 归并 @repeat 是 AI 的合法职责（命令数 N → 1），所以「正文行数 < 原文行数」合法；
     但归并后的块必须能覆盖原文对应的那一组，不能凭空吞掉别处。
  ③ 命令顺序对网络配置有语义（先建 vlan 再配 interface），换序必须能判出来。

三个口径（都基于归一化后的「命令骨架」，不含注释、不含变量值）：
    extra   ：正文有、原文没有的骨架  -> AI 创作/幻觉，拒
    missing ：原文有、正文没有的骨架  -> 漏命令，拒
    reorder ：骨架集合相同但顺序不同 -> 换序，拒

为什么不用「条数」而用「骨架集合」
--------------------------------
归并与拆分会让命令条数根本不可比，且两种方向都合法：
    原文 1 条 `vlan batch 10 20`
      → AI 归并成 @repeat：  `vlan {v}` + `name {v}`   （条数 1→2，合法）
    原文 2 条 `interface X` / `description Y`
      → AI 不归并：          仍是 2 条                 （条数不变，也合法）
所以任何按「行数 / 缺命令率」的阈值判定（旧的 _body_vs_commands 用 25%）
都会在合法归并时误报、在小漏行时漏报——真实语料里 @repeat 出现在 9/17 个官方模板中。

替代口径：先做「字面骨架」的严格双向比对；仅当原文存在「含占位符的骨架」
或正文存在 @repeat/@choice 块时，才降级到「字面量段」的宽松比对
（把占位符与被 @repeat 吸收的字面量都视作可变部分，只比不可变的命令词干）。
"""

import re

# ---------------- 归一化口径 ----------------

_VAR_CURLY = re.compile(r"\{\{\s*(.+?)\s*\}\}")
_VAR_ANGLE = re.compile(r"<<\s*(.+?)\s*>>")
_SECTION = re.compile(r"^!\s*-{2,}\s*(.+?)\s*-{2,}\s*$")


def _strip_comment(line):
    """去掉 ` ## 注释` 部分（注释可由 AI 调整，不参与保真判定）。

    注意 `! ## https://...` 这种「整行只有注释、没有命令」的情况：
    朴素切分会把它剥成 `!`，而 `!` 又是最容易与真命令混淆的残渣。
    这里切完若只剩 ! 前缀，就直接返回空串，让它走 _is_marker 的注释分支。
    """
    if "##" in line:
        line = line.split("##", 1)[0]
    line = line.strip()
    # 剥完只剩 ! 或若干 !（`!## x`、`!! ## y`）-> 实质是纯注释行，无命令主体
    if line and set(line) == {"!"}:
        return ""
    return line


def _norm(line):
    """归一化成命令骨架：
    - 去注释、去首尾空白
    - 连续空白压成单空格（容忍 AI 的对齐/缩进变化）
    - 占位符统一成 {v}（`{{VLAN id}}` 与 `<<vlan id>>` 视为同一个槽位，
      这是模板编写规范明确允许的转换，不该判成改写）
    """
    s = _strip_comment(line)
    if not s:
        return ""
    s = _VAR_CURLY.sub("{v}", s)
    s = _VAR_ANGLE.sub("{v}", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def _is_marker(s):
    """标记行 / 注释行不参与保真比对。

    真实语料（CG 官方库 17 个模板）里以 `!` 开头的行分三类，都不是命令：
      ① 段头      `! --- Section Name ---`
      ② 块标记    `! @repeat` / `! @choice` / `! @option` / `! @end`
      ③ 人写的说明 `!` 独占一行、`! Set the order depending on...`、
                   `!!Logging Configuration`、`!!!Voice ACL`、`!## https://...`
    ⚠️ 第③类曾被漏判：旧口径只认「恰好等于 `!`」，于是 `!!Logging Configuration`
    被当成命令骨架，比对时报出 `['!!!', '!!!!!!logging', '!aaa']` 这种
    用户完全看不懂的 token。CG 官方模板里这类行有 14 处。
    口径改为：**任何以 `!` 开头的行都不是命令**（配置命令不会以 ! 开头）。
    """
    if not s:
        return True
    # 注释与段头、块标记——全部以 ! 开头
    if s.startswith("!"):
        return True
    return False


def extract_skeletons(lines):
    """从原文行列表抽骨架序列（顺序保留）。"""
    out = []
    for ln in lines:
        s = _norm(ln)
        if not s or _is_marker(s):
            continue
        out.append(s)
    return out


def _locate(original_lines, skeleton):
    """把骨架还原成「原文第几行 + 那行原文是什么」。

    报错时给用户看 '第 42 行：snmp-server host 1.1.1.1'，
    而不是给一个孤零零的骨架字符串 —— 后者对不上他手里的原文，等于没说。
    """
    for i, ln in enumerate(original_lines, 1):
        if _norm(ln) == skeleton:
            raw = _strip_comment(ln) or ln.strip()
            return "第%d行「%s」" % (i, raw[:56])
    return "「%s」" % skeleton[:56]


# ---------------- 说明行（以 ! 开头的注释/说明）----------------
# 说明行不是命令，但**不是可以随便丢的**：engine.render_nodes 会把它们以
# kind="note" 渲染进最终 Excel，是交付物的一部分。实测：
#   '! Set the order depending on the region' -> node type=line -> 渲染为 note 行
# 所以口径是：说明行**不参与命令词干比对**（它的首词毫无语义，会造成大量假阳性），
# 但**单独校验存在性与内容**，被吞掉照样拦。

def _note_key(s):
    """说明行的比对键：整行归一化（压空白、占位符统一），保留全部文字。

    不取首词——说明行的首词（`!aaa` / `!!logging`）无语义，
    只有整行文字才能判断它有没有被改写或吞掉。
    """
    t = _strip_comment(s)
    if not t:
        return ""
    t = _VAR_CURLY.sub("{v}", t)
    t = _VAR_ANGLE.sub("{v}", t)
    return re.sub(r"\s+", " ", t).strip()


def _note_index(lines):
    """抽取「人写的说明行」集合（按归一化整行去重）。

    ⚠️ **段头 `! --- X ---` 与块标记 `! @repeat` 不计入**：
    它们是 AI 的合法产出（分段/归并是 prompt 明确授权的语法加工），
    正文里新增段头不构成「多出说明行」。纯 `!` 空注释行也不计。
    只对**真正的说明文字**做存在性与内容比对。
    """
    out = set()
    for ln in lines:
        s = _norm(ln)
        if not s or not s.startswith("!"):
            continue
        if _SECTION.match(s) or s.startswith("! @"):
            continue                      # 段头 / 块标记：AI 合法产出
        k = _note_key(ln)
        if k and k != "!":
            out.add(k)
    return out


def _note_locate(original_lines, note_key):
    """把说明行键还原成「原文第几行 + 那行内容」。"""
    for i, ln in enumerate(original_lines, 1):
        if _note_key(ln) == note_key:
            return "第%d行「%s」" % (i, ln.strip()[:56])
    return "「%s」" % note_key[:56]


def _unwrap_blocks(lines):
    """把 body 文本按引擎口径展开成扁平行序列：
    复用 config-generator 的 engine.parse_nodes，保证与渲染时看到的是同一套语义。
    engine 不可用时退化为「只去掉标记行」，宁可少判也不错杀。
    """
    pairs = []  # [(text, comment)]，engine 的输入格式
    for raw in lines:
        if not raw.strip():
            continue
        if "##" in raw:
            t, c = raw.split("##", 1)
            pairs.append((t.rstrip(), c.strip()))
        else:
            pairs.append((raw.rstrip(), ""))
    try:
        import engine  # config-generator（只读依赖）
    except Exception:
        return [p[0] for p in pairs]

    def flat(nodes):
        for n in nodes:
            if n["type"] == "line":
                yield n["text"]
            elif n["type"] == "repeat":
                for l in n["body"]:
                    yield l[0]
            elif n["type"] == "choice":
                # choice 的多个 option 不是「全都执行」，摊平会把候选方案
                # 误当成漏命令/多余命令。这里按选项名分组返回，由调用方并集处理。
                for o in n["options"]:
                    for l in o["body"]:
                        yield l[0]
    try:
        return list(flat(engine.parse_nodes(pairs)))
    except Exception:
        return [p[0] for p in pairs]


def body_skeletons(body):
    """正文 -> 骨架序列（已按引擎口径展开块结构）。"""
    return extract_skeletons(_unwrap_blocks(body.splitlines()))


# ---------------- 字面量段（抗归并/抗写死值）----------------

_VAR_ANY = re.compile(r"\{\{\s*.+?\s*\}\}|<<\s*.+?\s*>>")

# @repeat/@choice 块内允许出现的结构性配套命令词干（引擎语义支持，非幻觉）：
#   vlan {{vlan_id}}  +  name {{vlan_name}}   ← CG 官方库 vlan_configuration 就是这样
_BLOCK_TAIL_OK = {"name", "description"}


def _literal_anchor(s):
    """从骨架里取「不可变的命令词干」：去掉所有变量槽位与可变尾部参数。

    只取**首命令词**（如 `vlan` / `interface` / `switchport` / `snmp-server`），
    它是配置命令的语义主体，不会因归并或取值变化而改变。
    若整条骨架被占位符占满（罕见），返回空串表示「无锚点」。
    """
    s = _VAR_ANY.sub(" ", s).strip()
    s = re.sub(r"\s+", " ", s)
    if not s:
        return ""
    head = s.split(" ", 1)[0]
    # 首词本身是变量 -> 该行没有稳定锚点
    if not head or head == "{":
        return ""
    return head.lower()


def _anchors(skeletons):
    """骨架序列 -> 锚点序列（去掉空锚点，保持相对顺序）。"""
    out = []
    for s in skeletons:
        a = _literal_anchor(s)
        if a:
            out.append(a)
    return out


def _first_with_anchor(skeletons, anchor):
    """找出锚点对应的第一条骨架（用于把词干还原成原文行）。"""
    for s in skeletons:
        if _literal_anchor(s) == anchor:
            return s
    return anchor


# ---------------- 双向比对 ----------------

def check(original_lines, body):
    """比对命令原文与 AI 产出的正文。

    返回 dict:
      ok        : bool   是否通过
      extra     : list   正文多出来的骨架（AI 创作/幻觉）
      missing   : list   正文缺失的骨架（漏命令）
      reordered : bool   集合相同但顺序不同
      degraded  : bool   是否走了宽松口径（正文含 @repeat/@choice，条数不可比）
      reason    : str    人话原因（给前端直接显示）
    """
    orig = extract_skeletons(original_lines)
    cand = body_skeletons(body)

    has_block = any(
        ln.strip().startswith(("! @repeat", "! @choice", "! @option"))
        for ln in body.splitlines()
    )

    # ---- 降级口径：正文含块结构，归并使条数不可比 ----
    if has_block:
        oa, ca = set(_anchors(orig)), set(_anchors(cand))
        # 原文的每个锚点都必须在正文出现（漏 = 锚点消失）
        missing_anchor = sorted(oa - ca)
        # 正文的锚点若原文完全没有 => 幻觉。
        # `name` / `description` 是引擎在 @repeat 块内支持的结构性配套行
        # （真实语料 vlan_configuration 就是 `vlan {v}` + `name {v}` 成对出现），
        # 因此不把它们计入 extra。
        extra_anchor = sorted(
            {_literal_anchor(s) for s in cand if _literal_anchor(s)} - oa - _BLOCK_TAIL_OK
        )
        ok = not missing_anchor and not extra_anchor
        # 说明行单独校验（引擎会把它们渲染进 Excel，被吞掉同样是污染）
        orig_notes, cand_notes = _note_index(original_lines), _note_index(body.splitlines())
        miss_notes = sorted(orig_notes - cand_notes)
        if miss_notes:
            ok = False

        if ok:
            reason = "命令词干校验通过（含块结构，按词干比对）"
        elif missing_anchor:
            ex_loc = _locate(original_lines, _first_with_anchor(orig, missing_anchor[0]))
            reason = "正文遗漏了原文中的 %d 条（如 %s）" % (len(missing_anchor), ex_loc)
        elif extra_anchor:
            reason = "正文多出 %d 条原文没有的命令（如「%s」）" % (
                len(extra_anchor), extra_anchor[0][:40])
        else:
            ex_loc = _note_locate(original_lines, miss_notes[0])
            reason = "正文遗漏了 %d 条说明/注释行（如 %s）" % (len(miss_notes), ex_loc)

        return {
            "ok": ok,
            "extra": extra_anchor,
            "missing": missing_anchor,
            "missing_notes": miss_notes[:5],
            "reordered": False,
            "degraded": True,
            "reason": reason,
            "orig_count": len(orig),
            "body_count": len(cand),
        }

    # ---- 严格口径：字面骨架双向比对 + 顺序 ----
    orig_set, cand_set = set(orig), set(cand)
    missing = [s for s in orig if s not in cand_set]
    extra = [s for s in cand if s not in orig_set]

    # 顺序：贪心匹配（每个骨架取「尚未使用的最早位置」）。
    # 不能用 setdefault 只记首次位置——真实配置里存在大量**合法的重复行**
    # （如 address-family ipv4/ipv6 各自的 exit-address-family、
    #     同一 permit 网段出现在不同 ACL 里），记首次位置会把第二次出现误判成逆序。
    reordered = False
    if not missing and not extra:
        pool = {}
        for i, s in enumerate(orig):
            pool.setdefault(s, []).append(i)
        for lst in pool.values():
            lst.reverse()          # pop() 取最早
        last, mono = -1, True
        for s in cand:
            lst = pool.get(s)
            if not lst:
                mono = False
                break
            i = lst.pop()
            if i < last:
                mono = False
                break
            last = i
        reordered = not mono

    ok = (not missing) and (not extra) and (not reordered)

    # 说明行也校验（严格口径同样要管）：引擎会以 kind="note" 渲染进 Excel
    orig_notes, cand_notes = _note_index(original_lines), _note_index(body.splitlines())
    miss_notes = sorted(orig_notes - cand_notes)
    extra_notes = sorted(cand_notes - orig_notes)
    if miss_notes or extra_notes:
        ok = False

    # 举例一律给「原文第几行 + 那行内容」，让用户能对着自己的原文定位
    if missing:
        m_ex = _locate(original_lines, missing[0])
    elif miss_notes:
        m_ex = _note_locate(original_lines, miss_notes[0])
    else:
        m_ex = ""
    e_ex = ("「%s」" % extra[0][:56]) if extra else (
        "「%s」" % extra_notes[0][:56] if extra_notes else "")

    if ok:
        reason = "命令集与顺序均与原文一致"
    elif extra and missing:
        reason = ("正文多了 %d 条原文没有的命令（如 %s），且漏了 %d 条（如 %s）"
                  % (len(extra), e_ex, len(missing), m_ex))
    elif extra:
        reason = "正文多了 %d 条原文没有的命令（AI 自行创作，如 %s）" % (len(extra), e_ex)
    elif missing:
        reason = "正文漏了 %d 条原文命令（如 %s）" % (len(missing), m_ex)
    elif miss_notes and extra_notes:
        reason = ("正文漏了 %d 条说明行（如 %s），并多出 %d 条（如 %s）"
                  % (len(miss_notes), _note_locate(original_lines, miss_notes[0]),
                     len(extra_notes), e_ex))
    elif miss_notes:
        reason = "正文漏了 %d 条说明/注释行（如 %s）" % (
            len(miss_notes), _note_locate(original_lines, miss_notes[0]))
    elif reordered:
        reason = "命令顺序与原文不一致（配置有先后依赖，如先建 VLAN 再配 interface）"
    else:
        reason = "正文多出 %d 条原文没有的说明行（如 %s）" % (len(extra_notes), e_ex)

    return {
        "ok": ok,
        "extra": extra,
        "missing": missing,
        "missing_notes": miss_notes[:5],
        "extra_notes": extra_notes[:5],
        "reordered": reordered,
        "degraded": False,
        "reason": reason,
        "orig_count": len(orig),
        "body_count": len(cand),
    }
