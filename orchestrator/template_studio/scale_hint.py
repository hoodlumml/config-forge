#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""scale_hint —— 导入材料规模评估（只算规模，不判断内容对错）。

为什么需要
----------
用户导入的常常是「一整台交换机的配置模板」，动辄几百行。这里存在两条独立的失效路径，
但它们的处理方式完全不同，提示必须能区分开：

  A. **规模风险（长度导致）**：LLM 输出被截断、长上下文注意力失焦导致中间行被改写。
     → 这是**工程问题**，根治靠分批，重复点几次「重新生成」没用。
  B. **内容问题**：AI 漏命令 / 改写法 / 吞说明行。
     → 这是**AI 改写**，靠保真校验拦，与长度无关。

之前的报错把两者混在一起（只说「不一致」），用户无法判断该重试还是该分批。
本模块产出的 scale 信息专门用于把这两种情况分开说明。

阈值依据（CG 官方库 17 个真实模板实测）
----------------------------------------
    qos_configuration          244 行 / 238 条命令   <- 已入库、生产在用，不算大
    basic_configuration        140 行 / 139 条命令   <- 同上
    aaa_tacacs_configuration    73 行 /  66 条命令
    snmp_configs                68 行 /  68 条命令
    中位数                      46 行 /  25 条命令

初版把「偏大」门槛设成 120，结果**把上面两个在用模板全判成偏大**——这是误报。
说明「能否生成成功」与「条数」并非简单正相关（模板已有分段与归并时更耐长）。
故门槛按**覆盖全部在用模板**上界定，不拍脑袋：

    ok    : ≤ 240 条命令   正常（覆盖全部 17 个官方模板，含 238 条的 qos）
    large : 241 ~ 500      偏大：建议按段落分段导入
    huge  : > 500          很大：单次导入极可能截断，强烈建议分批

注意：这是**提示**而非**拦截**——规模大不直接拒绝生成（因为已有 238 条的成功先例），
只提示风险与建议，把决定权留给用户。
"""

# 门槛（按「命令条数」计，不按总行数——说明行与注释不算压力）
_OK_MAX = 240
_LARGE_MAX = 500


def assess(original_lines):
    """评估导入材料的规模。

    original_lines : list[str]  命令原文的各行
    返回 dict：
      total_lines 总行数（去空白）
      cmd_lines  命令条数（排除说明行/段头/块标记，与 seam_guard 同一口径）
      note_lines 说明行条数
      level      'ok' | 'large' | 'huge'
      label      中文档位名
      advice     给用户的建议（可直接显示）
      hint       一句话提示，用于状态栏
    """
    from template_studio import seam_guard as G

    total = 0
    notes = 0
    cmds = 0
    for ln in original_lines:
        s = G._norm(ln)
        if not s:
            continue
        total += 1
        if s.startswith("!"):
            notes += 1
        else:
            cmds += 1

    if cmds <= _OK_MAX:
        level = "ok"
        label = "规模正常"
        advice = ""
        hint = "共 %d 条命令，规模正常（官方模板中最大的一条为 238 条）。" % cmds
    elif cmds <= _LARGE_MAX:
        level = "large"
        label = "偏大"
        advice = ("这份材料有 %d 条命令，已超过单次输出的稳定区间。"
                  "一次生成失败时重复点「重新生成」通常没用——建议按段落分成"
                  "「基础 / 接口 / 路由 / 管理 / 策略」几张表分别导入，"
                  "每张 100~150 条命令，成功率明显更高。" % cmds)
        hint = "共 %d 条命令，材料偏大，建议分段导入。" % cmds
    else:
        level = "huge"
        label = "很大"
        advice = ("这份材料有 %d 条命令，单次导入极可能被截断（输出超长后，"
                  "后半段命令会静默消失）。强烈建议先按段落拆成多张表，"
                  "每张不超过 150 条命令，再分别入库。" % cmds)
        hint = "共 %d 条命令，材料很大，建议务必分段导入。" % cmds

    return {
        "total_lines": total,
        "cmd_lines": cmds,
        "note_lines": notes,
        "level": level,
        "label": label,
        "advice": advice,
        "hint": hint,
    }


def explain_reject(scale, seam_reason):
    """拦截时把「规模问题」和「内容问题」分开说。

    内容问题永远要说清是哪里不一致（seam_reason 已带原文行号）；
    若同时规模偏大，额外提示长度可能是诱因。
    """
    msg = seam_reason
    if scale and scale.get("level") in ("large", "huge"):
        msg += ("　另外：这份材料有 %d 条命令（%s），"
                "如果反复出现同类失败，建议按段落拆成每张 100~150 条命令分别导入。"
                % (scale["cmd_lines"], scale["label"]))
    return msg
