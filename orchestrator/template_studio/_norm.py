# -*- coding: utf-8 -*-
"""模板工作室：命令区文本归一（单一实现，probe 与 app.py 共用）。

为什么必须有这一层：
1) Excel / 网页 / PDF 粘进来的中文资料常带 NBSP(\\u00a0)、全角空格(\\u3000)、
   零宽空格(\\u200b)。NBSP 在 GBK 里**无法编码**，Windows 子进程按 locale(GBK)
   输出 stdout 时会 UnicodeEncodeError 崩溃（2026-10-05 现场报错）。
2) 就算传得回来，NBSP / 零宽字符进到 Cisco 配置命令区也是非法字符。

注意：必须用 chr() 显式构造字符表——编辑器/写入链路会把字面量 NBSP、零宽空格
归一成普通空格，写死字面量会静默失效。
"""

NBSP_LIKE = tuple(chr(c) for c in (0x00A0, 0x1680, 0x2000, 0x2001, 0x2002,
                                   0x2003, 0x2004, 0x2005, 0x2006, 0x2007,
                                   0x2008, 0x2009, 0x200A, 0x202F, 0x205F, 0x3000))
ZERO_WIDTH = tuple(chr(c) for c in (0x200B, 0x200C, 0x200D, 0xFEFF, 0x2060))


def norm_text(s):
    """把不可见空白归一成半角空格、零宽字符删除。普通空格数量不变，不动其它字符。"""
    if not isinstance(s, str):
        return s
    for ch in NBSP_LIKE:
        if ch in s:
            s = s.replace(ch, " ")
    for ch in ZERO_WIDTH:
        if ch in s:
            s = s.replace(ch, "")
    return s


def norm_ws(text):
    """整段文本归一（命令区/正文入口用）。"""
    return norm_text(text or "")


def norm_deep(obj):
    """递归归一：importer 返回的是 tuple（lines 为 tuple of tuple），
    tuple 与 list 都要处理，否则脏字符会从 tuple 分支原样漏出去。"""
    if isinstance(obj, str):
        return norm_text(obj)
    if isinstance(obj, (list, tuple)):
        return type(obj)(norm_deep(x) for x in obj)
    if isinstance(obj, dict):
        return {k: norm_deep(v) for k, v in obj.items()}
    return obj
