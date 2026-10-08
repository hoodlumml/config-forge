#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""xlsx 解析探针：仅在有 openpyxl 的 .venv 里由子进程执行，stdout 输出 sections JSON。

用法: python _xlsx_probe.py <xlsx路径> <CG_DIR用于import engine>
模板工作室自有文件，不属于 CG，CG 零改动。

两个必须做的编码防护（2026-10-05 现场报错修复）：
1) NBSP(\\xa0) 等字符在 GBK 里无法编码，而 Windows 子进程 stdout 重定向到管道时
   默认按 locale(GBK) 编码，一旦 Excel 内容带 NBSP 就会 UnicodeEncodeError 崩溃。
   → 显式把 stdout 改成 utf-8，与父进程 subprocess(encoding="utf-8") 对齐。
2) NBSP / 全角空格 / 零宽字符进到配置命令区是非法字符（Cisco 命令不认），
   → 解析结果里统一归一成半角空格 / 删除，避免脏字符流进模板正文。
"""
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))  # 让 template_studio 包内模块可导入
from _norm import norm_deep  # noqa: E402

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass  # 老版本/非管道场景忽略，退化到默认编码。关键：GBK 编不了 \xa0

sys.path.insert(0, sys.argv[2])  # CG_DIR：让 importer 的 `import engine` 能解析到
import importer  # noqa: E402  (template_studio 目录已在 sys.path[0])

print(json.dumps(norm_deep(importer.read_xlsx_sections(sys.argv[1])), ensure_ascii=False, default=str))
