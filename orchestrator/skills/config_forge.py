#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CONFIG FORGE skill —— 离线网络配置生成（确定性引擎，LLM 不做数）。

设计要点：
- 解耦：用包内 runtime/ 便携 Python 以子进程调用引擎，
  避免把 openpyxl / PySide 塞进 route-D 程序（managed python 无这些依赖）。
- 安全性：config-generator 源码只读依赖，本 skill 不修改它任何文件；
  仅在其 output/ 目录生成 Excel（即工具原本行为）。
- LLM 角色：只负责"听懂请求 → 选模板 → 按 prompts 填 data"，
  真正渲染由 engine.render_nodes 确定性完成。

输入 spec 示例：
{
  "devices": [
    {
      "hostname": "SW-ACCESS-01",
      "category": "switch/access",
      "templates": [{"name": "upgrade_configuration", "version": "v1.0"}],
      "data": {
        "upgrade_configuration": {
          "vars": {"ios_name_checked_from_pg_dtl": "cat9k_iosxe.17.12.01.SPA.bin"},
          "repeats": {}, "choices": {}
        }
      }
    }
  ]
}
"""
import json
import subprocess
import tempfile
from pathlib import Path

# 自包含包：不再依赖外部 config-generator 源码或 .venv。
# 运行时 Python 用包内 runtime/，引擎运行器用包内 engine_runtime/_forge_runner.py。
PACKAGE_ROOT = Path(__file__).resolve().parents[2]  # skills/ -> orchestrator/ -> config-forge-portable/
RUNTIME_PY = PACKAGE_ROOT / "runtime" / "python.exe"
RUNNER = PACKAGE_ROOT / "engine_runtime" / "_forge_runner.py"
CFG_DIR = PACKAGE_ROOT / "engine_runtime"

SKILL = {
    "name": "config_forge",
    "description": "根据模板库与设备参数，生成多 sheet 网络配置 Excel（每 sheet = 一台设备，A列命令/B列注释，样式沿用模板）。纯离线、确定性渲染。",
    "trigger": ["配置", "config", "forge", "模板", "交换机配置", "upgrade", "stack", "生成配置"],
    "input_schema": {
        "devices": "list，每项 {hostname, category?, templates:[{name, version?}], stack_member?, data?}",
        "data": "每模板 {vars:{slug:值}, repeats:{rid:[...]}, choices:{cid:选项名}}；slug=变量名 slugify 后",
    },
    "run": "run",  # orchestrator 调用 skills.config_forge.run(spec)
    "engine_isolation": "subprocess -> runtime/ python + engine_runtime/",
}


def run(spec: dict) -> dict:
    """调用 CONFIG FORGE 引擎生成配置 Excel，返回 {output: xlsx路径}。"""
    if not RUNTIME_PY.exists():
        raise RuntimeError(
            f"包内 Python 运行时缺失: {RUNTIME_PY}\n"
            f"请确认 runtime/ 目录已随包一起拷贝（不要只拷贝 orchestrator/）。"
        )
    if not RUNNER.exists():
        raise RuntimeError(f"引擎运行器缺失: {RUNNER}")

    spec_path = Path(tempfile.gettempdir()) / "forge_spec.json"
    spec_path.write_text(json.dumps(spec, ensure_ascii=False), encoding="utf-8")

    proc = subprocess.run(
        [str(RUNTIME_PY), str(RUNNER), str(spec_path)],
        capture_output=True,
        text=True,
        encoding="utf-8",
        timeout=300,
    )
    if proc.returncode != 0:
        raise RuntimeError(f"CONFIG FORGE 引擎执行失败:\n{proc.stderr.strip()}")
    return json.loads(proc.stdout.strip())
