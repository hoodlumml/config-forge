# Config Forge（配置锻造）

一个「双击即用」的网络配置生成工具：用自然语言描述需求，或手动选择模板填参数，
由确定性渲染引擎产出多 sheet 的 Excel 配置清单（每 sheet = 一台设备，A 列命令 / B 列注释）。

## 特点

- **便携**：内置 Python 运行时（`runtime/`），目标机器无需安装任何环境。
- **离线优先**：默认用本地大模型（dsh）完全离线生成，无需联网、无需 Key。
- **云端可选**：在网页「设置」里填入 DeepSeek / OpenAI 兼容端点的 Key 即可切换。
- **模板化**：内置交换机接入层（`switch/access`）等模板，支持版本化管理与工作室编辑。

## 介绍视频

> GitHub 不支持在 README 内直接播放视频（会剥离播放器标签）。点击下方链接前往 B 站观看完整演示。

- **[B 站 · config-forge（视频演示）](https://www.bilibili.com/video/BV1wqHD6qEsv/)** —— 便携网络配置生成器：自然语言生成交换机配置，确定性渲染为多 sheet Excel，内置运行时无需安装。
- [Releases 资产 config-forge.mp4（31MB，直接下载）](https://github.com/hoodlumml/config-forge/releases/download/v1.0.0/config-forge.mp4)

## 快速开始

### 方式一：便携包（推荐普通用户）

1. 从 [Releases](../../releases) 下载**带 `runtime/` 的完整压缩包**，解压到任意目录
   （必须连 `runtime/` 一起，只拿 `orchestrator/` 会跑不起来）。
2. 进入 `orchestrator/`，双击 `start.bat`。
3. 约 3 秒后浏览器自动打开 http://localhost:18000
   （没自动开就手动打开；改端口：`start.bat 9000`）。
4. 在网页里用自然语言描述需求，或手动选模板填参数；生成的 Excel 落在
   `engine_runtime/output/`。

### 方式二：从源码运行

仓库本身只含源码，不含 `runtime/`。若要从源码跑，需自行准备 Python 3.12+ 与
`openpyxl`，并补齐运行时后再启动 `orchestrator/start.bat`。多数用户直接用方式一的便携包即可。

## 配置模型

网页右上角「设置」可配置以下模型端点（Key 仅保存在本机
`orchestrator/config.local.json`，**不进仓库**，保存时自动重建）：

| 名称 | 类型 | 默认地址 |
| --- | --- | --- |
| `local_dsh` | 本地 | http://localhost:8080 |
| `local_ollama` | 本地 | http://localhost:11434 |
| `cloud_openai` | 云端 | DeepSeek / OpenAI 兼容端点（自填 Key） |

## 目录结构

```
config-forge/
├─ orchestrator/      # 网页服务（app.py / index.html / skills / template_studio）
│  └─ start.bat       # 启动入口，双击它
├─ engine_runtime/    # 渲染引擎 + 模板库（只读运行，勿改）
│  ├─ engine.py / library.py / main.py / _forge_runner.py
│  ├─ library/        # 模板库（index.json + 各模板版本目录）
│  └─ output/         # 生成的 Excel 落这里
├─ runtime/           # 便携 Python（发布包内含，仓库不纳入）
├─ README.md
├─ LICENSE
└─ .gitignore
```

## 许可证

[MIT](LICENSE)
