# LLM 编排程序（跨平台演示版）

一个**独立、零三方依赖**的本地 Web 程序：任意装了 Python 3 的操作系统（Windows / macOS / Linux）直接 `python app.py` 起服务，开浏览器即可演示。

## 核心特性（对应你的需求）

- **跨平台**：仅用 Python 标准库，无 pip 安装步骤。演示机只要有 Python 3 就能跑。
- **调用方式可选**：自动 / 手动两种模式。
  - **自动模式**：按 `config.json` 里的 `auto_rules` 规则选模型（默认：命中涉密关键词→本地，否则→云端；云端不可达→降级本地）。
  - **手动模式**：下拉框**只列出本地模型**，强制只能调用本地大模型（dsh），云端根本不出现（前端隐藏 + 后端二次拦截双保险）。
- **自动调用规则可调**：规则写在 `config.json` 的 `auto_rules`，支持 `keyword` / `regex` / `always` 三种匹配，不写死在代码里。
- **本地模型用 dsh 托管**：`local_dsh` 走 OpenAI 兼容接口，把 `base_url` 指向你的 dsh 实例即可。

## 目录结构

```
llm-orchestrator/
├── app.py          # 本地 Web 服务（标准库 http.server）
├── llm.py          # LLMProvider 抽象 + Orchestrator 调度核心
├── config.json     # 模型配置 + 自动规则（改这里）
├── index.html      # 浏览器前端
├── Dockerfile      # 可选：容器化分发
└── README.md
```

## 运行（跨平台）

### 方式一：直接跑（推荐演示用）
```bash
# 1) 配置云端密钥（Linux/macOS）
export OPENAI_API_KEY=sk-xxxx
# Windows:
# set OPENAI_API_KEY=sk-xxxx

# 2) 改 config.json，把 local_dsh.base_url 指向你的 dsh OpenAI 兼容端口

# 3) 启动
python app.py
# 浏览器打开 http://localhost:8000
```

### 方式二：Docker（零环境依赖）
```bash
docker build -t llm-orchestrator .
docker run -p 8000:8000 -e OPENAI_API_KEY=sk-xxxx llm-orchestrator
```
> 注意：容器内访问宿主机 dsh 需把 `local_dsh.base_url` 改为宿主机网关地址（如 `http://host.docker.internal:8080`），不能写 `localhost`。

## 配置说明（config.json）

```jsonc
{
  "providers": {
    "cloud_openai": { "kind": "cloud", "base_url": "https://api.openai.com", "api_key": "${OPENAI_API_KEY}", "model": "gpt-4o-mini" },
    "local_dsh":    { "kind": "local", "base_url": "http://localhost:8080", "api_key": "", "model": "deepseek-chat" }
  },
  "auto_rules": [
    { "name": "sensitive_to_local", "type": "keyword",
      "keywords": ["配置","拓扑","IP","BOM","密码","密钥","secret","192.168"], "provider": "local_dsh" },
    { "name": "default_cloud", "type": "always", "provider": "cloud_openai" }
  ],
  "fallback_to_local": true
}
```

- `providers` 里每个模型：`kind` 必须是 `cloud` 或 `local`；`base_url` 指向 OpenAI 兼容端点；`api_key` 支持 `${ENV}` 从环境变量取。
- `auto_rules` 按顺序匹配，命中第一条即选用对应 `provider`；最后一条用 `type: always` 兜底。
- 加国产大模型：复制一个 `cloud_xxx` 条目，`base_url` 填其 OpenAI 兼容网关即可。

## API

| 方法 | 路径 | 说明 |
|------|------|------|
| GET  | `/api/providers?mode=auto\|manual` | 列出可选模型（手动模式只返回本地） |
| GET  | `/api/health` | 各模型连通性探测 |
| POST | `/api/chat` | body：`{"mode","provider?","message"}` → 返回实际调用的模型与回答 |

## 下一步（接你的生成器 / skill）

当前程序完成了「LLM 调度层」。把配置/BOM/ROM 生成器接进来有两种做法：
1. **本地优先**：在 `llm.py` 的 `complete()` 之前加一层「意图路由」，识别要调哪个生成器 skill，再拼装 prompt 发模型。
2. **MCP 接 Office**：加一个 MCP 客户端，模型产出结构化结果后，调 Excel/Word/PPT 工具按模板落盘（手动模式写文件前回显参数让人确认）。

> 设计原则：手动模式永远不碰云端；涉密内容自动锁本地。这是合规底线，不要为了"省事"放开。
