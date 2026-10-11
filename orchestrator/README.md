# CONFIG FORGE 编排程序（便携分发版）

一个**独立、零三方依赖**的本地 Web 程序：Windows 下双击 `start.bat` 即可使用（包内自带便携 Python，无需安装任何东西）；macOS / Linux 下用包内 `runtime/python.exe` 启动。

---

## ⚠️ 升级后必读：模板导入多了一道「命令保真校验」

**如果你以前导入过模板，现在可能会看到以前没见过的拦截提示。这是设计升级，不是坏了。**

### 发生了什么

以前 AI 包装完模板，系统只查「语法能不能渲染」。于是 AI 悄悄漏掉一行、改掉一个写法、把写死的值当变量——系统都发现不了：**页面显示成功、引擎校验通过、Excel 也生成了，但配置是错的**，往往要到设备上试出来才发现。

现在多了一道**逐行核对**：AI 写出的每一行，都必须能在你粘贴的原文里找到对应，否则直接拦下，并告诉你**原文第几行对不上**。

### 看到拦截提示怎么办

提示长这样：

```
模板正文与命令原文不一致，已拦下：正文漏了 1 条原文命令（如 第8行「logging host <<syslog>>」）
```

它会明确告诉你是**多了还是少了、原文第几行、内容是什么**。按提示三选一处理：

| 情况 | 怎么办 |
|------|--------|
| 确实漏了 / 写错了 | 点「① 重新生成」再来一次（多数情况一次就好） |
| 原文写 `<<x>>`，AI 转成 `{{x}}` | **这是正常的，不算错**，不会拦 |
| 反复几次都失败 | 说明材料太长，见下面「材料规模」 |

### 材料规模提示

导入整台交换机的配置时，命令区上方会显示一条**规模档位**：

| 档位 | 命令条数 | 建议 |
|------|----------|------|
| 🟢 规模正常 | ≤ 240 | 直接生成 |
| 🟡 偏大 | 241 ~ 500 | 反复失败时，建议按段落拆成每张 100~150 条命令 |
| 🔴 很大 | > 500 | **建议先拆开再导入**，单次生成极易被截断 |

> 反复点「重新生成」对长材料没用（输出超长时后半段会被截断），**拆开才是正解**。
> 参照量：官方模板库中最大的一条是 238 条命令（`qos_configuration`），属于正常范围。

### 为什么这么做

CONFIG FORGE 的核心价值是「**确定性引擎渲染 + 人不确认不落盘**」。既然目标是确定性，AI 就不该有偷偷改命令的自由——这套校验就是为此设计的。

**AI 只做语法包装，绝不创作 / 改写命令**，这条铁律从 prompt 里的文字变成了代码强制。

---

## 核心特性

- **开箱即用**：仅用 Python 标准库，无 pip 安装。包内自带 `runtime/python.exe`（便携 Python）与 `engine_runtime/`（渲染引擎 + 官方模板库），整个文件夹拷到别的电脑上就能用。
- **调用方式可选**：自动 / 手动两种模式。
  - **自动模式**：按 `config.json` 里的 `auto_rules` 规则选模型（默认：命中涉密关键词 → 本地，否则 → 云端；云端不可达 → 降级本地）。
  - **手动模式**：下拉框**只列出本地模型**，强制只能调用本地大模型（dsh），云端根本不出现（前端隐藏 + 后端二次拦截双保险）。
- **自动调用规则可调**：规则写在 `config.json` 的 `auto_rules`，支持 `keyword` / `regex` / `always` 三种匹配，不写死在代码里。
- **本地模型用 dsh 托管**：`local_dsh` 走 OpenAI 兼容接口，把 `base_url` 指向你的 dsh 实例即可。

## 目录结构

```
config-forge-portable/
|-- orchestrator/
│   ├── app.py           # 本地 Web 服务（标准库 http.server）
│   ├── llm.py           # LLMProvider 抽象 + Orchestrator 调度核心
│   ├── config.json      # 模型配置 + 自动路由规则（改这里）
│   ├── index.html       # 浏览器前端
│   ├── start.bat        # Windows 双击启动（UTF-8 防乱码 + 端口处理）
│   └── template_studio/ # 模板工作室（导入 / 校验 / 入库 / 客户管理）
├── runtime/             # 便携 Python（系统无需装 Python）
├── engine_runtime/      # CG 渲染引擎 + 官方模板库（17 个，只读）
└── README.md
```

## 运行

### Windows（推荐）

```
双击 orchestrator\start.bat
```

自动起服务并打开浏览器，默认 <http://localhost:18000>。脚本已内置 `chcp 65001`，避免中文乱码。

### macOS / Linux

```bash
# 1) 配置云端密钥
export OPENAI_API_KEY=sk-xxxx

# 2) 改 orchestrator/config.json，把 local_dsh.base_url 指向你的 dsh OpenAI 兼容端口

# 3) 用包内 Python 启动
./runtime/python.exe orchestrator/app.py
# 浏览器打开 http://localhost:18000
```

> 端口可用环境变量覆盖：`PORT=19000 ./runtime/python.exe orchestrator/app.py`

## 配置说明（config.json）

```jsonc
{
  "providers": {
    "cloud_openai": { "kind": "cloud", "base_url": "https://api.openai.com", "api_key": "${OPENAI_API_KEY}", "model": "gpt-4o-mini" },
    "local_dsh":    { "kind": "local", "base_url": "http://localhost:8080", "api_key": "", "model": "deepseek-chat" }
  },
  "auto_rules": [
    {"name": "sensitive_to_local", "type": "keyword",
     "keywords": ["配置","拓扑","IP","BOM","密码","密钥","secret","192.168"], "provider": "local_dsh" },
    {"name": "default_cloud", "type": "always", "provider": "cloud_openai" }
  ],
  "fallback_to_local": true
}
```

- `providers` 里每个模型：`kind` 必须是 `cloud` 或 `local`；`base_url` 指向 OpenAI 兼容端点；`api_key` 支持 `${ENV}` 从环境变量取。
- `auto_rules` 按顺序匹配，命中第一条即选用对应 `provider`；最后一条用 `type: always` 兜底。
- 加国产大模型：复制一个 `cloud_xxx` 条目，`base_url` 填其 OpenAI 兼容网关即可。

## 模板与客户

- **官方模板库**（`engine_runtime/library`，17 个）是只读基准，程序不会修改它。
- 在「模板工作室」把官方模板**复制（fork）给某个客户**后，该客户即可使用自己那份；**未复制的模板仍用官方版**。各客户之间互不干扰。
- 所以给客户改配置时，**只需要复制你要改的那几个模板**，不必把 17 个全复制一遍。

## API

| 方法 | 路径 | 说明 |
|------|------|------|
| GET  | `/api/providers?mode=auto\|manual` | 列出可选模型（手动模式只返回本地） |
| GET  | `/api/health` | 各模型连通性探测 |
| POST | `/api/chat` | body：`{"mode","provider?","message"}` → 返回实际调用的模型与回答 |
| POST | `/api/template/draft` | AI 包装模板草稿。含命令保真校验，不合格返回 502 并指出原文行号 |
| POST | `/api/template/verify` | 运行校验：引擎硬校验 + 命令保真复检（手改正文后可再检一次） |
| POST | `/api/template/commit` | 确认入库（三重门控：校验通过 + 已逐行检查 + 二次确认） |
| POST | `/api/template/scale` | 只算材料规模（生成前预检），返回档位与建议 |

## 设计原则

> 手动模式永远不碰云端；涉密内容自动锁本地；**人不点确认键，配置不落盘**；**AI 不得创作命令**。
> 这些是合规与正确性底线，不要为了「省事」放开。
