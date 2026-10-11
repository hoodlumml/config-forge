"""LLM 提供商抽象 + 自动/手动调度核心（零三方依赖，仅用标准库）。

- LLMProvider：统一 OpenAI 兼容接口（云端 / 本地 dsh 都走 /v1/chat/completions）
- Orchestrator：自动模式按 auto_rules 选模型；手动模式强制只能用本地模型
"""
import json
import os
import re
import urllib.request
import urllib.error


def _resolve_env(val):
    """配置里写 ${ENV_VAR} 时从环境变量取值。"""
    if isinstance(val, str) and val.startswith("${") and val.endswith("}"):
        return os.environ.get(val[2:-1], "")
    return val


def load_config(path):
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


class LLMProvider:
    def __init__(self, name, kind, base_url, api_key="", model="", extra=None):
        self.name = name
        self.kind = kind  # "cloud" | "local"
        self.base_url = (base_url or "").rstrip("/")
        self.api_key = api_key
        self.model = model
        self.extra = extra or {}

    def chat(self, messages, json_mode=False, **kwargs):
        url = f"{self.base_url}/v1/chat/completions"
        payload = {"model": self.model, "messages": messages, "stream": False}
        # 默认给足输出长度，避免本地模型把长 spec JSON 截断（此前规划报错 502 的根因之一）。
        # max_tokens 是 OpenAI 标准字段，DeepSeek 与 Ollama 的 OpenAI 兼容端点都认。
        if "max_tokens" not in kwargs:
            payload["max_tokens"] = 8192
        # 本地模型默认上下文偏小，按需抬一下 num_ctx（仅本地生效；云端 DeepSeek 不传，避免未知字段）。
        if self.kind == "local" and "num_ctx" not in kwargs:
            payload["num_ctx"] = 8192
        # 规划类调用强制 JSON 输出：本地 Ollama 用 format:"json"，云端 OpenAI 兼容用
        # response_format=json_object。从协议层保证吐出的一定是合法 JSON 对象，
        # 杜绝「双层 JSON / 过度宽容砍内容 / 静默丢命令」三类脏输出（事后救援逻辑仍保留作备份）。
        if json_mode:
            if self.kind == "local":
                payload["format"] = "json"
            else:
                payload["response_format"] = {"type": "json_object"}
        # 结构化解析（json_mode）需确定性输出：本地模型默认 temperature≈0.8，
        # 会导致同一输入每次结果不同（导入模板时好时坏）。强制 temperature=0；
        # 若调用方显式传了 temperature，以调用方为准（必须放在 update(kwargs) 之前才生效）。
        if json_mode and "temperature" not in kwargs:
            payload["temperature"] = 0
        payload.update(kwargs)
        data = json.dumps(payload).encode("utf-8")
        req = urllib.request.Request(url, data=data, method="POST")
        req.add_header("Content-Type", "application/json")
        if self.api_key:
            req.add_header("Authorization", f"Bearer {self.api_key}")
        try:
            with urllib.request.urlopen(req, timeout=60) as resp:
                out = json.loads(resp.read().decode("utf-8"))
        except urllib.error.HTTPError as e:
            # 把服务端错误体带出来，别只报 "HTTP Error 400" 这种没头没尾的信息
            try:
                body = e.read().decode("utf-8", "replace")[:300]
            except Exception:
                body = ""
            self._last_error = "HTTP %s: %s" % (e.code, body or e.reason)
            raise RuntimeError(self._last_error) from None
        except Exception as e:
            # 连接类错误（端口未监听/超时/密钥无效）也要留痕，供回退提示用
            self._last_error = "%s: %s" % (type(e).__name__, e)
            raise
        self._last_error = None
        return out["choices"][0]["message"]["content"]

    def healthy(self):
        try:
            self.chat([{"role": "user", "content": "ping"}], max_tokens=1)
            return True
        except Exception:
            return False

    def last_error_summary(self):
        """一句话说明上次为什么失败（用于把「都试过了」讲清楚）。"""
        e = getattr(self, "_last_error", None)
        if not e:
            return "未记录"
        s = str(e).strip().replace("\n", " ")
        if "10061" in s or "actively refused" in s or "拒绝" in s:
            return "服务未启动 / 端口未监听"
        if "10060" in s or "timed out" in s.lower() or "超时" in s:
            return "连接超时"
        if "401" in s or "403" in s or "api_key" in s.lower() or "unauthorized" in s.lower():
            return "密钥无效或无权限"
        return s[:60]


def _is_simple(exc):
    """错误对象能否安全地按原类型重建（自定义异常可能构造签名不同，保守处理）。"""
    return isinstance(exc, (RuntimeError, ValueError, OSError, PermissionError))


def _fallback_hint(provider, tried):
    """全部 provider 都失败时，把「还试过谁、为什么不行」讲清楚。

    只说「连接被拒绝」而不说「本地服务没启动」，用户无法判断该重启服务还是改配置。
    """
    if not tried:
        return ""
    lines = ["\n\n（已依次尝试过以下模型，全部失败："]
    for p in tried:
        kind = "云端" if getattr(p, "kind", "local") == "cloud" else "本地"
        lines.append("  · %s（%s %s）：%s" % (p.name, kind, p.base_url, p.last_error_summary()))
    lines.append("）")
    lines.append("\n排查：若为本地模型，请确认其服务已启动（如 dsh 默认 http://localhost:8080、"
                 "ollama 默认 http://localhost:11434）。")
    return "".join(lines)


def build_orchestrator(cfg):
    providers = {}
    for name, p in cfg.get("providers", {}).items():
        providers[name] = LLMProvider(
            name=name,
            kind=p.get("kind", "cloud"),
            base_url=_resolve_env(p.get("base_url", "")),
            api_key=_resolve_env(p.get("api_key", "")),
            model=p.get("model", ""),
            extra=p.get("extra"),
        )
    return Orchestrator(providers, cfg.get("auto_rules", []), cfg.get("fallback_to_local", True))


class Orchestrator:
    def __init__(self, providers, rules, fallback_to_local=True):
        self.providers = providers
        self.rules = rules
        self.fallback_to_local = fallback_to_local

    def list_providers(self, mode):
        """按模式返回可选模型：手动模式只返回本地模型。"""
        items = []
        for name, p in self.providers.items():
            if mode == "manual" and p.kind != "local":
                continue
            items.append({"name": name, "kind": p.kind, "model": p.model})
        return items

    def _match(self, rule, text):
        t = rule.get("type")
        if t == "always":
            return True
        if t == "keyword":
            kws = rule.get("keywords", [])
            return any(kw.lower() in text.lower() for kw in kws)
        if t == "regex":
            return re.search(rule.get("pattern", ""), text, re.IGNORECASE) is not None
        return False

    def decide(self, mode, manual_provider, text):
        """只做选择，不发起真实调用。"""
        if mode == "manual":
            if not manual_provider:
                # 手动模式未显式选模型：默认取第一个本地 provider（保持"手动只走本地"约束；
                # 结合 complete() 的本地互备，不可用的本地模型会被自动跳过）。
                local = next((p for p in self.providers.values() if p.kind == "local"), None)
                if local is None:
                    raise ValueError("手动模式必须选择一个本地模型（未配置任何 kind=local 的 provider）")
                return local
            p = self.providers.get(manual_provider)
            if p is None:
                raise ValueError(f"未知 provider: {manual_provider}")
            if p.kind != "local":
                # 双保险：手动模式绝不调用云端
                raise PermissionError("手动模式仅允许调用本地大模型")
            return p
        # 自动模式：按规则顺序匹配
        chosen = None
        for rule in self.rules:
            if self._match(rule, text):
                chosen = self.providers.get(rule.get("provider"))
                if chosen:
                    break
        if chosen is None:
            chosen = next((p for p in self.providers.values() if p.kind == "cloud"), None)
        if chosen is None:
            chosen = next((p for p in self.providers.values() if p.kind == "local"), None)
        if chosen is None:
            raise RuntimeError("无可用模型")
        return chosen

    def complete(self, mode, manual_provider, message, route_text=None, json_mode=False):
        """选择并发起调用；失败时按「同类互备 → 另一类」顺序回退。

        - auto：首选按规则路由；失败后依次尝试 **其他 provider（同类优先，云端→本地→云端）**
        - manual：选中的本地失败 → 只在**本地之间**互备，**绝不碰云端**

        ⚠️ 2026-10-11 修正：原先只在「云端失败」时回退到本地，方向是单向的。
        于是当 auto 规则把涉密文本路由到 local_dsh、而该 dsh 服务没启动时，
        两个本地都连不上 → 直接把 WinError 10061 抛给用户，云端明明是通的却用不上。
        实测报错：`HTTP 502: upstream connect failed: 由于目标计算机积极拒绝，无法连接。
        (os error 10061)`。现改为**先试同类、再跨类兜底**。
        """
        provider = self.decide(mode, manual_provider, route_text or message)
        try:
            return provider, provider.chat([{"role": "user", "content": message}], json_mode=json_mode)
        except Exception as first_err:
            if mode == "manual":
                # 手动模式铁律：只在本地之间互备，绝不自动改用云端
                others = [p for p in self.providers.values()
                          if p.kind == "local" and p is not provider]
                last = first_err
                for p in others:
                    try:
                        return p, p.chat([{"role": "user", "content": message}], json_mode=json_mode)
                    except Exception as e:
                        last = e
                raise last
            # auto 模式：按「先同类、再另一类」兜底，最终都会尝试到云端
            locals_ = [p for p in self.providers.values() if p.kind == "local" and p is not provider]
            clouds_ = [p for p in self.providers.values() if p.kind == "cloud" and p is not provider]
            order = (locals_ if provider.kind == "local" else clouds_) + \
                    (clouds_ if provider.kind == "local" else locals_)
            last = first_err
            for p in order:
                try:
                    return p, p.chat([{"role": "user", "content": message}], json_mode=json_mode)
                except Exception as e:
                    last = e
            # 全都失败：把**首选**的错误抛出（更贴近用户预期），并附上后续尝试的失败原因
            hint = _fallback_hint(provider, order)
            raise type(first_err)(str(first_err) + hint) if _is_simple(first_err) else first_err

    def health(self):
        out = {}
        for name, p in self.providers.items():
            out[name] = {"kind": p.kind, "model": p.model, "healthy": p.healthy()}
        return out
