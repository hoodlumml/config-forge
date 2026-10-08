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
            raise RuntimeError(f"HTTP {e.code}: {body or e.reason}") from None
        return out["choices"][0]["message"]["content"]

    def healthy(self):
        try:
            self.chat([{"role": "user", "content": "ping"}], max_tokens=1)
            return True
        except Exception:
            return False


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
        """选择并发起调用；失败时本地互备回退。

        - auto：云端失败 → 逐个尝试本地（受 fallback_to_local 开关控制）
        - manual：选中的本地失败 → 逐个尝试其他本地（绝不碰云端）
        route_text：路由判定用的文本（auto 关键词规则），缺省用 message 本身。
        json_mode：规划类调用传 True，强制模型输出合法 JSON 对象（见 chat()）。
        """
        provider = self.decide(mode, manual_provider, route_text or message)
        try:
            return provider, provider.chat([{"role": "user", "content": message}], json_mode=json_mode)
        except Exception as first_err:
            if mode == "auto" and not self.fallback_to_local:
                raise
            others = [p for p in self.providers.values()
                      if p.kind == "local" and p is not provider]
            last = first_err
            for p in others:
                try:
                    return p, p.chat([{"role": "user", "content": message}], json_mode=json_mode)
                except Exception as e:
                    last = e
            raise last

    def health(self):
        out = {}
        for name, p in self.providers.items():
            out[name] = {"kind": p.kind, "model": p.model, "healthy": p.healthy()}
        return out
