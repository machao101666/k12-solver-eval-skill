"""统一配置加载层。

所有阶段脚本的接入点、模型、提示词、并发、列映射都从这里读取，
不再在代码里写死。配置来源与优先级：

1. ``config.yaml``（可用 ``AUTO_EVAL_CONFIG`` 指向别的文件）
2. 环境变量 / ``.env``（解析 yaml 里的 ``${VAR}`` 占位符，真实环境变量优先于 .env）
3. 命令行参数（各脚本自行覆盖，例如 ``--concurrency``）

典型用法::

    from config import cfg, secret, prompt

    url = cfg("providers.openai_compat.base_url")
    key = secret("providers.openai_compat.api_key")   # 缺失时报错并提示要设哪个环境变量
    tpl = prompt("solve")

密钥永远不落盘到 config.yaml，只放 .env（已被 .gitignore 忽略）。
"""

from __future__ import annotations

import os
import re
from typing import Any

REPO_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_CONFIG_PATH = os.path.join(REPO_DIR, "config.yaml")
DEFAULT_ENV_PATH = os.path.join(REPO_DIR, ".env")

_ENV_PATTERN = re.compile(r"^\$\{([A-Za-z_][A-Za-z0-9_]*)\}$")
_MISSING = object()


class MissingEnv(str):
    """未设置的环境变量占位符。

    继承 ``str`` 且值为空串，因此 ``bool()`` 为 False、误用也不会抛类型错误；
    ``secret()`` 会识别它并给出「该设哪个环境变量」的明确报错。
    """

    env_name: str

    def __new__(cls, env_name: str) -> "MissingEnv":
        obj = super().__new__(cls, "")
        obj.env_name = env_name
        return obj


def load_dotenv(path: str | None = None) -> dict[str, str]:
    """把 .env 载入 os.environ。已存在的真实环境变量不被覆盖。"""
    path = path or os.environ.get("AUTO_EVAL_ENV_FILE") or DEFAULT_ENV_PATH
    loaded: dict[str, str] = {}
    if not os.path.exists(path):
        return loaded
    with open(path, "r", encoding="utf-8") as handle:
        for raw in handle:
            line = raw.strip()
            if not line or line.startswith("#"):
                continue
            if line.startswith("export "):
                line = line[len("export "):].strip()
            if "=" not in line:
                continue
            key, value = line.split("=", 1)
            key = key.strip()
            value = value.strip()
            if len(value) >= 2 and value[0] == value[-1] and value[0] in "\"'":
                value = value[1:-1]
            if not key:
                continue
            loaded[key] = value
            os.environ.setdefault(key, value)
    return loaded


def _resolve(node: Any) -> Any:
    """递归把 ``${VAR}`` 替换成环境变量值，缺失则替换成 MissingEnv。"""
    if isinstance(node, dict):
        return {key: _resolve(value) for key, value in node.items()}
    if isinstance(node, list):
        return [_resolve(item) for item in node]
    if isinstance(node, str):
        match = _ENV_PATTERN.match(node.strip())
        if match:
            env_name = match.group(1)
            value = os.environ.get(env_name)
            return value if value not in (None, "") else MissingEnv(env_name)
    return node


def _load() -> dict[str, Any]:
    try:
        import yaml
    except ImportError as exc:  # pragma: no cover
        raise RuntimeError(
            "缺少 PyYAML，请先执行 pip3 install -r requirements.txt"
        ) from exc

    path = os.environ.get("AUTO_EVAL_CONFIG") or DEFAULT_CONFIG_PATH
    if not os.path.exists(path):
        raise RuntimeError(
            f"未找到配置文件: {path}\n"
            "请从仓库根目录的 config.yaml 复制一份，或设置 AUTO_EVAL_CONFIG 指向你的配置文件。"
        )
    load_dotenv()
    with open(path, "r", encoding="utf-8") as handle:
        data = yaml.safe_load(handle) or {}
    if not isinstance(data, dict):
        raise RuntimeError(f"配置文件顶层必须是映射: {path}")
    data["__config_path__"] = path
    return _resolve(data)


_CONFIG: dict[str, Any] | None = None
_PROMPT_CACHE: dict[str, str] = {}


def config() -> dict[str, Any]:
    """返回已解析的完整配置（首次调用时加载并缓存）。"""
    global _CONFIG
    if _CONFIG is None:
        _CONFIG = _load()
    return _CONFIG


def cfg(path: str, default: Any = _MISSING) -> Any:
    """按点号路径取配置，例如 ``cfg("runtime.score_concurrency")``。"""
    node: Any = config()
    for part in path.split("."):
        if not isinstance(node, dict) or part not in node:
            if default is _MISSING:
                raise KeyError(
                    f"配置项不存在: {path}（配置文件: {config().get('__config_path__')}）"
                )
            return default
        node = node[part]
    if isinstance(node, MissingEnv):
        return None if default is _MISSING else default
    return node


def secret(path: str) -> str:
    """取必须存在的凭证，缺失时给出可操作的报错。"""
    node: Any = config()
    for part in path.split("."):
        if not isinstance(node, dict) or part not in node:
            raise KeyError(f"配置项不存在: {path}")
        node = node[part]
    if isinstance(node, MissingEnv):
        raise RuntimeError(
            f"配置项 {path} 未就绪：环境变量 {node.env_name} 未设置。\n"
            "请复制 .env.example 为 .env 并填入真实值，或直接 export 该环境变量。"
        )
    if node in (None, ""):
        raise RuntimeError(f"配置项 {path} 为空，请在 config.yaml 或 .env 中补齐。")
    return str(node)


def soft_secret(path: str) -> str:
    """取凭证，缺失时返回空串而不报错。

    用于模块顶层赋值，避免「只是 import 一下 / 看 --help」也要求配齐密钥。
    真正发请求的入口应先调用 :func:`require_secrets` 做前置校验。
    """
    try:
        return secret(path)
    except (KeyError, RuntimeError):
        return ""


def require_secrets(*paths: str) -> None:
    """入口处一次性校验多个凭证，缺失时汇总报错。"""
    problems: list[str] = []
    for path in paths:
        try:
            secret(path)
        except (KeyError, RuntimeError) as exc:
            problems.append(f"  - {exc}".replace("\n", "\n    "))
    if problems:
        raise RuntimeError(
            "配置缺失，无法发起请求：\n" + "\n".join(problems)
            + "\n复制 .env.example 为 .env 填入真实值后，可执行 python3 config.py 自检。"
        )


def provider(name: str) -> dict[str, Any]:
    """取某个 provider 的完整配置。"""
    providers = cfg("providers")
    if name not in providers:
        raise KeyError(f"providers.{name} 未在配置中定义")
    return providers[name]


def prompt_dir() -> str:
    """提示词目录。

    ``prompts.dir`` 为相对路径时，先按「配置文件所在目录」解析，
    找不到再回退到「仓库目录」，这样把 config.yaml 放到别处也能用默认提示词。
    """
    configured = cfg("prompts.dir", "prompts")
    if os.path.isabs(configured):
        return configured
    candidates = [
        os.path.join(os.path.dirname(config()["__config_path__"]), configured),
        os.path.join(REPO_DIR, configured),
    ]
    for candidate in candidates:
        if os.path.isdir(candidate):
            return candidate
    return candidates[0]


def prompt(name: str) -> str:
    """按逻辑名读取提示词文件内容（带缓存）。

    ``name`` 取值见 config.yaml 的 ``prompts.files``：
    clean_classify / solve / consensus / score_target / score_reference
    """
    if name in _PROMPT_CACHE:
        return _PROMPT_CACHE[name]
    files = cfg("prompts.files")
    if name not in files:
        raise KeyError(
            f"未定义的提示词: {name}，可选: {', '.join(sorted(files))}"
        )
    path = os.path.join(prompt_dir(), files[name])
    if not os.path.exists(path):
        raise RuntimeError(
            f"提示词文件不存在: {path}\n"
            "检查 config.yaml 的 prompts.dir 与 prompts.files 是否正确。"
        )
    with open(path, "r", encoding="utf-8") as handle:
        text = handle.read()
    _PROMPT_CACHE[name] = text
    return text


def work_dir() -> str:
    """当前批次工作目录。run_pipeline.py 会通过 AUTO_EVAL_WORK_DIR 注入。"""
    return os.environ.get("AUTO_EVAL_WORK_DIR") or REPO_DIR


def excel_path() -> str:
    return os.path.join(work_dir(), cfg("paths.excel_file", "evaluation.xlsx"))


def column(name: str) -> int:
    """取 Excel 列号（1-based）。"""
    columns = cfg("excel.columns")
    if name not in columns:
        raise KeyError(f"excel.columns.{name} 未定义")
    return int(columns[name])


def doctor() -> int:
    """自检：打印配置来源、缺失的凭证和提示词文件状态。返回缺失项数量。"""
    data = config()
    print(f"配置文件: {data.get('__config_path__')}")
    print(f".env    : {os.environ.get('AUTO_EVAL_ENV_FILE') or DEFAULT_ENV_PATH}"
          f" ({'存在' if os.path.exists(os.environ.get('AUTO_EVAL_ENV_FILE') or DEFAULT_ENV_PATH) else '不存在'})")
    print()

    missing: list[str] = []

    def walk(prefix: str, body: dict) -> None:
        """递归打印 providers 子树；嵌套的 target_model.baidu_solve.* 也要覆盖到。"""
        for key, value in body.items():
            path = f"{prefix}.{key}"
            if isinstance(value, MissingEnv):
                missing.append(f"{path} -> ${{{value.env_name}}}")
                print(f"  ✖ {path}  未设置环境变量 {value.env_name}")
            elif isinstance(value, dict):
                walk(path, value)
            elif any(token in key for token in ("key", "secret", "password")):
                shown = str(value)
                masked = shown[:4] + "***" + shown[-2:] if len(shown) > 8 else "***"
                print(f"  ✔ {path}  {masked}")
            else:
                print(f"  ✔ {path}  {value}")

    print("[凭证与接入点]")
    _STYLE_BLOCKS = ("baidu_solve", "openai_compat")
    for provider_name, body in sorted(cfg("providers").items()):
        if not isinstance(body, dict):
            continue
        if provider_name == "target_model":
            # 待测模型下挂着多种接口风格，只检查当前 api_style 生效的那一份
            active = str(body.get("api_style", "baidu_solve")).strip().lower()
            body = {
                key: value for key, value in body.items()
                if key not in _STYLE_BLOCKS or key == active
            }
        walk(f"providers.{provider_name}", body)

    print("\n[提示词]")
    for name in sorted(cfg("prompts.files")):
        try:
            text = prompt(name)
            print(f"  ✔ {name:16s} {len(text)} 字符")
        except Exception as exc:
            missing.append(f"prompts.{name}")
            print(f"  ✖ {name:16s} {exc}")

    print("\n[模型]")
    print(f"  分类       : {cfg('models.classify.model')}")
    for item in cfg("models.reference_models"):
        print(f"  参考模型   : {item['label']:14s} {item['model']}  (provider={item['provider']})")
    print(f"  共识裁判   : {cfg('models.consensus_judge.model')}")
    print(f"  评分裁判   : {cfg('models.score_judge.model')}")

    print("\n[评测口径]")
    print(f"  只评置信度 : {cfg('evaluation.only_confidence')}")
    print(f"  达标线     : >= {cfg('evaluation.pass_score')} 分")
    print(f"  3分格式要求: {' + '.join(cfg('evaluation.required_output_headings'))}")

    if missing:
        print(f"\n共 {len(missing)} 项待配置：")
        for item in missing:
            print(f"  - {item}")
        print("参考 .env.example 补齐后重新执行本自检。")
    else:
        print("\n配置完整。")
    return len(missing)


if __name__ == "__main__":
    raise SystemExit(1 if doctor() else 0)
