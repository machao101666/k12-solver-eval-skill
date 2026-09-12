"""集中式凭证与外部服务配置。

设计目标：
- 所有敏感信息（API Token / Key / Secret）一律从环境变量读取，源码中不含任何真实密钥，
  因此本包可以安全打包分发（Comate / Cursor / Codex 等），使用者只需填写自己的凭证。
- 支持在“包根目录”（本文件所在目录）放置 `.env` 文件集中填写全部凭证；
  真正的环境变量（export / 系统注入）优先级更高，不会被 .env 覆盖。
- API 地址、BOS Endpoint/Bucket、小度 Host 等非敏感配置保留默认值，
  仍可按需通过同名环境变量覆盖。

使用方式（各 stage 脚本内）：
    import config
    api_key = config.require("HUAYAN_API_KEY", "华研评分/解题接口")
    url = config.get("HUAYAN_API_URL", "https://www.huayanapi.com/v1/chat/completions")
"""

import os

# 本文件所在目录即“包根目录”，.env 固定从这里读取，与运行时 CWD 无关。
_PKG_DIR = os.path.dirname(os.path.abspath(__file__))
_ENV_LOADED = False


def load_env(dotenv_path=None):
    """读取包根目录下的 .env 文件到环境变量。

    - 只在首次调用时真正加载，后续调用直接返回（幂等）。
    - 使用 setdefault：已存在的真实环境变量优先，不会被 .env 覆盖。
    - 零依赖：不需要 python-dotenv，便于在各类工具环境中直接运行。
    """
    global _ENV_LOADED
    if _ENV_LOADED:
        return
    path = dotenv_path or os.path.join(_PKG_DIR, ".env")
    if os.path.isfile(path):
        with open(path, "r", encoding="utf-8") as f:
            for raw in f:
                line = raw.strip()
                if not line or line.startswith("#") or "=" not in line:
                    continue
                key, _, val = line.partition("=")
                key = key.strip()
                val = val.strip()
                # 去除包裹的成对引号
                if len(val) >= 2 and val[0] == val[-1] and val[0] in ("'", '"'):
                    val = val[1:-1]
                if key:
                    os.environ.setdefault(key, val)
    _ENV_LOADED = True


# 模块导入时即加载 .env，保证每个 stage 子进程都能拿到集中配置。
load_env()


def get(name, default=None):
    """读取配置项；不存在时返回 default。用于非敏感、可有默认值的配置。"""
    val = os.environ.get(name)
    if val is None or val == "":
        return default
    return val


def require(name, service=""):
    """读取必填凭证；缺失时给出可读的中文报错并退出。"""
    val = os.environ.get(name, "")
    if val:
        val = val.strip()
    if not val:
        hint = f"（用于：{service}）" if service else ""
        raise SystemExit(
            f"[配置缺失] 缺少环境变量 {name}{hint}。\n"
            f"请在包根目录 {os.path.join(_PKG_DIR, '.env')} 中填写该项，"
            f"或用 `export {name}=...` 设置后重试。\n"
            f"可参考模板文件 .env.example。"
        )
    return val
