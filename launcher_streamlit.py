import os
import sys
from pathlib import Path


def resource_path(name: str) -> str:
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return str(base / name)


def load_runtime_env(exe_dir: Path) -> None:
    try:
        from dotenv import load_dotenv
    except Exception:
        return

    candidates = [
        exe_dir / ".env",
        Path(resource_path(".env")),
    ]
    seen = set()
    for path in candidates:
        resolved = str(path.resolve()) if path.exists() else ""
        if not resolved or resolved in seen:
            continue
        seen.add(resolved)
        load_dotenv(dotenv_path=path, override=False)

    if os.getenv("LLM_API_KEY") and not os.getenv("OPENAI_API_KEY"):
        os.environ["OPENAI_API_KEY"] = os.getenv("LLM_API_KEY", "")
    if os.getenv("LLM_API_BASE") and not os.getenv("OPENAI_API_BASE"):
        os.environ["OPENAI_API_BASE"] = os.getenv("LLM_API_BASE", "")


def main() -> None:
    exe_dir = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
    os.chdir(exe_dir)
    load_runtime_env(exe_dir)

    app_path = resource_path("app.py")
    if not os.path.exists(app_path):
        raise FileNotFoundError(f"???????: {app_path}")

    import streamlit.config as st_config
    import streamlit.web.bootstrap as bootstrap

    flag_options = {
        "global.developmentMode": False,
        "server.headless": False,
        "server.fileWatcherType": "none",
        "browser.gatherUsageStats": False,
        "logger.level": "info",
    }

    bootstrap.load_config_options(flag_options)
    for key, value in flag_options.items():
        st_config.set_option(key, value)

    bootstrap.run(app_path, False, [], flag_options)


if __name__ == "__main__":
    main()
