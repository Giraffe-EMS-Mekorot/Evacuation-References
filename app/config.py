"""Configuration and paths for the tteudot-agent pipeline."""
import os
from pathlib import Path

from dotenv import load_dotenv

load_dotenv()


def _secret_or_env(key: str) -> str:
    """Prefers st.secrets (how Streamlit Community Cloud injects secrets),
    falls back to an environment variable (.env locally via load_dotenv()
    above) - lets this module work unmodified for both the cloud deployment
    and local CLI/dev use.

    Wrapped in a broad try/except, not just KeyError: touching st.secrets at
    all raises when no secrets.toml exists anywhere (the normal case for a
    plain local CLI run with only a .env file), and that's exactly the case
    that should fall through quietly rather than surface as an error here.
    """
    try:
        import streamlit as st

        if key in st.secrets:
            return st.secrets[key]
    except Exception:
        pass
    return os.environ.get(key)


ANTHROPIC_API_KEY = _secret_or_env("ANTHROPIC_API_KEY")
# Vision-capable model used to read certificates. Override with CLAUDE_MODEL in .env if needed.
MODEL_NAME = os.environ.get("CLAUDE_MODEL", "claude-sonnet-5")

BASE_DIR = Path(__file__).resolve().parent.parent
INPUT_DIR = BASE_DIR / "input"
OUTPUT_DIR = BASE_DIR / "output"
OUTPUT_FILENAME = "ריכוז_תעודות.xlsx"

SUPPORTED_EXTENSIONS = {".pdf", ".jpg", ".jpeg", ".png"}
