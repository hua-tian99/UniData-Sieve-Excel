"""
tests/conftest.py — pytest 全局配置

1. 将 loguru 重定向到 pytest 的 capfd，避免测试输出被日志污染。
2. 为 fixtures 目录提供 Path 常量，各 test 模块可直接 import 使用。
"""
import sys
from pathlib import Path
import pytest
from loguru import logger


# ── 重定向 loguru 到 stderr (pytest 默认捕获) ────────────────────
logger.remove()
logger.add(sys.stderr, level="WARNING", format="{level}: {message}")

# ── Fixture 目录常量 ─────────────────────────────────────────────
FIXTURE_DIR = Path(__file__).parent / "fixtures"


@pytest.fixture(scope="session")
def fixture_dir() -> Path:
    return FIXTURE_DIR
