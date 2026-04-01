"""config/schema.py — Pydantic AppConfig, 严格对应计划书第 4 节。"""
from pydantic import BaseModel, Field
from enum import Enum
from typing import Optional
import json
from pathlib import Path


class ColumnMode(str, Enum):
    FULL   = "full"    # 保留所有列
    COMMON = "common"  # 仅保留共同列


class MatchRule(BaseModel):
    keywords:  list = []
    mode:      str = "and"       # "and" | "or" | "regex"
    pattern:   Optional[str] = None


class AppConfig(BaseModel):
    match_rule:       MatchRule
    column_mode:      ColumnMode = ColumnMode.FULL
    max_depth:        int = Field(5, ge=1, le=10)
    output_filename:  str = "result.xlsx"
    log_level:        str = "INFO"
    temp_dir:         str = "temp"
    logs_dir:         str = "logs"
    stream_threshold: int = Field(50_000, ge=1000)


def load_config(path: str = None) -> AppConfig:
    """从 path 加载 JSON，path=None 时使用 default_config.json。"""
    if path is None:
        path = Path(__file__).parent / "default_config.json"
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)
    return AppConfig(**data)
