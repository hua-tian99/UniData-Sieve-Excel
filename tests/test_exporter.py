"""
tests/test_exporter.py

Phase 2 Task 6c：测试 FULL 模式、COMMON 模式及优先列排序。
"""
import pytest
import pandas as pd
from pathlib import Path

from engine.exporter import to_xlsx
from config.schema import ColumnMode

FIXTURE_DIR = Path(__file__).parent / "fixtures"


def _make_multi_source_df() -> pd.DataFrame:
    """
    模拟两个来源文件 outer join 后的 DataFrame：
    - a.xlsx 有列: 姓名, 成绩, 班级
    - b.xlsx 有列: 姓名, 成绩, 备注
    共有列（COMMON）: 姓名, 成绩
    """
    return pd.DataFrame([
        {"source_file": "a.xlsx", "source_sheet": "S1", "姓名": "张三", "成绩": 90, "班级": "A", "备注": None},
        {"source_file": "b.xlsx", "source_sheet": "S1", "姓名": "李四", "成绩": 85, "班级": None, "备注": "良好"},
    ])


# ──────────────────────────────────────────────────────────────
# FULL 模式
# ──────────────────────────────────────────────────────────────

def test_full_mode_keeps_all_columns(tmp_path):
    df = _make_multi_source_df()
    out = tmp_path / "full.xlsx"
    to_xlsx(df, out, column_mode=ColumnMode.FULL)
    result = pd.read_excel(out)
    # 所有列都应保留
    for col in ["source_file", "source_sheet", "姓名", "成绩", "班级", "备注"]:
        assert col in result.columns


# ──────────────────────────────────────────────────────────────
# COMMON 模式
# ──────────────────────────────────────────────────────────────

def test_common_mode_keeps_only_shared_columns(tmp_path):
    df = _make_multi_source_df()
    out = tmp_path / "common.xlsx"
    to_xlsx(df, out, column_mode=ColumnMode.COMMON)
    result = pd.read_excel(out)
    # 非共有列应被移除
    assert "班级" not in result.columns
    assert "备注" not in result.columns
    # 共有列应存在
    assert "姓名" in result.columns
    assert "成绩" in result.columns
    # meta 列始终保留
    assert "source_file" in result.columns
    assert "source_sheet" in result.columns


def test_common_mode_result_cols_le_any_source(tmp_path):
    """COMMON 模式输出列数 ≤ 任意单个来源文件的列数（DoD 标准）。"""
    df = _make_multi_source_df()
    out = tmp_path / "common.xlsx"
    to_xlsx(df, out, column_mode=ColumnMode.COMMON)
    result = pd.read_excel(out)

    # 任意单个来源的列数
    non_meta = [c for c in df.columns if c not in ("source_file", "source_sheet")]
    max_per_source = max(
        df[df.source_file == src][non_meta].notna().any().sum()
        for src in df.source_file.unique()
    )
    result_non_meta = [c for c in result.columns if c not in ("source_file", "source_sheet")]
    assert len(result_non_meta) <= max_per_source


def test_single_source_common_equals_full(tmp_path):
    """只有一个来源文件时，COMMON 模式等同于 FULL。"""
    df = pd.DataFrame([
        {"source_file": "a.xlsx", "source_sheet": "S1", "姓名": "张三", "成绩": 90, "班级": "A"},
    ])
    out_common = tmp_path / "common.xlsx"
    out_full   = tmp_path / "full.xlsx"
    to_xlsx(df, out_common, column_mode=ColumnMode.COMMON)
    to_xlsx(df, out_full,   column_mode=ColumnMode.FULL)
    assert sorted(pd.read_excel(out_common).columns) == sorted(pd.read_excel(out_full).columns)


# ──────────────────────────────────────────────────────────────
# 优先列排序
# ──────────────────────────────────────────────────────────────

def test_priority_keywords_columns_sorted_first(tmp_path):
    df = _make_multi_source_df()
    out = tmp_path / "priority.xlsx"
    to_xlsx(df, out, column_mode=ColumnMode.FULL, priority_keywords=["成绩"])
    result = pd.read_excel(out)
    non_meta = [c for c in result.columns if c not in ("source_file", "source_sheet")]
    assert non_meta[0] == "成绩", f"期望 '成绩' 在最前，实际顺序: {non_meta}"
