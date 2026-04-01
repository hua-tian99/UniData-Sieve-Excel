"""
tests/test_excel_handler.py
———————————————————————————————
测试 excel_handler 的 _normalize_cell / _row_to_str / scan_excel
其中 test_normalize_cell_verbose 会逐值打印"预处理前 → 后"的变化（类断点演示）
"""
import datetime
from pathlib import Path
import tempfile
import shutil
import pytest
import pandas as pd
import openpyxl

from engine.excel_handler import _normalize_cell, _row_to_str, scan_excel
from engine.matcher import MatchRule, MatchMode


# ──────────────────────────────────────────────────────────────
# 1.  _normalize_cell 参数化测试（覆盖率 100%）
# ──────────────────────────────────────────────────────────────

NORMALIZE_CASES = [
    # (input_value, expected_output, description)
    (None,                                  "",          "None → 空字符串"),
    (float("nan"),                          "",          "NaN  → 空字符串"),
    (datetime.datetime(2026, 3, 31, 8, 0),  "2026/3/31", "datetime with time → YYYY/M/D"),
    (datetime.datetime(2026, 1, 5, 0, 0),   "2026/1/5",  "datetime 单位月日 → 不补零"),
    (pd.Timestamp("2026-03-31"),            "2026/3/31", "pd.Timestamp → YYYY/M/D"),
    (datetime.date(2026, 3, 31),            "2026/3/31", "date → YYYY/M/D"),
    (123.0,                                 "123",       "整数浮点 123.0 → '123'"),
    (0.0,                                   "0",         "整数浮点 0.0   → '0'"),
    (3.14,                                  "3.14",      "真浮点 3.14 保留原样"),
    ("张三",                                 "张三",      "字符串原样返回"),
    (42,                                    "42",        "整数 → str"),
    ("",                                    "",          "空字符串原样返回"),
]


@pytest.mark.parametrize("value, expected, desc", NORMALIZE_CASES)
def test_normalize_cell(value, expected, desc):
    result = _normalize_cell(value)
    assert result == expected, f"[{desc}] got {result!r}, want {expected!r}"


# ──────────────────────────────────────────────────────────────
# 2.  "断点演示"测试 —— 打印预处理前后对比（-s 可见输出）
# ──────────────────────────────────────────────────────────────

def test_normalize_cell_verbose():
    """
    仿断点测试：逐值打印 _normalize_cell 预处理前后的字符串变化。
    在 pytest -s 模式下可直接观察每一条转换结果，无需调试器。
    """
    print("\n" + "=" * 60)
    print(f"{'原始值 (repr)':<35}  {'raw type':<20}  →  预处理结果")
    print("=" * 60)
    for value, expected, desc in NORMALIZE_CASES:
        result = _normalize_cell(value)
        match_mark = "OK" if result == expected else "NG"
        print(f"  {repr(value):<33}  {type(value).__name__:<20}  ->  {result!r:<15}  {match_mark}  ({desc})")
    print("=" * 60)


# ──────────────────────────────────────────────────────────────
# 3.  _row_to_str
# ──────────────────────────────────────────────────────────────

def test_row_to_str():
    row = pd.Series({
        "姓名": "张三",
        "成绩": 95.0,
        "日期": datetime.datetime(2026, 3, 31),
        "备注": None,
    })
    result = _row_to_str(row)
    assert "张三" in result
    assert "95" in result
    assert "2026/3/31" in result
    assert result.count("  ") == 0  # 不应有双空格（None → ""，空格连续）


# ──────────────────────────────────────────────────────────────
# 4.  scan_excel —— 端到端，含 source_file / source_sheet 验证
# ──────────────────────────────────────────────────────────────

@pytest.fixture
def sample_xlsx(tmp_path):
    """创建一个包含目标行和非目标行的临时 xlsx 文件。"""
    filepath = tmp_path / "sample.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["姓名", "成绩", "备注"])            # header
    ws.append(["张三", 95.0, "优秀"])              # 含 "张三" → 应命中
    ws.append(["李四", 82.0, "良好"])              # 不含 "张三" → 不命中
    ws.append(["王五", 77.5, "张三推荐"])           # 备注含 "张三" → 应命中
    wb.save(filepath)
    return filepath


def test_scan_excel_and_mode(sample_xlsx):
    rule = MatchRule(keywords=["张三"], mode=MatchMode.AND)
    results = list(scan_excel(sample_xlsx, rule))
    assert len(results) == 2

    for r in results:
        assert r["source_file"] == "sample.xlsx"
        assert r["source_sheet"] == "Sheet1"


def test_scan_excel_or_mode(sample_xlsx):
    rule = MatchRule(keywords=["张三", "李四"], mode=MatchMode.OR)
    results = list(scan_excel(sample_xlsx, rule))
    assert len(results) == 3  # 所有行都含其中一个关键词


def test_scan_excel_regex_mode(sample_xlsx):
    rule = MatchRule(keywords=[], mode=MatchMode.REGEX, pattern=r"\d{2}\.\d")
    results = list(scan_excel(sample_xlsx, rule))
    # 77.5 → "77.5" 可匹配 \d{2}\.\d；82.0 → "82" 不匹配；95.0 → "95" 不匹配
    assert len(results) == 1


def test_scan_excel_empty_sheet(tmp_path):
    """空 Sheet 不应 yield 任何结果，且不抛出异常。"""
    filepath = tmp_path / "empty.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Empty"
    wb.save(filepath)

    rule = MatchRule(keywords=["任意"], mode=MatchMode.AND)
    results = list(scan_excel(filepath, rule))
    assert results == []
