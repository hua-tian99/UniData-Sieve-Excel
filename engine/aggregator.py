"""
engine/aggregator.py

核心变更（Phase 2 Task 6a）：超出阈值时切换流式写入，防止 OOM。
流式写入使用 openpyxl 直接 append 行，不依赖 pd.ExcelWriter 的 startrow hack。
"""
import pandas as pd
from typing import Iterable
from pathlib import Path
from loguru import logger

STREAM_THRESHOLD = 50_000  # 超过此行数切换为流式写入


def aggregate(
    records: Iterable[dict],
    output_path: Path,
    stream_threshold: int = STREAM_THRESHOLD
) -> tuple:
    """
    汇集 scan_excel 的所有 yield。

    正常模式（记录数 ≤ stream_threshold）：
      - 构建 DataFrame，source_file / source_sheet 列置于最左侧。
      - 列对齐采用 outer join，缺失值填 NaN。
      - 返回 (DataFrame, 行数)。

    流式模式（记录数 > stream_threshold）：
      - 跳过 DataFrame 构建，直接使用 openpyxl append 逐行写入 output_path。
      - 返回 (None, 行数)，调用方据此判断是否有 DataFrame 可预览。
      - 记录 WARNING 日志提示用户已切换流式模式。
    """
    META_COLS = ["source_file", "source_sheet"]

    buffer: list[dict] = []
    stream_mode = False
    total_rows = 0

    # 流式模式状态（延迟初始化，仅超阈值时才 import openpyxl）
    _wb = None
    _ws = None
    _headers = None  # 流式模式下的列顺序（首次写入时锁定）

    for record in records:
        total_rows += 1
        buffer.append(record)

        # 检查是否需要切换流式模式
        if not stream_mode and len(buffer) > stream_threshold:
            stream_mode = True
            logger.warning(
                f"记录数超过阈值 {stream_threshold}，已切换为流式写入模式: {output_path}"
            )
            import openpyxl
            output_path.parent.mkdir(parents=True, exist_ok=True)
            _wb = openpyxl.Workbook(write_only=True)
            _ws = _wb.create_sheet()

        # 流式模式：把 buffer 全部 flush 到 openpyxl
        if stream_mode and buffer:
            for row_dict in buffer:
                if _headers is None:
                    # 第一次写入：确定列顺序，meta 列置左
                    all_keys = list(row_dict.keys())
                    _headers = (
                        [c for c in META_COLS if c in all_keys]
                        + [c for c in all_keys if c not in META_COLS]
                    )
                    _ws.append(_headers)

                # 补齐缺失列后写入一行
                row_values = [_normalize_for_stream(row_dict.get(h)) for h in _headers]
                _ws.append(row_values)
            buffer = []

    # 收尾
    if stream_mode:
        # flush 剩余（不应有，但防御性处理）
        for row_dict in buffer:
            if _headers is None:
                all_keys = list(row_dict.keys())
                _headers = (
                    [c for c in META_COLS if c in all_keys]
                    + [c for c in all_keys if c not in META_COLS]
                )
                _ws.append(_headers)
            _ws.append([_normalize_for_stream(row_dict.get(h)) for h in _headers])

        if _wb is not None:
            _wb.save(str(output_path))
            _wb.close()
        return (None, total_rows)

    # 正常模式
    if not buffer:
        logger.warning("没有任何匹配记录，输出空 DataFrame。")
        return (pd.DataFrame(columns=META_COLS), 0)

    df = pd.DataFrame(buffer)
    df = _reorder_meta_cols(df, META_COLS)
    return (df, total_rows)


def _normalize_for_stream(value) -> object:
    """
    流式写入时的单元格值处理：保留 datetime 对象（openpyxl 可直接写入），
    将 NaN/None 转为 None（openpyxl 写空单元格），其余原样返回。
    """
    if value is None:
        return None
    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass
    return value


def _reorder_meta_cols(df: pd.DataFrame, meta_cols: list) -> pd.DataFrame:
    """将 meta_cols 中存在的列强制移到最左侧，其余列顺序不变。"""
    existing_meta = [c for c in meta_cols if c in df.columns]
    other_cols = [c for c in df.columns if c not in meta_cols]
    return df[existing_meta + other_cols]
