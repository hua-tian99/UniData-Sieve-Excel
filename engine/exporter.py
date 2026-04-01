"""
engine/exporter.py

Task 6c：实现 COMMON 列模式与优先列排序。
COMMON 模式：仅保留所有来源文件（source_file 分组）中均出现的列（交集）。
优先列：含有 priority_keywords 中任一关键词的列名排在 meta 列之后的最前面。

# EXTENSION: CSV export
"""
import pandas as pd
from pathlib import Path
from loguru import logger

_FULL   = "full"
_COMMON = "common"


def to_xlsx(
    df: pd.DataFrame,
    output_path: Path,
    column_mode,
    priority_keywords: list = None
) -> None:
    """
    将 df 写出为 .xlsx。

    column_mode = FULL：保留所有列（outer join 默认行为）。
    column_mode = COMMON：仅保留在所有来源文件中均出现的列（inner join 列集合）。
    无论哪种模式，含有 priority_keywords 中任一关键词的列名，
    均排在 source_file / source_sheet 之后的最前面。

    # EXTENSION: CSV export
    """
    META_COLS = ["source_file", "source_sheet"]
    mode_str = column_mode.value if hasattr(column_mode, "value") else str(column_mode)

    # ── COMMON 列模式过滤 ────────────────────────────────────────
    if mode_str == _COMMON:
        df = _apply_common_mode(df, META_COLS)

    # ── 优先列排序 ───────────────────────────────────────────────
    if priority_keywords:
        df = _apply_priority_sort(df, META_COLS, priority_keywords)

    # ── 写入 ────────────────────────────────────────────────────
    output_path.parent.mkdir(parents=True, exist_ok=True)
    df.to_excel(output_path, index=False, engine="openpyxl")
    logger.info(f"已导出 {len(df)} 行 x {len(df.columns)} 列 -> {output_path}")


def _apply_common_mode(df: pd.DataFrame, meta_cols: list) -> pd.DataFrame:
    """
    仅保留在所有来源文件中均出现的列（非全 NaN 的列集合的交集）。
    meta 列（source_file / source_sheet）始终保留。
    """
    if "source_file" not in df.columns or df["source_file"].nunique() <= 1:
        # 只有一个来源，COMMON == FULL，无需过滤
        logger.info("COMMON 模式：只有单一来源文件，等同于 FULL 模式。")
        return df

    # 对每个 source_file 分组，统计哪些（非 meta）列含有至少一个非空值
    non_meta_cols = [c for c in df.columns if c not in meta_cols]

    per_file_non_empty: list[set] = []
    for _, group in df.groupby("source_file"):
        non_empty = {
            col for col in non_meta_cols
            if group[col].notna().any()
        }
        per_file_non_empty.append(non_empty)

    if not per_file_non_empty:
        return df

    common_cols = set.intersection(*per_file_non_empty)
    meta_present = [c for c in meta_cols if c in df.columns]
    keep = meta_present + [c for c in non_meta_cols if c in common_cols]

    removed = len(df.columns) - len(keep)
    logger.info(f"COMMON 模式：保留 {len(keep)} 列，已移除 {removed} 列非共有列。")
    return df[keep]


def _apply_priority_sort(
    df: pd.DataFrame,
    meta_cols: list,
    priority_keywords: list,
) -> pd.DataFrame:
    """将含有 priority_keywords 关键词的列名移到 meta 列之后的最前面。"""
    meta_present = [c for c in meta_cols if c in df.columns]
    non_meta = [c for c in df.columns if c not in meta_cols]

    priority = [
        c for c in non_meta
        if any(kw.lower() in c.lower() for kw in priority_keywords)
    ]
    rest = [c for c in non_meta if c not in priority]
    return df[meta_present + priority + rest]
