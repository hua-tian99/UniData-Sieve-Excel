"""
engine/__init__.py

统一入口：run_pipeline 串联
  extract_all_excels → scan_excel → aggregate → to_xlsx
cli.py 和 app.py 均调用此函数，不直接调用子模块。
"""
from pathlib import Path
import tempfile
import contextlib
from loguru import logger

from engine.processor import extract_all_excels
from engine.excel_handler import scan_excel
from engine.aggregator import aggregate
from engine.exporter import to_xlsx


def run_pipeline(source: Path, config) -> tuple:
    """
    串联 extract_all_excels → scan_excel → aggregate → to_xlsx。
    cli.py 和 app.py 均调用此函数，不直接调用子模块。
    返回 (输出文件路径, 匹配行数)。

    Args:
        source:  输入的 .zip 文件或目录路径。
        config:  AppConfig 实例（来自 config.schema）。
    """
    from engine.matcher import MatchRule, MatchMode
    from config.schema import ColumnMode

    output_path = Path(config.output_filename)

    # 构建 MatchRule
    mode_map = {"and": MatchMode.AND, "or": MatchMode.OR, "regex": MatchMode.REGEX}
    match_mode = mode_map.get(config.match_rule.mode.lower(), MatchMode.AND)
    rule = MatchRule(
        keywords=config.match_rule.keywords,
        mode=match_mode,
        pattern=config.match_rule.pattern,
    )

    logger.info(f"开始处理: {source}")
    logger.info(f"匹配模式: {match_mode.value} | 关键词: {rule.keywords}")

    # CLI 模式下使用 contextlib.ExitStack 管理临时目录，确保异常时自动清理
    # Windows 上 pandas 可能持有文件句柄；使用 ignore_cleanup_errors 避免清理失败时崩溃
    with contextlib.ExitStack() as stack:
        tmp = stack.enter_context(
            tempfile.TemporaryDirectory(dir=config.temp_dir, ignore_cleanup_errors=True)
        )
        temp_dir = Path(tmp)

        # 1. 解压 → 迭代所有 Excel 路径
        excel_paths = list(
            extract_all_excels(source, temp_dir, max_depth=config.max_depth)
        )
        logger.info(f"共找到 {len(excel_paths)} 个 Excel 文件")

        # 2. 扫描 → 将所有匹配记录在 with 块内完全 materialize，
        #    确保 pandas ExcelFile 文件句柄在 TemporaryDirectory 清理前已释放
        all_records = []
        for excel_path in excel_paths:
            logger.debug(f"  扫描: {excel_path.name}")
            all_records.extend(scan_excel(excel_path, rule))

    # ExitStack 已关闭，临时目录已清理（或忽略清理错误）
    # 3. 聚合
    df, total_rows = aggregate(
        iter(all_records),
        output_path,
        stream_threshold=config.stream_threshold,
    )

    # 4. 导出（仅正常模式有 df；流式模式 aggregate 内部已完成写入）
    if df is not None:
        to_xlsx(
            df,
            output_path,
            column_mode=config.column_mode,
            priority_keywords=config.match_rule.keywords or None,
        )
    else:
        logger.info("流式模式：文件已由 aggregate 直接写入，跳过 to_xlsx。")

    logger.info(f"完成！共匹配 {total_rows} 行 -> {output_path}")
    return (output_path, total_rows)
