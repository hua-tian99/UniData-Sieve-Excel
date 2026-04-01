"""
cli.py — 命令行入口（第一阶段实现）

用法示例：
    python cli.py --input archive.zip --keywords 张三 李四 --mode and
    python cli.py --input archive.zip --keywords "\d{4}/\d+" --mode regex --output out.xlsx
    python cli.py --input archive.zip --config my_config.json
"""
import argparse
import sys
from pathlib import Path
from loguru import logger


def _setup_logger(log_level: str, logs_dir: str) -> None:
    """配置 loguru：同时输出到控制台和日志文件。"""
    log_path = Path(logs_dir)
    log_path.mkdir(parents=True, exist_ok=True)
    logger.remove()
    logger.add(sys.stderr, level=log_level, colorize=True,
               format="<green>{time:HH:mm:ss}</green> | <level>{level:<8}</level> | {message}")
    logger.add(log_path / "run.log", level="DEBUG", encoding="utf-8",
               format="{time:YYYY-MM-DD HH:mm:ss} | {level:<8} | {message}")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="unidata-sieve",
        description="从嵌套 ZIP 中提取并聚合匹配 Excel 行",
    )
    parser.add_argument(
        "--input", "-i", required=True,
        help="输入的 .zip 文件或目录路径"
    )
    parser.add_argument(
        "--keywords", "-k", nargs="*", default=[],
        help="匹配关键词（AND/OR 模式）"
    )
    parser.add_argument(
        "--mode", "-m", choices=["and", "or", "regex"], default="and",
        help="匹配模式（默认 and）"
    )
    parser.add_argument(
        "--pattern", "-p", default=None,
        help="正则表达式（仅 regex 模式）"
    )
    parser.add_argument(
        "--output", "-o", default="result.xlsx",
        help="输出文件名（默认 result.xlsx）"
    )
    parser.add_argument(
        "--column-mode", dest="column_mode", choices=["full", "common"], default="full",
        help="列模式：full 保留所有列，common 仅保留共同列"
    )
    parser.add_argument(
        "--max-depth", dest="max_depth", type=int, default=5,
        help="最大递归解压深度（1-10，默认 5）"
    )
    parser.add_argument(
        "--config", "-c", default=None,
        help="JSON 配置文件路径（优先级低于命令行参数）"
    )
    parser.add_argument(
        "--log-level", dest="log_level", default="INFO",
        choices=["DEBUG", "INFO", "WARNING", "ERROR"],
    )
    return parser


def main(argv=None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    # ── 加载配置 ──────────────────────────────────────────────
    from config.schema import AppConfig, ColumnMode, MatchRule, load_config

    if args.config:
        config = load_config(args.config)
    else:
        config = load_config(None)  # 从 default_config.json 加载基础配置

    # 命令行参数覆盖配置（CLI 优先）
    config = config.model_copy(update={
        "match_rule": MatchRule(
            keywords=args.keywords,
            mode=args.mode,
            pattern=args.pattern,
        ),
        "column_mode":     ColumnMode(args.column_mode),
        "max_depth":       args.max_depth,
        "output_filename": args.output,
        "log_level":       args.log_level,
    })

    _setup_logger(config.log_level, config.logs_dir)

    # 确保 temp 目录存在（run_pipeline 内 ExitStack 会在其下创建临时子目录）
    Path(config.temp_dir).mkdir(parents=True, exist_ok=True)

    # ── 运行主管线 ─────────────────────────────────────────────
    from engine import run_pipeline

    source = Path(args.input)
    if not source.exists():
        logger.error(f"输入路径不存在: {source}")
        return 1

    output_path, total_rows = run_pipeline(source, config)

    if total_rows == 0:
        logger.warning("没有找到任何匹配记录，输出文件为空表。")
    else:
        logger.info(f"处理完成，共 {total_rows} 条匹配记录 -> {output_path.resolve()}")

    return 0


if __name__ == "__main__":
    sys.exit(main())
