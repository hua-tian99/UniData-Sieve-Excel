from pathlib import Path
from typing import Generator
import zipfile
import tempfile
import shutil
import os
from loguru import logger

def _decode_zip_filename(raw: bytes) -> str:
    """
    按 UTF-8 → GBK → CP437 顺序尝试解码压缩包内的文件名。
    全部失败时使用 CP437 + replace 错误处理，确保不抛异常。
    """
    for encoding in ("utf-8", "gbk", "cp437"):
        try:
            return raw.decode(encoding)
        except (UnicodeDecodeError, AttributeError):
            continue
    return raw.decode("cp437", errors="replace")

def _get_raw_filename_bytes(zip_info: zipfile.ZipInfo) -> bytes:
    """
    从 ZipInfo 对象中恢复原始的文件名 bytes 数据。
    """
    if zip_info.flag_bits & 0x800:
        # bit 11 is set, it means the filename is UTF-8 encoded
        return zip_info.filename.encode('utf-8')
    else:
        # bit 11 is not set, Python's zipfile arbitrarily decodes it as CP437
        # we can recover the original bytes by encoding it back to cp437
        try:
            return zip_info.filename.encode('cp437')
        except UnicodeEncodeError:
            # Fallback for unexpected cases
            return zip_info.filename.encode('utf-8')

def extract_all_excels(
    source: Path,
    temp_dir: Path,
    max_depth: int = 5
) -> Generator[Path, None, None]:
    """
    递归解压 source（.zip 或目录），yield 每个 .xlsx/.xls 的临时路径。
    每个文件名通过 _decode_zip_filename() 处理，保证中文不乱码。
    超过 max_depth 层时记录 WARNING 并跳过，不抛异常。
    损坏的 zip 记录 ERROR/WARNING 后继续处理其他文件。
    """
    if max_depth < 0:
        logger.warning(f"已达到最大解压深度，跳过: {source}")
        return

    # 预创建 temp 目录
    temp_dir.mkdir(parents=True, exist_ok=True)

    if source.is_dir():
        for item in source.iterdir():
            if item.is_file():
                if item.suffix.lower() in ('.xlsx', '.xls'):
                    logger.info(f"  提取到表格: {item.name}")
                    yield item
                elif item.suffix.lower() == '.zip':
                    logger.info(f"  发现嵌套压缩包: {item.name}，深入解压层级 {max_depth - 1}")
                    yield from extract_all_excels(item, temp_dir, max_depth - 1)
            elif item.is_dir():
                yield from extract_all_excels(item, temp_dir, max_depth)
        return

    if source.is_file() and source.suffix.lower() == '.zip':
        try:
            with zipfile.ZipFile(source, 'r') as zf:
                for zip_info in zf.infolist():
                    if zip_info.is_dir():
                        continue
                    
                    raw_bytes = _get_raw_filename_bytes(zip_info)
                    decoded_filename = _decode_zip_filename(raw_bytes)
                    
                    # 仅关心我们支持的后缀
                    if decoded_filename.lower().endswith(('.xlsx', '.xls', '.zip')):
                        # 规避同名文件覆盖：为每个解压操作分配一个独立的随机子目录
                        safe_name = os.path.basename(decoded_filename.replace('\\', '/'))
                        out_dir = Path(tempfile.mkdtemp(dir=temp_dir))
                        out_path = out_dir / safe_name
                        
                        try:
                            with zf.open(zip_info) as source_f, open(out_path, 'wb') as target_f:
                                shutil.copyfileobj(source_f, target_f)
                                
                            if out_path.suffix.lower() in ('.xlsx', '.xls'):
                                logger.info(f"  提取到表: {safe_name}")
                                yield out_path
                            elif out_path.suffix.lower() == '.zip':
                                logger.info(f"  发现内嵌压缩包: {safe_name}，深入层级 {max_depth - 1}")
                                yield from extract_all_excels(out_path, temp_dir, max_depth - 1)
                        except Exception as extract_err:
                            logger.error(f"解压文件 {decoded_filename} 失败: {extract_err}")
                            continue

        except zipfile.BadZipFile:
            logger.warning(f"压缩包已损坏或不是标准的 zip 文件，跳过: {source}")
            return
        except Exception as e:
            logger.error(f"处理压缩包时发生意外错误 {source}: {e}")
            return
