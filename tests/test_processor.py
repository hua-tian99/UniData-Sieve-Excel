import os
import shutil
from pathlib import Path
import tempfile
import pytest

from engine.processor import extract_all_excels, _decode_zip_filename

# Fixture path
FIXTURE_DIR = Path(__file__).parent / 'fixtures'
GBK_ZIP = FIXTURE_DIR / 'gbk_named.zip'

def test_decode_zip_filename():
    # UTF-8 encoded
    utf8_bytes = "测试.xlsx".encode('utf-8')
    assert _decode_zip_filename(utf8_bytes) == "测试.xlsx"
    
    # GBK encoded
    gbk_bytes = "测试.xlsx".encode('gbk')
    assert _decode_zip_filename(gbk_bytes) == "测试.xlsx"
    
    # Invalid bytes
    invalid_bytes = b'\xff\xfe'
    decoded = _decode_zip_filename(invalid_bytes)
    assert decoded != "" 
    assert isinstance(decoded, str)

def test_extract_all_excels():
    # Setup temp directory
    temp_dir = Path(tempfile.mkdtemp())
    
    try:
        # Run extractor
        extracted = list(extract_all_excels(GBK_ZIP, temp_dir))
        
        # Check that we extracted the files
        assert len(extracted) == 2
        
        # Check if the filenames are correct (which means decoding worked)
        filenames = [path.name for path in extracted]
        assert "ascii_file.xlsx" in filenames
        assert "测试.xlsx" in filenames
        
        # Check files actually exist
        for path in extracted:
            assert path.exists()
            assert path.is_file()
            
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)
