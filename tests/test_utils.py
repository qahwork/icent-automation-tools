import os
import tempfile
from datetime import datetime

import pytest
from core.utils import parse_excel_date, normalize, rename_file_with_date, get_logger

def test_parse_excel_date_valid():
    # "10-Apr-2025"가 올바른 날짜 객체로 파싱되는지 확인
    result = parse_excel_date("10-Apr-2025")
    # datetime.date 객체이거나 문자열로 비교해도 됨.
    assert str(result) == "2025-04-10", f"Expected '2025-04-10' but got {result}"

def test_parse_excel_date_invalid():
    # 변환할 수 없는 날짜 문자열인 경우 "-" 반환
    result = parse_excel_date("invalid-date")
    assert result == "-", f"Expected '-' for invalid date but got {result}"

def test_normalize():
    # 개행문자와 공백 제거 테스트
    text = "  Hello\nWorld  "
    expected = "HelloWorld"
    result = normalize(text)
    assert result == expected, f"Expected '{expected}', got '{result}'"

def test_get_logger():
    # get_logger가 Logger 객체를 반환하는지
    logger = get_logger("test_logger")
    assert logger is not None
    # 적어도 하나 이상의 핸들러가 존재해야 함
    assert logger.handlers, "Logger should have at least one handler"

def test_rename_file_with_date():
    # 임시 폴더 내에 테스트용 파일을 생성한 후, 파일명이 올바르게 변경되는지 확인
    with tempfile.TemporaryDirectory() as tmpdir:
        original_file = os.path.join(tmpdir, "테스트파일.xlsx")
        with open(original_file, 'w') as f:
            f.write("Test content")
        # rename_file_with_date 함수를 호출하여 파일 이름 변경
        new_name = rename_file_with_date(original_file)
        # 새 파일이 존재하는지 확인
        assert os.path.exists(new_name), "Renamed file does not exist"
        # 파일명에 현재 날짜가 포함되었는지 확인 (날짜 형식은 config에 정의된 값에 따름)
        current_date = datetime.now().strftime("%Y_%m_%d")
        assert current_date in new_name, "Current date string not found in renamed file"
