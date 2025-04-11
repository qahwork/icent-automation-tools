import os
import tempfile
import pandas as pd
import pytest
from cli.serial_extractor import extract_serials_from_excel

def test_extract_serials_from_excel():
    # "유지보수 대상장비" 시트를 포함하는 테스트용 Excel 데이터를 생성
    sample_data = {
        "시리얼": ["A123 ", " B456", None, " C789\n"],
        "Other Column": ["data", "more data", "data", "data"]
    }
    df = pd.DataFrame(sample_data)
    
    with tempfile.TemporaryDirectory() as tmpdir:
        file_path = os.path.join(tmpdir, "sample.xlsx")
        with pd.ExcelWriter(file_path) as writer:
            df.to_excel(writer, sheet_name="유지보수 대상장비", index=False)
        # Excel 파일에서 시리얼 정보를 추출
        serials = extract_serials_from_excel(file_path)
        # 예상되는 결과는 공백 및 개행문자가 제거된 형태여야 함
        expected = ["A123", "B456", "C789"]
        assert serials == expected, f"Expected {expected} but got {serials}"
