import pandas as pd
import pytest
from cli.excel_updater import build_serial_map

def test_build_serial_map():
    # LineDetails CSV 데이터를 모사하는 테스트 DataFrame 생성
    data = {
        "PAK/Serial Number": ["A123", "B456", "A123"],
        "Status": ["ACTIVE", "SIGNED", "ACTIVE"],
        "Last Date of Support": ["10-Apr-2025", "invalid", "20-May-2025"],
        "Service Level/Offer Type": ["Standard", "", "Premium"],
        "End Date": ["15-Apr-2025", "16-Apr-2025", "17-May-2025"],
        "Description": ["", "meraki test", ""],
        "Product /Offer Name": ["ModelX", "ModelY", "ModelZ"],
        "Subscription ID/Contract Number": ["123", "456", "789"]
    }
    df = pd.DataFrame(data)
    serial_map = build_serial_map(df)
    
    # 키 "A123"와 "B456"가 serial_map에 포함되어 있는지 확인
    assert "A123" in serial_map
    assert "B456" in serial_map
    
    # "B456"의 경우 description에 meraki가 포함되어 있으므로 LDoS만 업데이트 되어야 함
    assert serial_map["B456"]["LDoS"] != "-", "For meraki row, LDoS should be updated if valid"
    
    # "A123"의 계약번호가 숫자로 변환되었는지 확인
    assert isinstance(serial_map["A123"]["계약번호"], int)
