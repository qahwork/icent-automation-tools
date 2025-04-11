import os
import re
import pandas as pd
import logging
from core.utils import normalize, get_logger

logger = get_logger("serial_extractor")

def extract_serials_from_excel(file_path):
    """
    Excel 파일에서 '시리얼' 열에 해당하는 값을 추출합니다.
    """
    try:
        # 지정된 시트(기본적으로 config나 여기서 수정) 로드
        df = pd.read_excel(file_path, sheet_name="유지보수 대상장비")
    except Exception as e:
        logger.error(f"[오류] '{file_path}' 처리 중 예외 발생: {e}")
        return []

    serial_col = None
    for col in df.columns:
        if '시리얼' in str(col):
            serial_col = col
            break
    if not serial_col:
        logger.warning(f"[경고] '{file_path}'에서 '시리얼' 열을 찾을 수 없습니다.")
        return []

    serials = []
    for serial in df[serial_col].dropna().astype(str).tolist():
        cleaned = re.sub(r'\s+', '', serial.strip())
        if cleaned:
            serials.append(cleaned)
    return serials

def main():
    current_path = os.getcwd()
    serial_records = []
    for filename in os.listdir(current_path):
        if filename.endswith('.xlsx'):
            file_path = os.path.join(current_path, filename)
            serials = extract_serials_from_excel(file_path)
            for s in serials:
                serial_records.append({'Serial': s, 'Filename': filename})

    # 중복 제거
    serial_df = pd.DataFrame(serial_records).drop_duplicates()
    csv_path = os.path.join(current_path, 'serials.csv')
    serial_df.to_csv(csv_path, index=False, encoding="utf-8-sig")
    txt_path = os.path.join(current_path, 'serials.txt')
    with open(txt_path, 'w', encoding="utf-8") as f:
        f.write(','.join(serial_df['Serial'].tolist()))
    logger.info(f"✅ 작업 완료!\n시리얼 개수: {len(serial_df)}개\nCSV 저장 위치: {csv_path}\nTXT 저장 위치: {txt_path}")

if __name__ == '__main__':
    main()
