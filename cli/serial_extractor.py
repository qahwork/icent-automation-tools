import os
import re
import pandas as pd

# 현재 작업 디렉토리
current_path = os.getcwd()
serial_records = []

# 현재 폴더의 모든 .xlsx 파일 조회
for filename in os.listdir(current_path):
    if filename.endswith('.xlsx'):
        file_path = os.path.join(current_path, filename)
        try:
            # '유지보수 대상장비' 시트만 로드
            df = pd.read_excel(file_path, sheet_name='유지보수 대상장비')

            # 시리얼 열 찾기
            serial_col = None
            for col in df.columns:
                if '시리얼' in str(col):
                    serial_col = col
                    break

            if serial_col:
                raw_serials = df[serial_col].dropna().astype(str).tolist()
                
                # 공백 제거 및 문자만 필터링
                for serial in raw_serials:
                    cleaned = serial.strip()
                    cleaned = re.sub(r'\s+', '', cleaned)
                    if cleaned:
                        serial_records.append({'Serial': cleaned, 'Filename': filename})
            else:
                print(f"[경고] '{filename}'에서 '시리얼' 열을 찾을 수 없습니다.")

        except Exception as e:
            print(f"[오류] '{filename}' 처리 중 예외 발생: {e}")

# 중복 제거 (동일한 시리얼 + 파일명 쌍만 허용)
serial_df = pd.DataFrame(serial_records).drop_duplicates()

# CSV 저장
csv_path = os.path.join(current_path, 'serials.csv')
serial_df.to_csv(csv_path, index=False, encoding='utf-8-sig')

# TXT 저장 (쉼표로 구분된 시리얼만 추출)
txt_path = os.path.join(current_path, 'serials.txt')
with open(txt_path, 'w', encoding='utf-8') as f:
    f.write(','.join(serial_df['Serial'].tolist()))

print(f"\n✅ 작업 완료!")
print(f"🔹 시리얼 개수: {len(serial_df)}개")
print(f"📄 CSV 저장 위치: {csv_path}")
print(f"📄 TXT 저장 위치: {txt_path}")
