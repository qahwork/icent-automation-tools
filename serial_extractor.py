import os
import pandas as pd

# 현재 작업 디렉토리
current_path = os.getcwd()
all_serials = []

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
                serials = df[serial_col].dropna().astype(str).tolist()
                all_serials.extend(serials)
            else:
                print(f"[경고] '{filename}'에서 '시리얼' 열을 찾을 수 없습니다.")

        except Exception as e:
            print(f"[오류] '{filename}' 처리 중 예외 발생: {e}")

# 중복 제거 (선택사항)
# all_serials = list(set(all_serials))

# CSV 저장
csv_path = os.path.join(current_path, 'serials.csv')
pd.DataFrame({'Serials': all_serials}).to_csv(csv_path, index=False, encoding='utf-8-sig')

# TXT 저장 (쉼표 구분)
txt_path = os.path.join(current_path, 'serials.txt')
with open(txt_path, 'w', encoding='utf-8') as f:
    f.write(','.join(all_serials))

print(f"\n✅ 작업 완료!")
print(f"🔹 시리얼 개수: {len(all_serials)}개")
print(f"📄 CSV 저장 위치: {csv_path}")
print(f"📄 TXT 저장 위치: {txt_path}")
