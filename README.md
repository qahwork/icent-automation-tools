# 🔧 Serial Tools for Excel Maintenance

Cisco 유지보수 대상 장비 관리를 위한 엑셀 자동화 도구입니다.

## 📂 Scripts

### 1. `serial_extractor.py`
- 모든 `.xlsx` 파일의 `'유지보수 대상장비'` 시트에서 시리얼 번호 추출
- `serials.csv` 및 `serials.txt` 생성

### 2. `excel_updater.py`
- `LineDetails.csv` 기반으로 엑셀 파일 업데이트 (LDoS, 서비스 종류, 종료일 등)
- 시리얼 번호 기준으로 매칭
- 디버깅 메시지 1회 출력

## 🧪 설치 방법

```bash
pip install -r requirements.txt
```
## 🚀 실행 방법
bash
Copy
Edit
# 시리얼 추출
python serial_extractor.py

# 엑셀 업데이트
python excel_updater.py
yaml
Copy
Edit

