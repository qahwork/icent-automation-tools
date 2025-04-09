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
## 🚀 가상환경 구성

### 1. Windows 기준
```bash
python -m venv venv
venv\Scripts\activate
```
2. 의존성 설치
```bash
pip install -r requirements.txt
```
