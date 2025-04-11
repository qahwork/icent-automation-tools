# 📊 Excel & CSV 데이터 업데이트 프로젝트

이 프로젝트는 최신의의 `LineDetails*.csv` 파일과 Excel 파일을 처리하여 시리얼 정보를 업데이트하고, 문제 시리얼 리스트를 생성하는 도구입니다.  

## 📚 목차
- [✨ 특징](#-특징)
- [🗂️ 프로젝트 구조](#️-프로젝트-구조)
- [🔧 설치 및 환경 설정](#-설치-및-환경-설정)
- [🚀 사용법](#-사용법)
  - [💻 명령줄 인터페이스 (CLI)](#-명령줄-인터페이스-cli)
- [⚙️ 설정 파일 구성(예정정)](#️-설정-파일-구성)
- [📈 개선 사항 및 향후 계획](#-개선-사항-및-향후-계획)
- [📞 문의 및 지원](#-문의-및-지원)

## ✨ 특징
- (개선 중)`LineDetails*.csv` 파일 중 가장 최신의 시리얼 정보를 매핑합니다.
- `serials.csv` 파일과의 비교를 통해 누락된 시리얼 정보를 기록하고 `problem_serial.csv`를 생성합니다.
- Excel 파일의 특정 시트(`유지보수 대상장비`)를 업데이트하고, 파일명을 날짜와 함께 변경합니다.
- 유연한 설정 파일을 통해 사용자 정의가 가능하며, 명령줄 인자 지원으로 다양한 실행 옵션 제공.


## 🗂️ 프로젝트 구조
```
project-root/
├── cli/                    
│   ├── excel_updater.py       # Excel 업데이트 관련 스크립트
│   ├── serial_extractor.py    # 시리얼 추출 관련 스크립트
│   └── __init__.py            
├── core/                    
│   ├── utils.py               # 날짜 파싱, 파일명 변경, 문자열 정규화 등 공통 함수
│   ├── logger.py              # 로깅 관련 모듈 (필요 시 추가)
│   └── __init__.py            
├── config/                  
│   └── config.ini             # 환경 및 컬럼 설정 등 (전체 모듈 공용)
├── docs/                    
│   └── README.md              # 프로젝트 설명 및 사용법 문서 
└── requirements.txt           # 필요한 Python 패키지 목록
```

## 🔧 설치 및 환경 설정

### 1️⃣ 환경 준비
- 프로젝트 다운로드
```bash
  git clone https://github.com/qahwork/icent-automation-tools.git
  cd icent-automation-tools
  ```

- **Python 3.7 이상**이 설치되어 있어야 합니다.
- 가상 환경 생성 및 활성화 권장:
  ```bash
  python -m venv venv
  source venv/bin/activate    # Windows: venv\Scripts\activate
  ```

### 2️⃣ 필수 패키지 설치
```bash
pip install -r requirements.txt
```

### 3️⃣ 설정 파일 수정 (예정정)
- `config.ini` 파일에서 기본 설정(예: `sheet_name`, 컬럼명 등)을 필요에 따라 수정합니다.

## 🚀 사용법

### 💻 명령줄 인터페이스 (CLI)
1. **시리얼 추출 실행:**  
   Excel 파일들의 시트 `유지보수 대상장비`에서 시리얼 정보를 추출하여 `serials.csv`와 `serials.txt` 파일을 생성합니다.
   각 파일은 다음과 같습니다.
   - serials.csv: 해당 시리얼의 위치 매핑
   - serials.txt: 해당 파일의 전체 내용을 복사후 CCW 검색하여 LineDetails.csv 생성성
   ```bash
   python serial_extractor.py --input-folder "./data"
   ```
   *--input-folder 옵션은 Excel 파일들이 있는 폴더 경로를 지정합니다.*

2. **Excel 업데이트 실행:**  
   CSV와 Excel 데이터를 기반으로 Excel 파일 내 데이터를 업데이트하고 파일명을 변경합니다.
   ```bash
   python excel_updater.py --input-folder "./data"
   ```
   실행 결과로 업데이트된 파일과 디버그 로그가 출력됩니다.


## ⚙️ 설정 파일 구성 (예정)
`config.ini` 파일 예시:
```ini
[DEFAULT]
header_skip_lines = 5
dated_format = %Y_%m_%d

[Excel]
sheet_name = 유지보수 대상장비
serial_column_candidates = 시리얼
ldos_date_candidates = H/WLDoSDate,H/WEOS(Support)날짜
service_type_candidates = 서비스종류Subscription/ServiceLevel,서비스종류
# 필요한 컬럼 후보를 추가합니다.

[CSV]
encoding = utf-8-sig
```
*필요한 설정값은 코드 내에서 활용되도록 반영합니다.*

## 📈 개선 사항 및 향후 계획
- **코드 모듈화:** 공통 기능을 `utils.py` 등으로 모듈화하여 유지보수성 개선
- **설정 파일 활용:** 하드코딩된 상수를 설정 파일에서 읽어올 수 있도록 편의성 개선
- **명령줄 옵션 확장:** `argparse`를 활용하여 다양한 실행 옵션 (예: 로그 레벨, 디버그 모드 등)을 제공합니다.

## 📞 문의 및 지원
문제가 있거나 개선 사항에 대한 제안은 ksj1304@icent.co.kr WEBEX 또는 [깃허브 이슈 트래커]를 통해 문의해주시기 바랍니다.