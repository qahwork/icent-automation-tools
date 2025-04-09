import os
import pandas as pd
from openpyxl import load_workbook
from datetime import datetime

def parse_excel_date(date_str):
    if not date_str or date_str.strip() in ["-", "N"]:
        return None

    date_str = date_str.strip()
    formats = ["%d-%b-%y", "%d-%b-%Y", "%d-%B-%Y", "%Y/%m/%d", "%m/%d/%Y", "%d/%m/%Y"]
    for fmt in formats:
        try:
            return datetime.strptime(date_str, fmt).date()
        except ValueError:
            continue
    print(f"⚠️ 날짜 변환 실패: {date_str}")
    return None

csv_path = "LineDetails.csv"
if not os.path.exists(csv_path):
    raise FileNotFoundError(f"CSV 파일 '{csv_path}'을 찾을 수 없습니다.")
csv_df = pd.read_csv(csv_path, dtype=str).fillna("")

serial_map = {}
for _, row in csv_df.iterrows():
    serial = row.get("PAK/Serial Number", "").strip()
    status = row.get("Status", "").strip().upper()
    if not serial:
        continue

    entry = serial_map.setdefault(serial, {
        "LDoS": "-",
        "서비스 종류": "-",
        "ACTIVE 종료일": "-",
        "SIGNED 종료일": "-",
        "모델명PID": "-"
    })

    ldos_raw = row.get("Last Date of Support", "").strip()
    entry["LDoS"] = parse_excel_date(ldos_raw) or "-"

    offer_type = row.get("Service Level/Offer Type", "").strip()
    if offer_type:
        entry["서비스 종류"] = offer_type

    end_date_raw = row.get("End Date", "").strip()
    end_date_parsed = parse_excel_date(end_date_raw) or "-"
    if status == "ACTIVE":
        entry["ACTIVE 종료일"] = end_date_parsed
    elif status == "SIGNED":
        entry["SIGNED 종료일"] = end_date_parsed

    model_name = row.get("Product /Offer Name", "").strip()
    if model_name:
        entry["모델명PID"] = model_name

def normalize(text):
    return str(text).replace("\n", "").replace(" ", "").strip()

column_candidates = {
    "serial_col": ["시리얼"],
    "ldos_date_col": ["H/WLDoSDate", "H/WEOS(Support)날짜"],
    "service_type_col": ["서비스종류Subscription/ServiceLevel", "서비스종류"],
    "end_date_active_col": ["서비스종료일(Active)"],
    "end_date_signed_col": ["서비스종료일(SIGNED)"],
    "model_pid_col": ["모델명PID", "모델명\nPID"]
}

for filename in os.listdir():
    if not filename.endswith(".xlsx"):
        continue

    try:
        wb = load_workbook(filename)
        if "유지보수 대상장비" not in wb.sheetnames:
            print(f"[SKIP] 시트 없음: {filename}")
            continue

        ws = wb["유지보수 대상장비"]
        headers = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
        header_map = {normalize(h): i for i, h in enumerate(headers)}

        col_indices = {}
        for key, candidates in column_candidates.items():
            found = None
            for c in candidates:
                if normalize(c) in header_map:
                    found = header_map[normalize(c)]
                    break
            col_indices[key] = found

        if None in col_indices.values():
            print(f"\n[❌ 헤더 매핑 실패] {filename}")
            print("엑셀 헤더 목록:")
            for i, h in enumerate(headers):
                print(f"  - [{i}] '{h}' → 정규화: '{normalize(h)}'")
            for k, v in col_indices.items():
                if v is None:
                    print(f"  ❌ '{k}'에 해당하는 열을 찾을 수 없음 (후보: {column_candidates[k]})")
            continue

        updated = 0
        debug_printed = False

        for row in ws.iter_rows(min_row=2):
            serial = str(row[col_indices["serial_col"]].value).strip()
            if not serial or serial not in serial_map:
                continue

            info = serial_map[serial]

            if not debug_printed:
                print(f"\n🔍 디버그 (1회) - {serial} → {info}")
                debug_printed = True

            row[col_indices["ldos_date_col"]].value = info["LDoS"] if info["LDoS"] != "-" else "-"
            row[col_indices["service_type_col"]].value = info["서비스 종류"] if info["서비스 종류"] else "-"
            row[col_indices["end_date_active_col"]].value = info["ACTIVE 종료일"] if info["ACTIVE 종료일"] != "-" else "-"
            row[col_indices["end_date_signed_col"]].value = info["SIGNED 종료일"] if info["SIGNED 종료일"] != "-" else "-"
            row[col_indices["model_pid_col"]].value = info["모델명PID"] if info["모델명PID"] else "-"
            updated += 1

        wb.save(filename)
        print(f"✅ [UPDATED] {filename} - 총 {updated}건 업데이트 완료")

    except Exception as e:
        print(f"[ERROR] {filename} 처리 중 오류 발생: {e}")
