import re
import os
import logging
from datetime import datetime
from configparser import ConfigParser

# 전역 config 불러오기 (필요시 cli/에서 config.ini 경로를 지정할 수도 있음)
config = ConfigParser()
config.read(os.path.join(os.path.dirname(__file__), '..', 'config', 'config.ini'))

def parse_excel_date(date_str, formats=None):
    """
    주어진 날짜 문자열을 여러 포맷으로 파싱하여 date 객체를 반환합니다.
    실패 시 "-"를 리턴합니다.
    """
    if not date_str or date_str.strip() in ["-", "N"]:
        return "-"
    date_str = date_str.strip()
    # 기본 날짜 포맷 (필요시 config 또는 인자로 전달)
    if formats is None:
        formats = ["%d-%b-%y", "%d-%b-%Y", "%d-%B-%Y", "%Y/%m/%d", "%m/%d/%Y", "%d/%m/%Y"]
    for fmt in formats:
        try:
            return datetime.strptime(date_str, fmt).date()
        except ValueError:
            continue
    logging.warning(f"⚠️ 날짜 변환 실패: {date_str}")
    return "-"

def rename_file_with_date(original_file, date_format=None):
    """
    파일명을 변경합니다. (마지막 한글 문자 뒤에 _yyyy_mm_dd_update 문자열 삽입)
    """
    if date_format is None:
        date_format = config.get('DEFAULT', 'date_format', fallback="%Y_%m_%d")
    base, ext = os.path.splitext(original_file)
    # 마지막 한글 문자 위치 찾기
    matches = list(re.finditer(r"[가-힣]", base))
    date_str = datetime.now().strftime(date_format)
    if matches:
        last_match = matches[-1]
        new_base = base[:last_match.end()] + f"_{date_str}_update"
    else:
        new_base = base + f"_{date_str}_update"
    new_name = new_base + ext
    try:
        os.rename(original_file, new_name)
    except Exception as e:
        logging.error(f"파일명 변경 실패: {original_file} -> {new_name} / {e}")
        raise
    return new_name

def normalize(text):
    """
    텍스트에서 개행문자, 공백 제거 후 반환 (None인 경우 빈 문자열 반환)
    """
    return str(text).replace("\n", "").replace(" ", "").strip() if text is not None else ""

def get_logger(name=__name__, level=logging.INFO):
    """
    기본 로깅 설정을 반환하는 함수.
    """
    logger = logging.getLogger(name)
    if not logger.handlers:
        # 핸들러가 없을 경우 기본 설정 추가 (콘솔 출력)
        logger.setLevel(level)
        handler = logging.StreamHandler()
        formatter = logging.Formatter(
            fmt='[%(levelname)s] %(asctime)s - %(name)s: %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        handler.setFormatter(formatter)
        logger.addHandler(handler)
    return logger
