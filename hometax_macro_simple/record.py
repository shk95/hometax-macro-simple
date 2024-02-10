import logging
import re
from typing import Iterator, Optional

from pandas import DataFrame, Series

from hometax_macro_simple.exception import InvalidDataException


# 데이터중 한 행을 레코드로 관리.
class Record:
    def __init__(self, dataframe: DataFrame):
        self._df_generator: Iterator[Series] = _df_row_generator(dataframe)
        self._current_series: Optional[Series] = None  # 현재 행의 데이터프레임
        self._data: dict[str:str] = {
            "name": "",
            "personal_id": "",
            "start_date": "",
            "end_date": "",
            "salary": "",
            "income_tax": "",
            "local_income_tax": "",
            "national_pension": "",
            "health_insurance": "",
            "employment_insurance": "",
        }

    # 레코드의 현재 작업 행을 반환.
    def get_current_data(self) -> dict[str:str]:
        return self._data

    # 원본 데이터프레임의 현재 행을 반환.
    def get_current_series(self) -> Series:
        return self._current_series

    # 다음 행을 레코드에 초기화. 레코드는 초기 비어있는 상태로 시작. 데이터 검증실패시 False 와 함께 데이터 반환.
    def next(self) -> bool:
        try:
            self._current_series = next(self._df_generator)
            self._data["name"] = str(self._current_series[0])
            self._data["personal_id"] = str(int(self._current_series[1]))
            self._data["start_date"] = str(int(self._current_series[2]))
            self._data["end_date"] = str(int(self._current_series[3]))
            self._data["salary"] = str(int(self._current_series[5]))
            self._data["income_tax"] = str(int(self._current_series[9]))
            self._data["local_income_tax"] = str(int(self._current_series[10]))
            self._data["national_pension"] = str(int(self._current_series[12]))
            self._data["health_insurance"] = str(int(self._current_series[17]))
            self._data["employment_insurance"] = str(int(self._current_series[18]))
            _validate(self._data)
        except StopIteration:  # next 단계에서 발생. 더이상 행이 없음.
            return False
        except InvalidDataException as e:  # 데이터 검증 실패.
            logging.info(e.args[0])
            raise e
        except Exception as e:  # 캐스팅 오류 등.
            logging.info(e.args[0])
            raise InvalidDataException(f"검증오류: 행을 읽는중 오류 발생. (값 캐스팅 오류 등). : [{e}]")
        return True

    # 계속 근로자.
    def is_ongoing(self) -> bool:
        return str(self._data["end_date"])[4:] == "1231"

    # 남성 : 세대주.
    def is_male(self) -> bool:
        return str(self._data["personal_id"])[6] in ["1", "3", "5", "7"]

    # 부녀자 세액공제(근로소득 연 3천만원 이하인 경우).
    def is_woman_deduction_eligible(self) -> bool:
        if self.is_male():
            return False
        return int(self._data["salary"]) < 41470589


def _df_row_generator(dataframe: DataFrame) -> Iterator[Series]:
    for _, row in dataframe.iterrows():
        yield row


def _validate_name(name: str) -> None:
    if not re.match(r"^[가-힣]+$", name):
        raise InvalidDataException(f"검증오류: 이름은 한글만 포함해야 합니다. [{name}]")


def _validate_personal_id(personal_id: str) -> None:
    if not re.match(r"^\d{13}$", personal_id):
        raise InvalidDataException(f"검증오류: 개인 식별번호는 13자리 숫자여야 합니다. [{personal_id}]")


def _validate_date(date: str) -> None:
    if not re.match(r"^\d{8}$", date):
        raise InvalidDataException(f"검증오류: 날짜는 8자리 숫자여야 합니다. [{date}]")


def _validate_number(value: str, field_name) -> None:
    if not value.isdigit():
        raise InvalidDataException(f"검증오류: {field_name}은(는) 숫자만 포함해야 하며 양의 정수여야 합니다. [{value}]")


def _validate_salary(value: str) -> None:
    # 숫자이며 0보다 큰 값. 이전 단계에서 문자->숫자 변환 하였기 때문에 캐스팅 오류는 없음.
    if not (value.isdigit() and int(value) >= 1):
        raise InvalidDataException(f"검증오류: 급여는 1 보다 커야 합니다. [{value}]")


def _validate(data) -> None:
    try:
        _validate_name(data["name"])
        _validate_personal_id(data["personal_id"])
        _validate_date(data["start_date"])
        _validate_date(data["end_date"])
        _validate_salary(data["salary"])
        _validate_number(data["income_tax"], "income_tax")
        _validate_number(data["local_income_tax"], "local_income_tax")
        _validate_number(data["national_pension"], "national_pension")
        _validate_number(data["health_insurance"], "health_insurance")
        _validate_number(data["employment_insurance"], "employment_insurance")
    except InvalidDataException as e:  # 잘못된 데이터
        raise e
