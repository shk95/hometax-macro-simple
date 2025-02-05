import logging
import time

import pandas
from pandas import DataFrame

from hometax_macro_simple.exception import InvalidDataException
from hometax_macro_simple.record import Record
from hometax_macro_simple.webdriver import WebDriverManager

# 근소로득 지급명세서 제출용 (적용 내용)
#
# 0  : 성명                        name
# 1  : 주민등록번호                  personal_id
# 2  : 시작일자                     start_date
# 3  : 종료일자                     end_date
# ## : 총급여
# 5  : 급여                        salary
# ## : 상여
# ## : 인정상여
# ## : 주식매수선택권 행사이익
# 9  : 소득세                      income_tax
# 10 : 지방소득세                   local_income_tax
# ## : 농어촌특별세
# 12 : 국민연금보험료                 national_pension
# ## : 공무원연금
# ## : 군인연금
# ## : 사립학교직원연금
# ## : 별정우체국연금
# 17 : 건강보험료                    health_insurance
# 18 : 고용보험료                    employment_insurance
# ## : 법정기부금
# ## : 종교단체 지정기부금
# ## : 종교단체 외 지정기부금


_COLUMN_NAMES: list[str] = [
    "성명",
    "주민등록번호",
    "시작일자",
    "종료일자",
    "총급여",
    "급여",
    "상여",
    "인정상여",
    "주식매수선택권 행사이익",
    "소득세",
    "지방소득세",
    "농어촌특별세",
    "국민연금보험료",
    "공무원연금",
    "군인연금",
    "사립학교직원연금",
    "별정우체국연금",
    "건강보험료",
    "고용보험료",
    "법정기부금",
    "종교단체 지정기부금",
    "종교단체 외 지정기부금"
]


class Macro:
    def __init__(self, webdriver: WebDriverManager, path: str, selected_sheet_name: str):
        self.webdriver: WebDriverManager = webdriver
        self.record: Record = Record(
            pandas.read_excel(path, header=None, skiprows=6, sheet_name=selected_sheet_name))
        self.error_data_list: list[pandas.Series] = []

    def start(self) -> None:
        time.sleep(1)
        self.webdriver.control().switch_to_default_content()

        # 홈텍스 근로소득 지급명세서 제출 페이지에서 작업페이지를 확인.
        if not self.webdriver.control().is_working_page():
            # FIXME: 예외를 발생시켜서 오류파일을 만들지않도록
            logging.info("잘못된 시작위치 입니다.")
            return

        logging.info("메크로 반복 시작")
        i = 0
        while True:
            time.sleep(2)
            try:
                self.webdriver.control().reset()
                i += 1
                logging.info(f"메크로 반복 [{i}].")
                # 레코드의 다음 행 가져오기. 레코드의 다음 행이 없으면 반복 종료.
                if not self.record.next():
                    break
                logging.info(f"메크로 반복 [{i}].\n현재 데이터 : {self.record.get_current_data()}")

                # 소득자 인적사항 단계
                self.webdriver.control().set_name(self.record.get_current_data()["name"])
                self.webdriver.control().set_personal_id(
                    self.record.get_current_data()["personal_id"])  # 주민등록번호 검증 실패 가능성.
                self.webdriver.control().set_head_of_household(self.record.is_male())
                self.webdriver.control().set_continues_to_work(self.record.is_ongoing())

                # 근무처별 소득명세 -> 주(현) 단계
                self.webdriver.control().next_step_1()
                self.webdriver.control().set_step_1_start_date(self.record.get_current_data()["start_date"])
                self.webdriver.control().set_step_1_end_date(self.record.get_current_data()["end_date"])
                self.webdriver.control().set_step_1_salary(self.record.get_current_data()["salary"])
                self.webdriver.control().set_step_1_income_tax(self.record.get_current_data()["income_tax"])
                self.webdriver.control().set_step_1_local_income_tax(self.record.get_current_data()["local_income_tax"])
                self.webdriver.control().confirm_step_1()

                # 소득ㆍ세액공제명세 단계
                self.webdriver.control().set_step_2_woman_deduction(self.record.is_woman_deduction_eligible())
                self.webdriver.control().set_step_2_health_insurance(self.record.get_current_data()["health_insurance"])
                self.webdriver.control().set_step_2_employment_insurance(
                    self.record.get_current_data()["employment_insurance"])
                self.webdriver.control().confirm_step_2()

                # 연금보험료 공제 단계 -> 국민연금
                self.webdriver.control().set_step_3_national_pension(self.record.get_current_data()["national_pension"])

                # 계산 및 추가 단계
                self.webdriver.control().confirm_final_step()

            except InvalidDataException as e:
                logging.info(f"메크로 실행 중 데이터 오류 발생. 다음 순서로 넘김. : [{e}]")
                self.error_data_list.append(self.record.get_current_series())
                continue
            except Exception as e:
                logging.info(f"메크로 실행 중 오류 발생. 다음 순서로 넘김. : [{e}]")
                self.error_data_list.append(self.record.get_current_series())
                continue

    # 에러가 발생한 데이터 확인.
    def get_error_dataframe(self) -> DataFrame:
        df_default = DataFrame(columns=_COLUMN_NAMES)
        if not self.error_data_list:
            logging.info(f"에러 데이터 없음")
            return df_default

        df_error = DataFrame(self.error_data_list)

        df_default_columns = list(df_default.columns)
        df_error_columns = list(df_error.columns)
        if len(df_error_columns) > len(df_default_columns):
            df_error.columns = df_default_columns + df_error_columns[len(df_default_columns):]
        else:
            df_error.columns = df_default_columns[:len(df_error_columns)]

        df_concat = pandas.concat([df_default, df_error], ignore_index=True)
        logging.info(f"에러 데이터 반환 : {df_concat}")
        return df_concat


# test
if __name__ == "__main__":
    pass
