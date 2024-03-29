import logging
import os
import platform
import sys
import time
from enum import Enum
from typing import Optional

import browsers
from selenium import webdriver
from selenium.common import NoSuchElementException
from selenium.webdriver import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.wait import WebDriverWait
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium.webdriver.support import expected_conditions as ec

from hometax_macro_simple.exception import InvalidDataException

_SITE_URL: str = "https://www.hometax.go.kr"
_USER_AGENT: str = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36 Edg/121.0.0.0"
_TARGET_BROWSER: str = "msedge"


class InputID(Enum):
    NAME: str = "edtIeNm"
    PERSONAL_ID: str = "edtNtplTxprDscmNoEncCntn"
    HEAD_OF_HOUSEHOLD: str = "cmbHshrClCd"
    CONTINUES_TO_WORK_Y: str = "cmbYrsClCd_input_0"
    CONTINUES_TO_WORK_N: str = "cmbYrsClCd_input_1"
    STEP_1_START_DATE: str = "edtAttrYrStrtDt_input"
    STEP_1_END_DATE: str = "edtAttrYrEndDt_input"
    STEP_1_SALARY: str = "edtSnwAmt"
    STEP_1_INCOME_TAX: str = "edtClusInctxPpmTxamt"
    STEP_1_LOCAL_INCOME_TAX: str = "edtClusRestxPpmTxamt"
    STEP_2_WOMEN_DEDUCTION: str = "cmbWmnDdcClCd_input_0"
    STEP_2_HEALTH_INSURANCE: str = "edtNtsEtMateHife"
    STEP_2_EMPLOYMENT_INSURANCE: str = "edtNtsEtMateEmpInfee"
    STEP_3_NATIONAL_PENSION: str = "edtNpInfeeUseAmt"


class ButtonID(Enum):
    CHECK_PERSONAL_ID: str = "trigger49"
    STEP_1_NEXT: str = "trigger70"
    STEP_1_CONFIRM: str = "trigger88"
    STEP_2_CONFIRM: str = "trigger102"
    FINAL_1_CONFIRM: str = "trigger125"  # 재계산하기
    FINAL_2_CONFIRM: str = "trigger57"  # 추가하기
    RESET: str = "trigger68"


class ElementID(Enum):
    WORKING_PAGE_IFRAME: str = "txppIframe"


class WebDriverManager:
    def __init__(self):
        self._webdriver: Optional[_EdgeDriver] = None

    def create(self) -> None:
        if self._webdriver is None:
            self._webdriver = _EdgeDriver()

    def close(self) -> None:
        if self._webdriver is not None:
            self._webdriver.driver.quit()
            self._webdriver = None

    def control(self):
        return _Control(self._webdriver.driver)


class _Control:
    def __init__(self, driver: webdriver.Edge):
        self._driver: webdriver.Edge = driver

    # 홈텍스 사이트 열기
    def open(self) -> None:
        self._driver.get(_SITE_URL)

    # 메크로를 시작할 페이지가 맞는지 확인
    def is_working_page(self) -> bool:
        self._driver.switch_to.default_content()
        self._driver.execute_script("window.scrollTo(0, 0);")
        try:
            self._driver.find_element(By.ID, ElementID.WORKING_PAGE_IFRAME.value)
        except NoSuchElementException as e:
            logging.info(f"메크로 시작 페이지가 아닙니다. [{e.msg}]")
            return False
        except Exception as e:
            logging.info(f"알수없는 오류가 발생했습니다. [{e}]")
            return False
        return True

    # 메크로 작업할 영역으로 이동
    def switch_to_working_page(self) -> None:
        self._driver.switch_to.default_content()
        self._driver.switch_to.frame(ElementID.WORKING_PAGE_IFRAME.value)

    def set_name(self, name: str) -> None:
        _set_input_value(self._driver, InputID.NAME.value, name)

    def set_personal_id(self, personal_id: str) -> None:
        _set_input_value(self._driver, InputID.PERSONAL_ID.value, personal_id)
        self._driver.find_element(By.ID, ButtonID.CHECK_PERSONAL_ID.value).click()

        # 주민등록번호 확인창
        WebDriverWait(self._driver, 4).until(ec.alert_is_present())
        alert = self._driver.switch_to.alert
        alert_message = alert.text
        logging.info(f"{alert_message}")
        alert.dismiss()

        self.switch_to_working_page()
        ok_message = "확인완료되었습니다."
        # error_message = "주민등록번호를 확인 해주세요."
        if not alert_message.startswith(ok_message):
            self.reset()
            # 현재 오류가난 행 반환, 해당 행의 매크로 종료
            raise InvalidDataException(f"Invalid personal ID [{personal_id}]")

    def set_head_of_household(self, head_of_household: bool) -> None:
        select = Select(self._driver.find_element(By.ID, InputID.HEAD_OF_HOUSEHOLD.value))
        if head_of_household:
            select.select_by_visible_text("세대주")
        else:
            select.select_by_visible_text("세대원")

    def set_continues_to_work(self, continues_to_work: bool) -> None:
        if continues_to_work:
            self._driver.find_element(By.ID, InputID.CONTINUES_TO_WORK_Y.value).click()
        else:
            self._driver.find_element(By.ID, InputID.CONTINUES_TO_WORK_N.value).click()

    # 근무처별 소득명세 단계
    def next_step_1(self) -> None:
        self._driver.find_element(By.ID, ButtonID.STEP_1_NEXT.value).click()
        time.sleep(1)

    def set_step_1_start_date(self, start_date: str) -> None:
        _set_input_value(self._driver, InputID.STEP_1_START_DATE.value, start_date)

    def set_step_1_end_date(self, end_date: str) -> None:
        _set_input_value(self._driver, InputID.STEP_1_END_DATE.value, end_date)

    def set_step_1_salary(self, salary: str) -> None:
        _set_input_value(self._driver, InputID.STEP_1_SALARY.value, salary)

    def set_step_1_income_tax(self, income_tax: str) -> None:
        _set_input_value(self._driver, InputID.STEP_1_INCOME_TAX.value, income_tax)

    def set_step_1_local_income_tax(self, local_income_tax: str) -> None:
        _set_input_value(self._driver, InputID.STEP_1_LOCAL_INCOME_TAX.value, local_income_tax)

    def confirm_step_1(self) -> None:
        self._driver.find_element(By.ID, ButtonID.STEP_1_CONFIRM.value).click()

        confirm = self._driver.switch_to.alert
        confirm.accept()

        self.switch_to_working_page()
        time.sleep(1)

    def set_step_2_woman_deduction(self, eligible: bool) -> None:
        if not eligible:
            return

        check = self._driver.find_element(By.ID, InputID.STEP_2_WOMEN_DEDUCTION.value)
        if not check.is_selected():
            check.click()
        time.sleep(1)

    def set_step_2_health_insurance(self, health_insurance: str) -> None:
        _set_input_value(self._driver, InputID.STEP_2_HEALTH_INSURANCE.value, health_insurance)

    def set_step_2_employment_insurance(self, employment_insurance: str) -> None:
        _set_input_value(self._driver, InputID.STEP_2_EMPLOYMENT_INSURANCE.value, employment_insurance)

    def confirm_step_2(self) -> None:
        self._driver.find_element(By.ID, ButtonID.STEP_2_CONFIRM.value).click()
        confirm = self._driver.switch_to.alert
        confirm.accept()

        self.switch_to_working_page()
        time.sleep(1)

    def set_step_3_national_pension(self, national_pension: str) -> None:
        _set_input_value(self._driver, InputID.STEP_3_NATIONAL_PENSION.value, national_pension)

    def confirm_final_step(self) -> None:
        self._driver.find_element(By.ID, ButtonID.FINAL_1_CONFIRM.value).click()

        # 재계산 실행
        WebDriverWait(self._driver, 10).until(ec.alert_is_present())
        confirm = self._driver.switch_to.alert
        confirm_message = confirm.text
        logging.info(f"계산하기 단계 : [{confirm_message}]")
        confirm.dismiss()
        self.switch_to_working_page()
        time.sleep(1)

        # 재계산 실행후 완료 alert 창
        message_ok_1 = "재계산이 완료되었습니다."
        if not confirm_message.startswith(message_ok_1):
            logging.info(f"재계산 실패. [{confirm_message}]")
            self.reset()
            raise InvalidDataException("웹드라이버: 재계산 실패.")

        # 추가하기 confirm 창
        self._driver.find_element(By.ID, ButtonID.FINAL_2_CONFIRM.value).click()
        WebDriverWait(self._driver, 10).until(ec.alert_is_present())
        submit = self._driver.switch_to.alert
        submit.accept()

        # 추가하기 confirm 창 확인
        WebDriverWait(self._driver, 12).until(ec.alert_is_present())
        alert_dialog = self._driver.switch_to.alert
        alert_message = alert_dialog.text
        logging.info(f"추가하기 단계 : [{alert_message}]")
        alert_dialog.accept()

        message_ok_2 = "처리가 완료되었습니다"
        message_duplicated = "기존 수록자료가 존재합니다"
        if alert_message.startswith(message_ok_2):
            logging.info(f"추가하기 성공.")
            self.switch_to_working_page()
            return
        elif alert_message.startswith(message_duplicated):
            logging.info(f"데이터 중복입력. 무시됨.")
            self.reset()
            return
        else:
            logging.info(f"추가하기 실패. [{alert_message}]")
            self.reset()
            raise InvalidDataException("웹드라이버: 입력 오류 발생. 추가하기 실패.")

    # 작성내역 초기화
    def reset(self) -> None:  # 메인콘텐츠로 전환 -> 맨 위로 스크롤 -> 작업페이지로 전환 -> 초기화 버튼 클릭
        self._driver.switch_to.default_content()
        self._driver.execute_script("window.scrollTo(0, 0);")
        self.switch_to_working_page()
        reset_button = self._driver.find_element(By.ID, ButtonID.RESET.value)
        reset_button.click()
        time.sleep(1)


class _EdgeDriver:
    def __init__(self):
        options = webdriver.EdgeOptions()
        options.add_argument(f"--user-data-dir={_create_profile_path()}")
        options.add_argument(f"--user-agent={_USER_AGENT}")
        options.add_argument("--disable-extensions")
        options.add_argument("--disable-extensions-file-access-check")
        options.add_argument("--disable-extensions-http-throttling")
        options.add_argument("--disable-infobars")
        options.add_argument("--disable-automation")
        options.add_experimental_option("detach", True)  # keep the browser open after the process has ended
        options.add_experimental_option('excludeSwitches', ['enable-logging'])
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option("useAutomationExtension", False)

        self.driver: webdriver = webdriver.Edge(service=EdgeService(EdgeChromiumDriverManager().install()),
                                                options=options)
        self.driver.implicitly_wait(7)


def _set_input_value(driver: webdriver.Edge, input_id: str, value: str) -> None:
    _clear_input(driver, input_id)
    logging.info(f"Set value: [{value}]")
    driver.find_element(By.ID, input_id).send_keys(value)


def _clear_input(driver: webdriver.Edge, input_id: str) -> None:
    logging.info(f"Clear input. input_id : [{input_id}]")
    control_key = Keys.COMMAND if platform.system() == "Darwin" else Keys.CONTROL
    input_element = driver.find_element(By.ID, input_id)
    input_element.click()
    input_element.send_keys(control_key + "a")
    input_element.send_keys(Keys.DELETE)


def _create_profile_path(profile_name="Macro") -> str:
    # 운영 체제별 기본 프로필 경로.
    if os.name == 'nt':  # Windows.
        base_path = os.path.join(os.environ['USERPROFILE'], 'AppData', 'Local', 'Microsoft', 'Edge', 'User Data')
    elif os.name == 'posix':
        if 'darwin' in sys.platform:  # MacOS.
            base_path = os.path.join(os.path.expanduser('~'), 'Library', 'Application Support', 'Microsoft Edge')
        else:  # Linux.
            base_path = os.path.join(os.path.expanduser('~'), '.config', 'microsoft-edge')
    else:
        raise Exception("Unsupported operating system.")

    # 프로필 경로 생성.
    profile_path = os.path.join(base_path, profile_name)
    if not os.path.exists(profile_path):
        os.makedirs(profile_path)  # 프로필 디렉토리가 없으면 생성.

    return profile_path


# 엣지 브라우저 지원
def is_supported() -> tuple[bool, Optional[str]]:
    version = browsers.get(_TARGET_BROWSER)

    if version is None:
        logging.info(f"Unsupported browser: {_TARGET_BROWSER} required. Aborting...")
        return False, None
    else:
        return True, version.get("version")


if __name__ == "__main__":
    pass
