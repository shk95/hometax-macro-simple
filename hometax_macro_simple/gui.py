import logging
import os
import subprocess
import sys
from typing import Optional

import pandas as pd
from PySide6 import QtCore, QtWidgets
from PySide6.QtWidgets import QVBoxLayout

from hometax_macro_simple.macro import Macro
from hometax_macro_simple.webdriver import is_supported, WebDriverManager

_LOG_LEVEL = logging.INFO


# 위젯 선언
class MyWidget(QtWidgets.QWidget):
    def __init__(self, webdriver: WebDriverManager):
        super().__init__()
        self.webdriver: WebDriverManager = webdriver

        self.file_name: str = ""
        self.selected_sheet_name: str = ""
        self.exel_file: Optional[pd.ExcelFile] = None

        self.layout: QVBoxLayout = QtWidgets.QVBoxLayout(self)

        # "홈텍스 열기" 버튼
        self.open_button = QtWidgets.QPushButton("홈텍스 열기")
        self.layout.addWidget(self.open_button)
        self.open_button.clicked.connect(self.open)

        # "브라우저 확인" 버튼
        self.check_browser_button = QtWidgets.QPushButton("브라우저 확인")
        self.layout.addWidget(self.check_browser_button)
        self.check_browser_button.clicked.connect(self.check_browser)

        # 파일 관련 버튼을 위한 수평 레이아웃
        self.fileLayout = QtWidgets.QHBoxLayout()

        # "파일 불러오기" 버튼
        self.load_file_button = QtWidgets.QPushButton("파일 불러오기")
        self.fileLayout.addWidget(self.load_file_button)
        self.load_file_button.clicked.connect(self.load_file)

        # "파일 불러오기 취소" 버튼
        self.unload_file_button = QtWidgets.QPushButton("파일 제거")
        self.fileLayout.addWidget(self.unload_file_button)
        self.unload_file_button.clicked.connect(self.unload_file)

        # 파일 이름 표시 영역
        self.file_name_label = QtWidgets.QLabel("선택된 파일 없음")
        self.fileLayout.addWidget(self.file_name_label)

        # 파일 관련 버튼과 파일 이름 표시 영역을 메인 레이아웃에 추가
        self.layout.addLayout(self.fileLayout)

        # "매크로 시작" 버튼
        self.start_macro_button = QtWidgets.QPushButton("매크로 시작")
        self.layout.addWidget(self.start_macro_button)
        self.start_macro_button.clicked.connect(self.start_macro)

        # "근로소득 지급명세서 제출 바로가기" 버튼 추가
        # self.shortcut_1_button = QtWidgets.QPushButton("근로소득 지급명세서 제출 바로가기")
        # self.layout.addWidget(self.shortcut_1_button)
        # self.shortcut_1_button.clicked.connect(self.go_shortcut_1)

        # 예시 파일 열기 버튼
        self.show_example_button = QtWidgets.QPushButton("예시 파일 열기")
        self.layout.addWidget(self.show_example_button)
        self.show_example_button.clicked.connect(self.show_example)

        # 메시지 표시 영역
        self.message_label = QtWidgets.QLabel("")
        self.layout.addWidget(self.message_label)

        # 로그 출력을 위한 QPlainTextEdit 위젯 추가
        log_text_box = QTextEditLogger(self)
        self.layout.addWidget(log_text_box.widget)

        # 로거 설정
        logger = logging.getLogger()
        logger.addHandler(log_text_box)
        logger.setLevel(_LOG_LEVEL)

        # 테스트 로깅 메시지 발생
        logging.info("프로그램이 시작되었습니다.")

        self.check_browser()

    @QtCore.Slot()
    def open(self):
        self.webdriver.create()
        self.webdriver.control().open()

    @QtCore.Slot()
    def check_browser(self):
        status, version = is_supported()
        if status:
            self.message_label.setText(f"브라우저 확인 : 지원됨 (버전: {version})")
            logging.info(f"브라우저 확인 : 지원됨 (버전: {version})")
        else:
            self.message_label.setText("브라우저 확인 : 지원되지 않음.")
            logging.info("브라우저 확인 : 지원되지 않음.")

    @QtCore.Slot()
    def load_file(self):
        file_name, _ = QtWidgets.QFileDialog.getOpenFileName(self, "파일 불러오기", "", "Excel Files (*.xls *.xlsx)")
        if file_name:
            self.file_name = file_name
            # 엑셀 파일을 열고 시트 이름 목록을 가져옴.
            self.exel_file = pd.ExcelFile(file_name)
            sheet_names = self.exel_file.sheet_names

            # 시트 이름을 선택하는 대화상자를 생성.
            sheet, ok = QtWidgets.QInputDialog.getItem(self, "시트 선택", "시트:", sheet_names, 0, False)
            if ok and sheet:
                self.selected_sheet_name = sheet  # 선택된 시트 이름 저장
                self.file_name_label.setText(f"선택된 파일: {self.file_name} ({self.selected_sheet_name})")
            else:
                self.selected_sheet_name = ""  # 선택이 취소된 경우
                self.exel_file.close()
                self.exel_file = None
                self.file_name_label.setText("선택된 파일 없음")

        else:
            self.file_name = ""
            self.selected_sheet_name = ""  # 파일 선택이 취소된 경우
            self.file_name_label.setText("선택된 파일 없음")

    @QtCore.Slot()
    def unload_file(self):
        self.file_name = ""
        self.selected_sheet_name = ""
        self.file_name_label.setText("선택된 파일 없음")
        if self.exel_file:
            self.exel_file.close()
            self.exel_file = None

    @QtCore.Slot()
    def start_macro(self):
        logging.info("매크로 시작")
        macro = Macro(self.webdriver, self.file_name, self.selected_sheet_name)
        macro.start()
        self.webdriver.close()

        # 에러 데이터를 엑셀 파일로 저장
        error_data = macro.get_error_dataframe()
        # 파일 이름 설정
        base_file_name = os.path.splitext(os.path.basename(self.file_name))[0] + '_' + self.selected_sheet_name
        # 홈 폴더의 바탕화면 경로 설정
        desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
        output_file_path = os.path.join(desktop_path, "오류사항_" + base_file_name + ".xlsx")

        # 에러 데이터프레임을 엑셀 파일로 저장
        logging.info(f"에러 데이터를 저장합니다: {output_file_path}")
        error_data.to_excel(output_file_path, index=False)

        # 사용자에게 저장 완료 알림
        QtWidgets.QMessageBox.information(self, "저장 완료", f"에러 데이터가 저장되었습니다:\n{output_file_path}")

    @QtCore.Slot()
    def show_example(self):
        # PyInstaller가 생성한 임시 디렉터리 경로를 얻기.
        # 애플리케이션이 PyInstaller로 패키징되지 않았다면, 현재 파일의 디렉터리를 사용.
        if getattr(sys, 'frozen', False):
            path = os.path.join(sys._MEIPASS, 'data/sample.xlsx')
        else:
            path = '../data/sample.xlsx'

        # 예시 파일의 경로를 확인
        if os.path.exists(path):
            # Windows 환경
            if os.name == 'nt':
                os.startfile(path)
            # macOS 환경
            elif os.name == 'posix':
                subprocess.run(['open', path])
            # Linux 환경 (대부분의 데스크탑 환경에서 동작)
            else:
                subprocess.run(['xdg-open', path])
        else:
            QtWidgets.QMessageBox.warning(self, "파일 없음", "예시 파일을 찾을 수 없습니다.")


class QTextEditLogger(logging.Handler):
    def __init__(self, parent):
        super().__init__()
        self.widget = QtWidgets.QPlainTextEdit(parent)
        self.widget.setReadOnly(True)

    def emit(self, record):
        msg = self.format(record)
        self.widget.appendPlainText(msg)
