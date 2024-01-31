# 읽어드린 파일 Record 클래스로 변환 검증
from hometax_macro_simple.macro import Macro
from hometax_macro_simple.webdriver import WebDriverManager

path = input("파일 경로를 입력하세요: ")
sheet_name = input("시트 이름을 입력하세요: ")

macro = Macro(WebDriverManager(), path, sheet_name)
check = True
while macro.record.next():
    data = macro.record.get_current_data()
    print(f"#### data : {data}")
