#!/bin/bash

# PyInstaller로 main.spec 파일을 사용하여 빌드
pyinstaller hometax_macro_simple.spec

# 빌드 완료 메시지
echo "빌드가 완료되었습니다."
