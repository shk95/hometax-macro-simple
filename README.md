국세청 홈텍스 근로소득 지급명세서 제출 매크로
================================================

### 적용 항목

- 성명
- 주민등록번호
- 시작일자
- 종료일자
- 급여
- 소득세
- 지방소득세
- 국민연금보험료
- 건강보험료
- 고용보험료

### 사용 방법

**지원 환경**: 엣지 브라우저, 윈도우 권장

1. 프로그램을 실행하여 홈텍스 사이트를 오픈.
2. 홈텍스 사이트 내에서 로그인 후 '근로소득 지급명세서 제출' 페이지로 이동.
3. '파일 불러오기' 버튼을 클릭하여 적용할 엑셀 파일을 선택.
4. '매크로 시작' 버튼을 클릭하여 엑셀 파일에 기재된 정보를 바탕으로 자동 제출을 시작.
5. 엑셀 파일의 데이터 형식은 '예시 파일 저장하기' 버튼을 통해 저장 가능한 예시 파일에서 확인할 수 있음.
6. 매크로 작업이 종료되면, 오류 사항을 담은 엑셀 파일이 바탕화면에 저장됨.

### 빌드 방법

#### macOS & Linux:

```bash
$ poetry install
$ poetry shell
$ ./build.sh
```

#### Windows:

```cmd
> poetry install
> poetry shell
> .\build.ps1
```

---
<div align="center">
    <img src="https://github.com/shk95/hometax-macro-simple/assets/101378576/4ffb6966-a7b1-44c9-8629-acabed1a54d6" alt="image1" style="width: 60%; margin-bottom: 20px;"/>
    <img src="https://github.com/shk95/hometax-macro-simple/assets/101378576/269b65a1-ca78-44ec-a0e6-df101878c417" alt="image2" style="width: 60%;"/>
</div>
