# Windows 프로그램 빌드/사용

## 개발 중 실행

Windows에서 아래 파일을 더블클릭합니다.

```bat
run_windows_dev.bat
```

브라우저가 자동으로 열리고, 월별 엑셀 업로드와 검색/그래프 기능을 사용할 수 있습니다.

## exe 만들기

Windows PC에서 아래 파일을 더블클릭합니다.

```bat
build_windows.bat
```

완료되면 아래 실행파일이 생성됩니다.

```text
dist\약품사용량추세.exe
```

이 exe를 배포하면 됩니다.

## 데이터 저장 위치

exe로 실행하면 기준 피벗과 출력 파일은 exe 옆 폴더에 저장됩니다.

```text
data\master_pivot.xlsx
outputs\
```

기준 피벗을 초기화하려면 `data\master_pivot.xlsx`를 교체하거나 삭제한 뒤 프로그램을 다시 실행하면 됩니다.
