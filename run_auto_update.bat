@echo off
REM 부태리 신고가 자동 업데이트 배치 파일
REM Windows 작업 스케줄러용

cd /d "d:\python work\1. max"

REM Python 실행 (가상환경을 사용하는 경우 경로 수정 필요)
python auto_update_html.py

REM 오류 코드 확인
if %ERRORLEVEL% NEQ 0 (
    echo 오류 발생: %ERRORLEVEL%
    exit /b %ERRORLEVEL%
)

exit /b 0
