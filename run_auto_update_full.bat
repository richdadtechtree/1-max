@echo off
REM 부태리 신고가 완전 자동 업데이트 배치 파일
REM 데이터 갱신 + HTML 생성 통합
REM Windows 작업 스케줄러용

cd /d "d:\python work\1. max"

REM Python 실행
python auto_update_full.py

REM 오류 코드 확인
if %ERRORLEVEL% NEQ 0 (
    echo 오류 발생: %ERRORLEVEL%
    exit /b %ERRORLEVEL%
)

exit /b 0
