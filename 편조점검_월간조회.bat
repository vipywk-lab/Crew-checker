@echo off
chcp 65001 > nul
cd /d "%~dp0"
echo ==========================================
echo  편조점검 월간 자동 조회
echo  crew_monthly_checker.py 를 실행합니다.
echo ==========================================
echo.
python crew_monthly_checker.py
pause
