@echo off
cd /d %~dp0
echo.
echo  아파트 대시보드 서버 시작 중...
echo  브라우저가 자동으로 열립니다.
echo  종료하려면 이 창을 닫으세요.
echo.
python dashboard_server.py
pause
