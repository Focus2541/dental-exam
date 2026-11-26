@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Room: a2 a3 p1 p2 p3 p4 p5 p6  (a = anuban, p = prathom)
set /p ROOM=Room: 
set /p MODE=Type save to really save, or just Enter to practice: 
set ARGS=
if /i "%MODE%"=="save" set ARGS=--save --auto
set PYTHONUTF8=1
python jhcis_dental_room.py %ROOM% %ARGS%
echo.
echo exit code %ERRORLEVEL%
pause
