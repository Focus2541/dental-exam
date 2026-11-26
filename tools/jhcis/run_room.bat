@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Room: a2 a3 p1 p2 p3 p4 p5 p6  (a = anuban, p = prathom)
echo Setup: calib = scrollbar arrows, provider = choose provider name
set /p ROOM=Room: 
set /p MODE=Enter = practice, save = save, redo = save again incl. done kids: 
set ARGS=
if /i "%MODE%"=="save" set ARGS=--save --auto
if /i "%MODE%"=="redo" set ARGS=--save --auto --redo
set PYTHONUTF8=1
python jhcis_dental_room.py %ROOM% %ARGS%
echo.
pause
