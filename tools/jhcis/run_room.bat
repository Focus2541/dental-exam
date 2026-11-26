@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Room: a2 a3 p1 p2 p3 p4 p5 p6  or  all = every class automatically  or  all a3 = every class starting at a3
echo Setup (once): calib = scrollbar arrows, provider = provider name, school = class box + dental button, tprov = provider column in table
set /p ROOM=Room: 
set /p MODE=Enter = practice, save = save, redo = save again incl. done kids, test = practice incl. done kids, night = save all again with no questions: 
set ARGS=
if /i "%MODE%"=="save" set ARGS=--save --auto
if /i "%MODE%"=="redo" set ARGS=--save --auto --redo
if /i "%MODE%"=="test" set ARGS=--redo
if /i "%MODE%"=="night" set ARGS=--save --auto --redo --night
set PYTHONUTF8=1
python jhcis_dental_room.py %ROOM% %ARGS%
echo.
pause
