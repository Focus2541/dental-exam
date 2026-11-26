@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo === 1/3 install libraries ===
python -m pip install --upgrade pyinstaller pymysql openpyxl python-docx pyautogui pygetwindow pyperclip pillow pyscreeze
if errorlevel 1 goto fail
echo === 2/3 build JHCIS_School.exe ===
python -m PyInstaller --noconfirm --clean --onedir --console --name JHCIS_School --collect-data docx --hidden-import jhcis_roster_compare --hidden-import jhcis_student_fill --hidden-import jhcis_newperson_fill --hidden-import jhcis_dental_fill --hidden-import jhcis_dental_room --hidden-import jhcis_link_existing --hidden-import docx_roster_to_xlsx --hidden-import jhcis_birth_fill --hidden-import jhcis_check_report --hidden-import jhcis_check_detail --hidden-import jhcis_codes --hidden-import jhcis_compare_person --hidden-import pymysql --hidden-import openpyxl --hidden-import docx --hidden-import pyautogui --hidden-import pygetwindow --hidden-import pyperclip --hidden-import PIL --hidden-import pymsgbox --hidden-import pyscreeze --hidden-import pytweening --hidden-import mouseinfo --hidden-import tkinter jhcis_school_app.py
if errorlevel 1 goto fail
echo === 3/3 done ===
echo Program folder: %~dp0dist\JHCIS_School
echo Copy the WHOLE folder JHCIS_School to the flash drive, then double-click JHCIS_School.exe
explorer "%~dp0dist"
pause
exit /b 0
:fail
echo BUILD FAILED - take a photo of this window
pause
exit /b 1
