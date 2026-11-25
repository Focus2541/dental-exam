@echo off
chcp 65001 >nul
cd /d "%~dp0"
where python >nul 2>nul || (echo ไม่พบ Python ในเครื่องนี้ & pause & exit /b)
python -c "import pymysql" 2>nul || python -m pip install --quiet pymysql
python jhcis_capture.py
