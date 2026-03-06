@echo off
cd /d %~dp0
if exist .venv\Scripts\activate.bat (
    call .venv\Scripts\activate.bat
)
start "" python run_listener.py
start "" python manage.py runserver

