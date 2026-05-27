@echo off
cd /d C:\Users\SAMSUNG\eland-hr-faq

:: Flask 시작
start "" /B pythonw C:\Users\SAMSUNG\eland-hr-faq\app.py

:: ngrok 시작 (5초 대기 후)
timeout /t 5 /nobreak > nul
start "" /B "C:\Users\SAMSUNG\AppData\Local\Microsoft\WinGet\Packages\Ngrok.Ngrok_Microsoft.Winget.Source_8wekyb3d8bbwe\ngrok.exe" http --domain=tubby-settle-sadden.ngrok-free.dev 5000
