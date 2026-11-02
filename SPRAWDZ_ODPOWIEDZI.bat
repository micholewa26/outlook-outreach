@echo off
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0sprawdz_odpowiedzi.ps1"
echo.
echo [Okno nie zamknie sie samo - nacisnij dowolny klawisz]
pause >nul
