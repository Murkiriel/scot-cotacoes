@echo off
rem Coleta as cotações da Scot e atualiza scot_cotacoes.xlsx.
rem Argumentos extras são repassados ao script (ex.: rodar.bat --paginas boi-gordo).
cd /d "%~dp0"
"C:\Python\python.exe" "%~dp0scot_cotacoes.py" %*
if errorlevel 1 pause
