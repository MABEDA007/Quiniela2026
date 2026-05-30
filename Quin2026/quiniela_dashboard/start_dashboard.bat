@echo off
set "PYTHON=C:\Users\eduar\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
set "APP=C:\Users\eduar\OneDrive - PETROLEOS MEXICANOS\Quin2026\quiniela_dashboard\app.py"
cd /d "C:\Users\eduar\OneDrive - PETROLEOS MEXICANOS\Quin2026"
echo Dashboard Quiniela Mundial 2026
echo.
echo Abre http://127.0.0.1:8765 en tu navegador.
echo Deja esta ventana abierta mientras uses el dashboard.
echo.
"%PYTHON%" "%APP%"
pause
