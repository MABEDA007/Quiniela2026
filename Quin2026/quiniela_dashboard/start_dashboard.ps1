$python = "C:\Users\eduar\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
$app = "C:\Users\eduar\OneDrive - PETROLEOS MEXICANOS\Quin2026\quiniela_dashboard\app.py"
Set-Location "C:\Users\eduar\OneDrive - PETROLEOS MEXICANOS\Quin2026"
Write-Host "Dashboard Quiniela Mundial 2026"
Write-Host "Abre http://127.0.0.1:8765 en tu navegador."
Write-Host "Deja esta ventana abierta mientras uses el dashboard."
& $python $app
