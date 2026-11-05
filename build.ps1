# Script de compilación de RENO v2.0.0 a ejecutable autónomo (.exe)
Write-Host "====================================================" -ForegroundColor Cyan
Write-Host "  Iniciando empaquetado de RENO v2.0.0..." -ForegroundColor Cyan
Write-Host "====================================================" -ForegroundColor Cyan

$pythonCmd = "python"
if (Test-Path ".\venv\Scripts\python.exe") {
    $pythonCmd = ".\venv\Scripts\python.exe"
} elseif (Test-Path ".\.venv\Scripts\python.exe") {
    $pythonCmd = ".\.venv\Scripts\python.exe"
}

$args = @("-m", "PyInstaller", "--clean", "--noconfirm", "main.spec")

Write-Host "Ejecutando: $pythonCmd $($args -join ' ')" -ForegroundColor Yellow
& $pythonCmd $args

if ($LASTEXITCODE -eq 0) {
    Write-Host "`n====================================================" -ForegroundColor Green
    Write-Host "  ¡Compilación completada con éxito!" -ForegroundColor Green
    Write-Host "  Ejecutable generado en: dist\RENO.exe" -ForegroundColor Green
    Write-Host "====================================================" -ForegroundColor Green
    Write-Host "`nSiguiente paso opcional (Instalador Windows):" -ForegroundColor Cyan
    Write-Host "  Abra RENO.iss con Inno Setup Compiler y presione Build (Ctrl+F9)." -ForegroundColor Gray
} else {
    Write-Host "`n[ERROR] Falló el proceso de compilación de PyInstaller." -ForegroundColor Red
    exit $LASTEXITCODE
}
