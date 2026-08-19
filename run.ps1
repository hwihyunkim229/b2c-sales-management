$ErrorActionPreference = "Stop"

$projectDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
$pythonExecutable = Join-Path $projectDirectory ".venv\Scripts\python.exe"

if (-not (Test-Path -LiteralPath $pythonExecutable)) {
    throw "가상환경을 찾을 수 없습니다: $pythonExecutable"
}

Set-Location -LiteralPath $projectDirectory

Write-Host "B2C Sales Management 서버를 시작합니다."
Write-Host "이 PC 접속 주소: http://127.0.0.1:8002"
Write-Host "다른 PC 접속 주소: http://<서버 PC의 IPv4 주소>:8002"
Write-Host "서버 종료: Ctrl+C"

& $pythonExecutable `
    -m uvicorn `
    app.main:app `
    --app-dir $projectDirectory `
    --host 0.0.0.0 `
    --port 8002
