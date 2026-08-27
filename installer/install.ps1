<#
.SYNOPSIS
    Blue Page를 설치한다(관리자 권한 불필요, 현재 사용자 전용).
.DESCRIPTION
    게시된(publish) 빌드 결과물을 %LOCALAPPDATA%\Programs\BluePage 로 복사하고,
    BluePage.exe --register 를 실행해 .docx/.xlsx/.pptx 등을 "연결 프로그램" 후보로 등록한다.
.PARAMETER PublishDir
    dotnet publish 결과물이 있는 폴더. 기본값은 src/BluePage/publish.
.PARAMETER InstallDir
    설치 대상 폴더. 기본값은 %LOCALAPPDATA%\Programs\BluePage.
.PARAMETER SkipRegister
    파일 연결 등록을 건너뛰려면 지정한다.
.PARAMETER NoRestart
    설치 완료 후 자동으로 다시 실행하지 않으려면 지정한다(기본값은 자동 재시작).
#>
param(
    [string]$PublishDir = (Join-Path $PSScriptRoot "..\src\BluePage\publish"),
    [string]$InstallDir = (Join-Path $env:LOCALAPPDATA "Programs\BluePage"),
    [switch]$SkipRegister,
    [switch]$NoRestart
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path $PublishDir)) {
    throw "게시된 빌드를 찾을 수 없습니다: $PublishDir`n먼저 다음을 실행하세요:`n  dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -o `"$PublishDir`""
}

# 실행 중인 이전 버전이 있으면 파일이 잠겨 있어 복사가 실패하므로 먼저 종료한다(업데이트 설치 시나리오).
$wasRunning = $false
$runningProcess = Get-Process -Name "BluePage" -ErrorAction SilentlyContinue
if ($runningProcess) {
    $wasRunning = $true
    Write-Host "실행 중인 이전 버전을 종료합니다..."
    $runningProcess | Stop-Process -Force
    Start-Sleep -Seconds 1
}

New-Item -ItemType Directory -Force -Path $InstallDir | Out-Null
Copy-Item -Path (Join-Path $PublishDir '*') -Destination $InstallDir -Recurse -Force

Write-Host "설치 완료: $InstallDir"

$launcherExe = Join-Path $InstallDir "BluePage.exe"

if (-not $SkipRegister) {
    & $launcherExe --register
}

$configPath = Join-Path $env:LOCALAPPDATA "Microsoft365OfficeWebLauncher\config.json"
Write-Host ""
Write-Host "다음 단계:"
Write-Host "  1) $configPath 파일을 열어 auth.clientId를 Entra App Registration에서 발급받은 값으로 설정하세요."
Write-Host "     (README.txt의 'Entra App Registration 설정' 절차 참고)"
Write-Host "  2) 탐색기에서 .docx/.xlsx/.pptx 파일을 우클릭 → 연결 프로그램 → 다른 앱 선택 →"
Write-Host "     'Blue Page' 선택 후 '항상 이 앱 사용'을 체크하세요."
Write-Host "     (Windows 정책상 앱이 스스로 기본값을 강제 지정할 수 없어 이 1회 확인이 필요합니다.)"

# 업데이트로 인해(또는 개발 중 재빌드로) 종료했던 경우, 설치가 끝나면 자동으로 다시 띄워준다.
if (-not $NoRestart -and $wasRunning) {
    Write-Host ""
    Write-Host "이전에 실행 중이던 프로그램을 새 버전으로 다시 시작합니다..."
    Start-Process -FilePath $launcherExe
}
