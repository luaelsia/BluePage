<#
.SYNOPSIS
    Blue Page를 설치한다(관리자 권한 불필요, 현재 사용자 전용).
.DESCRIPTION
    게시된(publish) 빌드 결과물을 %LOCALAPPDATA%\Programs\BluePage 로 복사하고,
    BluePage.exe --register 를 실행해 .docx/.xlsx/.pptx 등을 "연결 프로그램" 후보로 등록한다.
.PARAMETER PublishDir
    배포 파일이 있는 폴더. Release ZIP에서는 스크립트가 있는 폴더를 자동으로 사용하며,
    소스 저장소에서는 src/BluePage/publish를 사용한다.
.PARAMETER InstallDir
    설치 대상 폴더. 기본값은 %LOCALAPPDATA%\Programs\BluePage.
.PARAMETER SkipRegister
    파일 연결 등록을 건너뛰려면 지정한다.
.PARAMETER NoRestart
    설치 완료 후 자동으로 다시 실행하지 않으려면 지정한다(기본값은 자동 재시작).
#>
param(
    [string]$PublishDir = "",
    [string]$InstallDir = (Join-Path $env:LOCALAPPDATA "Programs\BluePage"),
    [switch]$SkipRegister,
    [switch]$NoRestart
)

$ErrorActionPreference = "Stop"

if ([string]::IsNullOrWhiteSpace($PublishDir)) {
    $releaseExe = Join-Path $PSScriptRoot "BluePage.exe"
    $PublishDir = if (Test-Path -LiteralPath $releaseExe) {
        $PSScriptRoot
    }
    else {
        Join-Path $PSScriptRoot "..\src\BluePage\publish"
    }
}

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
$packageOnlyFiles = @('install.ps1', 'uninstall.ps1', 'README.txt', 'LICENSE.txt')
Get-ChildItem -LiteralPath $PublishDir -File |
    Where-Object { $_.Name -notin $packageOnlyFiles } |
    Copy-Item -Destination $InstallDir -Force

Write-Host "설치 완료: $InstallDir"

$launcherExe = Join-Path $InstallDir "BluePage.exe"

if (-not $SkipRegister) {
    & $launcherExe --register
}

Write-Host ""
Write-Host "다음 단계:"
Write-Host "  1) 탐색기에서 .docx/.xlsx/.pptx 파일을 우클릭 → 연결 프로그램 → 다른 앱 선택 →"
Write-Host "     'Blue Page' 선택 후 '항상 이 앱 사용'을 체크하세요."
Write-Host "     (Windows 정책상 앱이 스스로 기본값을 강제 지정할 수 없어 이 1회 확인이 필요합니다.)"
Write-Host "  2) 시작 메뉴에서 Blue Page를 실행하고 Microsoft 계정으로 로그인하세요."

# 새 설치와 업데이트 모두 설치가 끝나면 Blue Page를 실행한다.
if (-not $NoRestart) {
    Write-Host ""
    Write-Host "Blue Page를 시작합니다..."
    Start-Process -FilePath $launcherExe
}
