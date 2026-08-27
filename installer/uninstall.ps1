<#
.SYNOPSIS
    Blue Page를 제거한다.
.DESCRIPTION
    BluePage.exe --unregister 로 파일 연결 등록을 해제한 뒤 설치 폴더를 삭제한다.
.PARAMETER InstallDir
    설치 폴더. 기본값은 %LOCALAPPDATA%\Programs\BluePage.
.PARAMETER KeepUserData
    로그/설정/토큰 캐시/업로드 매니페스트(%LOCALAPPDATA%\Microsoft365OfficeWebLauncher)를 보존하려면 지정한다.
#>
param(
    [string]$InstallDir = (Join-Path $env:LOCALAPPDATA "Programs\BluePage"),
    [switch]$KeepUserData
)

$ErrorActionPreference = "Stop"

$launcherExe = Join-Path $InstallDir "BluePage.exe"
if (Test-Path $launcherExe) {
    & $launcherExe --unregister
}

if (Test-Path $InstallDir) {
    Remove-Item -Recurse -Force $InstallDir -Confirm:$false
    Write-Host "설치 폴더를 삭제했습니다: $InstallDir"
}

if (-not $KeepUserData) {
    $userDataDir = Join-Path $env:LOCALAPPDATA "Microsoft365OfficeWebLauncher"
    if (Test-Path $userDataDir) {
        Remove-Item -Recurse -Force $userDataDir -Confirm:$false
        Write-Host "사용자 데이터를 삭제했습니다: $userDataDir"
    }
}

Write-Host "제거가 완료되었습니다."
