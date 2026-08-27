<#
.SYNOPSIS
    개발 중 소스를 고친 뒤 다시 게시(publish)하고, 실행 중이던 Blue Page를 자동으로 재시작한다.
.DESCRIPTION
    dotnet publish는 실행 중인 BluePage.exe가 파일을 잠그고 있으면 실패한다. 이 스크립트는
    (1) 실행 중이면 종료 → (2) self-contained 단일 파일로 재게시 → (3) 원래 실행 중이었다면 자동으로 다시 실행
    까지 한 번에 처리한다. 코드를 수정할 때마다 이 스크립트만 실행하면 된다.
.PARAMETER ProjectDir
    BluePage.csproj이 있는 폴더. 기본값은 src/BluePage.
#>
param(
    [string]$ProjectDir = (Join-Path $PSScriptRoot "..\src\BluePage")
)

$ErrorActionPreference = "Stop"

$wasRunning = $false
$runningProcess = Get-Process -Name "BluePage" -ErrorAction SilentlyContinue
if ($runningProcess) {
    $wasRunning = $true
    Write-Host "실행 중인 Blue Page를 종료합니다..."
    $runningProcess | Stop-Process -Force
    Start-Sleep -Seconds 1
}

Write-Host "게시 중..."
Push-Location $ProjectDir
try {
    dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:PublishTrimmed=false -o publish
}
finally {
    Pop-Location
}

$exePath = Join-Path $ProjectDir "publish\BluePage.exe"

if ($wasRunning) {
    Write-Host "새 버전으로 다시 실행합니다: $exePath"
    Start-Process -FilePath $exePath
}
else {
    Write-Host "게시 완료: $exePath (원래 실행 중이 아니었으므로 자동으로 켜지 않았습니다)"
}
