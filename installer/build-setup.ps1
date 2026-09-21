<#
.SYNOPSIS
    Blue Page 설치 파일(BluePage-Setup-vX.Y.Z.exe)을 만든다.
.DESCRIPTION
    (1) src/BluePage 를 self-contained 단일 파일로 게시(publish)하고,
    (2) Inno Setup 컴파일러(ISCC.exe)로 installer/BluePage.iss 를 컴파일해
    release/BluePage-Setup-vX.Y.Z.exe 를 만든다. 버전은 BluePage.csproj 의 <Version> 을 그대로 쓴다.

    Inno Setup 6 이 필요하다. 없으면 다음으로 설치한다.
        winget install --id JRSoftware.InnoSetup -e
.PARAMETER SkipPublish
    이미 게시된 src/BluePage/publish 를 그대로 쓰고 dotnet publish 를 건너뛴다.
.PARAMETER OutputDir
    설치 파일을 놓을 폴더. 기본값은 저장소의 release 폴더.
#>
param(
    [switch]$SkipPublish,
    [string]$OutputDir = (Join-Path $PSScriptRoot "..\release")
)

$ErrorActionPreference = "Stop"

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$projectDir = Join-Path $repoRoot "src\BluePage"
$publishDir = Join-Path $projectDir "publish"
$issPath = Join-Path $PSScriptRoot "BluePage.iss"

# 버전은 csproj 한 곳에서만 관리한다.
[xml]$csproj = Get-Content -LiteralPath (Join-Path $projectDir "BluePage.csproj")
$version = ($csproj.Project.PropertyGroup | Where-Object { $_.Version } | Select-Object -First 1).Version
if ([string]::IsNullOrWhiteSpace($version)) {
    throw "BluePage.csproj 에서 <Version> 을 찾을 수 없습니다."
}

$isccCandidates = @(
    (Join-Path $env:LOCALAPPDATA "Programs\Inno Setup 6\ISCC.exe"),
    (Join-Path ${env:ProgramFiles(x86)} "Inno Setup 6\ISCC.exe"),
    (Join-Path $env:ProgramFiles "Inno Setup 6\ISCC.exe")
)
$iscc = $isccCandidates | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
if (-not $iscc) {
    throw "Inno Setup 6 (ISCC.exe) 을 찾을 수 없습니다. 설치: winget install --id JRSoftware.InnoSetup -e"
}

if (-not $SkipPublish) {
    # 실행 중인 Blue Page 가 publish 폴더의 exe 를 잠그고 있을 수 있다.
    Get-Process -Name "BluePage" -ErrorAction SilentlyContinue | Stop-Process -Force
    if (Test-Path -LiteralPath $publishDir) {
        Remove-Item -LiteralPath $publishDir -Recurse -Force
    }

    Write-Host "게시 중 (v$version)..."
    Push-Location $projectDir
    try {
        dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:PublishTrimmed=false -o $publishDir
        if ($LASTEXITCODE -ne 0) { throw "dotnet publish 실패 (exit $LASTEXITCODE)" }
    }
    finally {
        Pop-Location
    }
}

if (-not (Test-Path -LiteralPath (Join-Path $publishDir "BluePage.exe"))) {
    throw "게시된 BluePage.exe 가 없습니다: $publishDir"
}

New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null
$OutputDir = (Resolve-Path -LiteralPath $OutputDir).Path

Write-Host "설치 파일 컴파일 중..."
& $iscc /Q "/DAppVersion=$version" "/DPublishDir=$publishDir" "/DOutputDir=$OutputDir" $issPath
if ($LASTEXITCODE -ne 0) { throw "ISCC 실패 (exit $LASTEXITCODE)" }

$setupExe = Join-Path $OutputDir "BluePage-Setup-v$version.exe"
Write-Host "완료: $setupExe"
