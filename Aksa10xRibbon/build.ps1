param(
    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Release"
)

$ErrorActionPreference = "Stop"
$msbuild = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\MSBuild.exe"

if (-not (Test-Path $msbuild)) {
    $msbuild = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\MSBuild.exe"
}
if (-not (Test-Path $msbuild)) {
    $msbuild = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe"
}
if (-not (Test-Path $msbuild)) {
    $vsWhere = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"
    if (Test-Path $vsWhere) {
        $vsPath = & $vsWhere -latest -property installationPath
        $msbuild = Join-Path $vsPath "MSBuild\Current\Bin\MSBuild.exe"
    }
}
if (-not (Test-Path $msbuild) -and (Get-Command "msbuild" -ErrorAction SilentlyContinue)) {
    $msbuild = (Get-Command "msbuild").Source
}
if (-not (Test-Path $msbuild)) {
    Write-Error "MSBuild not found."
    exit 1
}

$solutionPath = Join-Path $PSScriptRoot "src\WordToExeAddin.sln"

Write-Host "=== Building AKSA 10X FASTER ===" -ForegroundColor Cyan
Write-Host "Configuration: $Configuration" -ForegroundColor Gray

& $msbuild "$PSScriptRoot\src\WordToExeAddin\WordToExeAddin.csproj" /p:Configuration=$Configuration /t:Build /v:m
if ($LASTEXITCODE -ne 0) {
    Write-Error "Build failed"
    exit 1
}

Write-Host "[OK] Build complete" -ForegroundColor Green
Write-Host "Output: src\WordToExeAddin\bin\$Configuration\WordAddIn.dll" -ForegroundColor Green
Write-Host "`nTo register, run: .\register.ps1" -ForegroundColor Yellow
