param(
    [switch]$Register
)

$ErrorActionPreference = "Stop"
$csc = "C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe"
$solutionDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$libDir = Join-Path $solutionDir "lib"
$outDir = Join-Path $solutionDir "bin"

if (-not (Test-Path $libDir)) { New-Item -ItemType Directory -Path $libDir -Force | Out-Null }
if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }

Write-Host "=== AKSA 10X FASTER - Build ===" -ForegroundColor Cyan

Write-Host "Step 1: Copying reference assemblies..." -ForegroundColor Yellow
$refs = @(
    @("C:\Windows\assembly\GAC_MSIL\Microsoft.Office.Interop.Word\15.0.0.0__71e9bce111e9429c\Microsoft.Office.Interop.Word.dll", "Microsoft.Office.Interop.Word.dll"),
    @("C:\Windows\assembly\GAC_MSIL\office\15.0.0.0__71e9bce111e9429c\office.dll", "office.dll"),
    @("C:\Windows\assembly\GAC\stdole\7.0.3300.0__b03f5f7f11d50a3a\stdole.dll", "stdole.dll"),
    @("C:\Windows\assembly\GAC\Extensibility\7.0.3300.0__b03f5f7f11d50a3a\Extensibility.dll", "Extensibility.dll")
)

foreach ($ref in $refs) {
    $src, $name = $ref[0], $ref[1]
    $dst = Join-Path $libDir $name
    if (Test-Path $src) { Copy-Item -Path $src -Destination $dst -Force; Write-Host "[OK] $name" -ForegroundColor Green }
    else { Write-Warning "Not found: $src" }
}

Write-Host "`nStep 2: Building WordAddIn.dll..." -ForegroundColor Yellow
$addinOut = Join-Path $outDir "WordAddIn.dll"

# Force-delete old file (handles locks from cmd)
cmd /c "del /f /q `"$addinOut`" 2>nul"

& $csc /target:library /out:$addinOut `
    /reference:"System.dll" /reference:"System.Core.dll" `
    /reference:"System.Windows.Forms.dll" /reference:"System.Drawing.dll" `
    /reference:"System.Web.dll" /reference:"System.Web.Extensions.dll" `
    /reference:"Microsoft.VisualBasic.dll" `
    /reference:"Microsoft.CSharp.dll" `
    /reference:"$libDir\Microsoft.Office.Interop.Word.dll" `
    /reference:"$libDir\office.dll" /reference:"$libDir\stdole.dll" `
    /reference:"$libDir\Extensibility.dll" `
    "/resource:$solutionDir\src\WordToExeAddin\WordToExeAddin_Ribbon.xml,WordToExeAddin.WordToExeAddin_Ribbon.xml" `
    "$solutionDir\src\WordToExeAddin\WordAddIn.cs" `
    "$solutionDir\src\WordToExeAddin\ApiSettingsForm.cs" `
    "$solutionDir\src\WordToExeAddin\AiMiniToolbar.cs" `
    "$solutionDir\src\WordToExeAddin\LicenseService.cs" `
    "$solutionDir\src\WordToExeAddin\RegistrationForm.cs" `
    "$solutionDir\src\WordToExeAddin\Properties\AssemblyInfo.cs" 2>&1

if ($LASTEXITCODE -ne 0) {
    Write-Error "Build failed"
    exit 1
}

Write-Host "[OK] WordAddIn.dll" -ForegroundColor Green
Write-Host "`n=== Build Complete ===" -ForegroundColor Cyan

if ($Register) {
    Write-Host "`nRegistering..." -ForegroundColor Yellow
    & ".\register.ps1"
}
