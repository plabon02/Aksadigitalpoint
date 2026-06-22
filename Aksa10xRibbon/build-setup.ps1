$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path

Write-Host "=== Building AKSA 10X FASTER Setup ===" -ForegroundColor Cyan

Write-Host "[1/4] Building WordAddIn.dll..." -ForegroundColor Yellow
$buildLog = & "powershell" -ExecutionPolicy Bypass -File (Join-Path $root "build.ps1") 2>&1
$buildLog | ForEach-Object { Write-Host $_ }
if ($LASTEXITCODE -ne 0) {
    Write-Error "DLL build failed (exit code: $LASTEXITCODE)"
    exit 1
}

# Copy built DLL to bin/ for packaging
$builtDll = Join-Path $root "src\WordToExeAddin\bin\Release\WordAddIn.dll"
$pkgDll = Join-Path $root "bin\WordAddIn.dll"
if (Test-Path $builtDll) {
    Copy-Item -Path $builtDll -Destination $pkgDll -Force
    Write-Host "  DLL copied from Release build" -ForegroundColor Gray
}

Write-Host "[2/4] Creating package.zip..." -ForegroundColor Yellow
$zipPath = Join-Path $root "package.zip"
if (Test-Path $zipPath) { Remove-Item $zipPath -Force }

# Files to include
$items = @(
    (Join-Path $root "bin"),
    (Join-Path $root "icons"),
    (Join-Path $root "Document Template")
)

$tmpDir = Join-Path $env:TEMP "aksapkg_$(Get-Random)"
New-Item -ItemType Directory -Path $tmpDir -Force | Out-Null
try {
    foreach ($item in $items) {
        $name = Split-Path -Leaf $item
        $dest = Join-Path $tmpDir $name
        if (Test-Path $item) {
            Copy-Item -Path $item -Destination $dest -Recurse -Force
            Write-Host "  Added: $name"
        }
    }
    Compress-Archive -Path "$tmpDir\*" -DestinationPath $zipPath -Force
    Write-Host "[OK] package.zip created ($((Get-Item $zipPath).Length / 1KB) KB)" -ForegroundColor Green
}
finally {
    Remove-Item $tmpDir -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Host "[3/4] Building setup executable..." -ForegroundColor Yellow
$csc = "C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\Roslyn\csc.exe"
if (-not (Test-Path $csc)) {
    $csc = "C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe"
}

$refs = @(
    "/reference:System.dll",
    "/reference:System.Core.dll",
    "/reference:System.Windows.Forms.dll",
    "/reference:System.Drawing.dll",
    "/reference:C:\Program Files (x86)\Reference Assemblies\Microsoft\Framework\.NETFramework\v4.8\System.IO.Compression.FileSystem.dll",
    "/reference:C:\Program Files (x86)\Reference Assemblies\Microsoft\Framework\.NETFramework\v4.8\System.IO.Compression.dll"
)

$outExe = Join-Path $root "Setup.exe"
$ico = Join-Path $root "app.ico"
$setupCs = Join-Path $root "Setup.cs"

& $csc /target:winexe $refs "/win32icon:$ico" "/resource:$zipPath,Data.zip" "/out:$outExe" $setupCs 2>&1
if ($LASTEXITCODE -eq 0) {
    Write-Host "[OK] Setup.exe created ($((Get-Item $outExe).Length / 1KB) KB)" -ForegroundColor Green
    Write-Host ""
    Write-Host "=== Build Complete ===" -ForegroundColor Cyan
    Write-Host "Output: $outExe" -ForegroundColor Green
} else {
    Write-Error "Build failed (exit code: $LASTEXITCODE)"
    exit 1
}

Write-Host "[4/4] Install the add-in..." -ForegroundColor Yellow
Write-Host "  Run: Setup.exe (double-click)" -ForegroundColor Gray
