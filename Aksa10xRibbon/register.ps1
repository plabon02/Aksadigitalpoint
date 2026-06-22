param(
    [switch]$Unregister
)

$ErrorActionPreference = "Stop"
$GUID = "{A1B2C3D4-E5F6-7890-ABCD-EF1234567893}"
$ProgId = "Aksa10xFaster.Connect"
$addinPath = Join-Path $PSScriptRoot "bin\WordAddIn.dll"

function Register-AddIn {
    Write-Host "=== Registering AKSA 10X FASTER ===" -ForegroundColor Cyan

    $tmpPath = Join-Path $PSScriptRoot "bin\WordAddIn.tmp.dll"
    if (-not (Test-Path $addinPath) -and (Test-Path $tmpPath)) {
        Write-Host "[INFO] Main DLL locked, using temp file" -ForegroundColor Yellow
        $addinPath = $tmpPath
    }
    if (-not (Test-Path $addinPath)) {
        Write-Error "WordAddIn.dll not found at: $addinPath"
        Write-Host "Build first: .\compile.ps1" -ForegroundColor Yellow
        exit 1
    }

    # Make path a proper file URI
    $codeBase = "file:///" + $addinPath.Replace("\", "/")

    # 1. Register CLSID under HKCU
    Write-Host "[1/3] Registering COM CLSID..." -ForegroundColor Yellow
    $clsidKey = "HKCU:\Software\Classes\CLSID\$GUID"
    New-Item -Path $clsidKey -Force | Out-Null
    Set-ItemProperty -Path $clsidKey -Name "(default)" -Value $ProgId

    New-Item -Path "$clsidKey\InprocServer32" -Force | Out-Null
    Set-ItemProperty -Path "$clsidKey\InprocServer32" -Name "(default)" -Value "mscoree.dll"
    Set-ItemProperty -Path "$clsidKey\InprocServer32" -Name "ThreadingModel" -Value "Both"
    Set-ItemProperty -Path "$clsidKey\InprocServer32" -Name "Assembly" -Value "WordAddIn, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null"
    Set-ItemProperty -Path "$clsidKey\InprocServer32" -Name "Class" -Value "WordToExeAddin.WordAddIn"
    Set-ItemProperty -Path "$clsidKey\InprocServer32" -Name "RuntimeVersion" -Value "v4.0.30319"
    Set-ItemProperty -Path "$clsidKey\InprocServer32" -Name "CodeBase" -Value $codeBase

    # 2. Register ProgId
    Write-Host "[2/3] Registering ProgId..." -ForegroundColor Yellow
    New-Item -Path "HKCU:\Software\Classes\$ProgId" -Force | Out-Null
    Set-ItemProperty -Path "HKCU:\Software\Classes\$ProgId" -Name "(default)" -Value "AKSA 10X FASTER"
    New-Item -Path "HKCU:\Software\Classes\$ProgId\CLSID" -Force | Out-Null
    Set-ItemProperty -Path "HKCU:\Software\Classes\$ProgId\CLSID" -Name "(default)" -Value $GUID

    # 3. Register Word Add-in
    Write-Host "[3/3] Registering Word Add-in..." -ForegroundColor Yellow
    $addinKey = "HKCU:\Software\Microsoft\Office\Word\Addins\$ProgId"
    New-Item -Path $addinKey -Force | Out-Null
    Set-ItemProperty -Path $addinKey -Name "Description" -Value "AKSA 10X FASTER - Word shortcut tools"
    Set-ItemProperty -Path $addinKey -Name "FriendlyName" -Value "AKSA 10X FASTER"
    Set-ItemProperty -Path $addinKey -Name "LoadBehavior" -Value 3

    Write-Host "`n=== Registration Complete ===" -ForegroundColor Green
    Write-Host "Please restart Microsoft Word." -ForegroundColor Yellow
    Write-Host "The 'AKSA 10X FASTER' tab will appear in the ribbon." -ForegroundColor Green
}

function Unregister-AddIn {
    Write-Host "=== Unregistering AKSA 10X FASTER ===" -ForegroundColor Cyan

    $paths = @(
        "HKCU:\Software\Classes\CLSID\$GUID",
        "HKCU:\Software\Classes\$ProgId",
        "HKCU:\Software\Microsoft\Office\Word\Addins\$ProgId"
    )

    foreach ($path in $paths) {
        if (Test-Path $path) {
            Remove-Item -Path $path -Recurse -Force
            Write-Host "[OK] Removed: $path" -ForegroundColor Green
        }
    }

    Write-Host "`n=== Unregistration Complete ===" -ForegroundColor Green
    Write-Host "Please restart Microsoft Word." -ForegroundColor Yellow
}

if ($Unregister) {
    Unregister-AddIn
} else {
    Register-AddIn
}
