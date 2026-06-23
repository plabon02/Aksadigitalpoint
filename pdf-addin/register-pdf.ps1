param(
    [switch]$Unregister
)

$ErrorActionPreference = "Stop"
$GUID = "{33F39564-802B-412B-A713-E947F98DCBB9}"
$ProgId = "AksaPdfTools.Connect"
$addinPath = Join-Path $PSScriptRoot "src\AksaPdfAddin\bin\Release\AksaPdfAddin.dll"

function Register-AddIn {
    Write-Host "=== Registering AKSA PDF Tools ===" -ForegroundColor Cyan

    if (-not (Test-Path $addinPath)) {
        Write-Error "AksaPdfAddin.dll not found at: $addinPath"
        Write-Host "Build first: open AksaPdfAddin.sln in VS2022 and build" -ForegroundColor Yellow
        exit 1
    }

    $codeBase = "file:///" + $addinPath.Replace("\", "/")

    Write-Host "[1/3] Registering COM CLSID..." -ForegroundColor Yellow
    $clsidKey = "HKCU:\Software\Classes\CLSID\$GUID"
    New-Item -Path $clsidKey -Force | Out-Null
    Set-ItemProperty -Path $clsidKey -Name "(default)" -Value $ProgId

    New-Item -Path "$clsidKey\InprocServer32" -Force | Out-Null
    Set-ItemProperty -Path "$clsidKey\InprocServer32" -Name "(default)" -Value "mscoree.dll"
    Set-ItemProperty -Path "$clsidKey\InprocServer32" -Name "ThreadingModel" -Value "Both"
    Set-ItemProperty -Path "$clsidKey\InprocServer32" -Name "Assembly" -Value "AksaPdfAddin, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null"
    Set-ItemProperty -Path "$clsidKey\InprocServer32" -Name "Class" -Value "AksaPdfAddin.PdfAddIn"
    Set-ItemProperty -Path "$clsidKey\InprocServer32" -Name "RuntimeVersion" -Value "v4.0.30319"
    Set-ItemProperty -Path "$clsidKey\InprocServer32" -Name "CodeBase" -Value $codeBase

    Write-Host "[2/3] Registering ProgId..." -ForegroundColor Yellow
    New-Item -Path "HKCU:\Software\Classes\$ProgId" -Force | Out-Null
    Set-ItemProperty -Path "HKCU:\Software\Classes\$ProgId" -Name "(default)" -Value "AKSA PDF Tools"
    New-Item -Path "HKCU:\Software\Classes\$ProgId\CLSID" -Force | Out-Null
    Set-ItemProperty -Path "HKCU:\Software\Classes\$ProgId\CLSID" -Name "(default)" -Value $GUID

    Write-Host "[3/3] Registering Word Add-in..." -ForegroundColor Yellow
    $addinKey = "HKCU:\Software\Microsoft\Office\Word\Addins\$ProgId"
    New-Item -Path $addinKey -Force | Out-Null
    Set-ItemProperty -Path $addinKey -Name "Description" -Value "AKSA PDF Tools - PDF conversion, merge, split, and info"
    Set-ItemProperty -Path $addinKey -Name "FriendlyName" -Value "AKSA PDF Tools"
    Set-ItemProperty -Path $addinKey -Name "LoadBehavior" -Value 3

    Write-Host "`n=== Registration Complete ===" -ForegroundColor Green
    Write-Host "Please restart Microsoft Word." -ForegroundColor Yellow
    Write-Host "The 'AKSA PDF Tools' tab will appear in the ribbon." -ForegroundColor Green
}

function Unregister-AddIn {
    Write-Host "=== Unregistering AKSA PDF Tools ===" -ForegroundColor Cyan

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
