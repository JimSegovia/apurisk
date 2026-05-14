param(
    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Debug"
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$assemblyPath = Join-Path $root "build\$Configuration\Apurisk.ExcelAddIn.dll"

if (-not (Test-Path $assemblyPath)) {
    throw "No existe $assemblyPath. Ejecuta scripts\build.ps1 primero."
}

$assembly = [System.Reflection.AssemblyName]::GetAssemblyName($assemblyPath)
$clsid = "{7BD16DC9-26B6-4C37-8E23-A4E80504D9E4}"
$progId = "Apurisk.ExcelAddIn"
$className = "Apurisk.ExcelAddIn.Connect"
$codeBase = (New-Object System.Uri($assemblyPath)).AbsoluteUri

$classesRoot = "HKCU:\Software\Classes"
$progIdRoot = Join-Path $classesRoot $progId
$clsidRoot = Join-Path $classesRoot "CLSID\$clsid"
$inprocRoot = Join-Path $clsidRoot "InprocServer32"
$versionRoot = Join-Path $inprocRoot $assembly.Version.ToString()
$implementedCategoriesRoot = Join-Path $clsidRoot "Implemented Categories\{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}"
$officeRoot = "HKCU:\Software\Microsoft\Office\Excel\Addins\$progId"

New-Item -Path $progIdRoot -Force | Out-Null
New-ItemProperty -Path $progIdRoot -Name "(default)" -Value "Apurisk Excel Add-in" -Force | Out-Null
New-Item -Path (Join-Path $progIdRoot "CLSID") -Force | Out-Null
New-ItemProperty -Path (Join-Path $progIdRoot "CLSID") -Name "(default)" -Value $clsid -Force | Out-Null

New-Item -Path $inprocRoot -Force | Out-Null
New-ItemProperty -Path $inprocRoot -Name "(default)" -Value "mscoree.dll" -Force | Out-Null
New-ItemProperty -Path $inprocRoot -Name "ThreadingModel" -Value "Both" -Force | Out-Null
New-ItemProperty -Path $inprocRoot -Name "Class" -Value $className -Force | Out-Null
New-ItemProperty -Path $inprocRoot -Name "Assembly" -Value $assembly.FullName -Force | Out-Null
New-ItemProperty -Path $inprocRoot -Name "RuntimeVersion" -Value "v4.0.30319" -Force | Out-Null
New-ItemProperty -Path $inprocRoot -Name "CodeBase" -Value $codeBase -Force | Out-Null

New-Item -Path $versionRoot -Force | Out-Null
New-ItemProperty -Path $versionRoot -Name "Class" -Value $className -Force | Out-Null
New-ItemProperty -Path $versionRoot -Name "Assembly" -Value $assembly.FullName -Force | Out-Null
New-ItemProperty -Path $versionRoot -Name "RuntimeVersion" -Value "v4.0.30319" -Force | Out-Null
New-ItemProperty -Path $versionRoot -Name "CodeBase" -Value $codeBase -Force | Out-Null

New-Item -Path (Join-Path $clsidRoot "ProgId") -Force | Out-Null
New-ItemProperty -Path (Join-Path $clsidRoot "ProgId") -Name "(default)" -Value $progId -Force | Out-Null
New-Item -Path $implementedCategoriesRoot -Force | Out-Null

New-Item -Path $officeRoot -Force | Out-Null
New-ItemProperty -Path $officeRoot -Name "FriendlyName" -Value "Apurisk" -Force | Out-Null
New-ItemProperty -Path $officeRoot -Name "Description" -Value "Herramientas de estadistica y gestion de riesgos. Primer modulo: Analisis BowTie." -Force | Out-Null
New-ItemProperty -Path $officeRoot -Name "LoadBehavior" -Value 3 -PropertyType DWord -Force | Out-Null

Write-Host "Apurisk registrado para el usuario actual. Abre Excel y busca la pestana Apurisk."
