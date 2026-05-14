$ErrorActionPreference = "Stop"

$resiliencyPaths = @(
    "HKCU:\Software\Microsoft\Office\16.0\Excel\Resiliency\DisabledItems",
    "HKCU:\Software\Microsoft\Office\16.0\Common\Resiliency\DisabledItems"
)

foreach ($path in $resiliencyPaths) {
    if (Test-Path $path) {
        Get-ChildItem $path -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
    }
}

Set-ItemProperty "HKCU:\Software\Microsoft\Office\Excel\Addins\Apurisk.ExcelAddIn" -Name "LoadBehavior" -Value 3
Write-Host "Se limpio el estado deshabilitado de Excel y Apurisk quedo habilitado otra vez."
