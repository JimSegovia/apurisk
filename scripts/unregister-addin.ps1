$ErrorActionPreference = "Stop"
$clsid = "{7BD16DC9-26B6-4C37-8E23-A4E80504D9E4}"
$progId = "Apurisk.ExcelAddIn"

Remove-Item "HKCU:\Software\Microsoft\Office\Excel\Addins\$progId" -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item "HKCU:\Software\Classes\$progId" -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item "HKCU:\Software\Classes\CLSID\$clsid" -Recurse -Force -ErrorAction SilentlyContinue

Write-Host "Apurisk desregistrado para el usuario actual."
