$ErrorActionPreference = "Stop"

$editorPath = "D:\JimRisk\RiskCode\Apurisk\tools\OfficeRibbonXEditor\OfficeRibbonXEditor.exe"

if (-not (Test-Path $editorPath)) {
    throw "No se encontro Office RibbonX Editor en $editorPath"
}

Start-Process -FilePath $editorPath
