$xlamPath = "D:\JimRisk\RiskCode\Apurisk.xlam"
$basDir = "$PSScriptRoot\..\src\Apurisk.XlamShell\vba"
$formsDir = "$PSScriptRoot\..\src\Apurisk.XlamShell\forms"

if (-not (Test-Path $xlamPath)) {
    Write-Error "No se encontro $xlamPath"
    exit 1
}

if (-not (Test-Path $basDir)) {
    Write-Error "No se encontro la carpeta VBA: $basDir"
    exit 1
}

if (-not (Test-Path $formsDir)) {
    Write-Warning "No se encontro la carpeta forms: $formsDir. Se omitira la importacion de formularios."
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$wb = $null

try {
    $wb = $excel.Workbooks.Open($xlamPath)

    try {
        $vba = $wb.VBProject
    }
    catch {
        Write-Error "No se puede acceder al proyecto VBA."
        Write-Host "Asegurate de habilitar en Excel:"
        Write-Host "  Archivo > Opciones > Centro de confianza > Configuracion"
        Write-Host "  > Configuracion de macros > Confiar en el acceso al"
        Write-Host "  modelo de objetos de proyectos de VBA"
        exit 1
    }

    Get-ChildItem $basDir -Filter *.bas | ForEach-Object {
        $name = $_.BaseName
        $existing = $vba.VBComponents | Where-Object { $_.Name -eq $name }
        if ($existing) {
            $vba.VBComponents.Remove($existing)
        }
        $vba.VBComponents.Import($_.FullName)
        Write-Host "Importado (modulo): $name"
    }

    if (Test-Path $formsDir) {
        Get-ChildItem $formsDir -Filter *.frm | ForEach-Object {
            $name = $_.BaseName
            $existing = $vba.VBComponents | Where-Object { $_.Name -eq $name }
            if ($existing) {
                $vba.VBComponents.Remove($existing)
            }
            $vba.VBComponents.Import($_.FullName)
            Write-Host "Importado (formulario): $name"
        }
    }

    $wb.Save()
    Write-Host "Guardado exitosamente."
}
catch {
    Write-Error $_.Exception.Message
}
finally {
    if ($wb) { $wb.Close() }
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}
