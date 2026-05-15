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

function Add-VbaProjectToXlam {
    param([string]$sourceXlam, [string]$targetXlam)

    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $vbaBin = $null
    $zip = [System.IO.Compression.ZipFile]::Open($sourceXlam, 'Read')
    try {
        $entry = $zip.GetEntry('xl/vbaProject.bin')
        if (-not $entry) { throw "xl/vbaProject.bin not found in source" }
        $ms = [System.IO.MemoryStream]::new()
        $entry.Open().CopyTo($ms)
        $vbaBin = $ms.ToArray()
    }
    finally {
        $zip.Dispose()
    }

    $bak = $targetXlam + '.bak'
    Copy-Item $targetXlam $bak -Force

    try {
        $zip = [System.IO.Compression.ZipFile]::Open($targetXlam, 'Update')
        try {
            if (-not $zip.GetEntry('xl/vbaProject.bin')) {
                $entry = $zip.CreateEntry('xl/vbaProject.bin')
                $sw = [System.IO.StreamWriter]::new($entry.Open())
                $sw.BaseStream.Write($vbaBin, 0, $vbaBin.Length)
                $sw.Dispose()
            }

            $ctEntry = $zip.GetEntry('[Content_Types].xml')
            if ($ctEntry) {
                $r = [System.IO.StreamReader]::new($ctEntry.Open())
                $ct = $r.ReadToEnd()
                $r.Dispose()
                if ($ct -notmatch 'vbaProject') {
                    $insert = '<Override PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/>'
                    $ct = $ct -replace '</Types>', "  $insert`r`n</Types>"
                    $w = [System.IO.StreamWriter]::new($zip.GetEntry('[Content_Types].xml').Open())
                    $w.Write($ct)
                    $w.Dispose()
                }
            }

            $relsEntry = $zip.GetEntry('xl/_rels/workbook.xml.rels')
            if ($relsEntry) {
                $r = [System.IO.StreamReader]::new($relsEntry.Open())
                $rels = $r.ReadToEnd()
                $r.Dispose()
                if ($rels -notmatch 'vbaProject') {
                    $insertRel = '<Relationship Id="rIdVba" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/>'
                    $rels = $rels -replace '</Relationships>', "  $insertRel`r`n</Relationships>"
                    $w = [System.IO.StreamWriter]::new($zip.GetEntry('xl/_rels/workbook.xml.rels').Open())
                    $w.Write($rels)
                    $w.Dispose()
                }
            }
        }
        finally {
            $zip.Dispose()
        }
        Remove-Item $bak -Force
        Write-Host "VBA project injected."
    }
    catch {
        Move-Item $bak $targetXlam -Force
        throw
    }
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

    if (-not $vba -or -not $vba.VBComponents) {
        Write-Warning "El xlam no tiene proyecto VBA. Creando uno automaticamente..."
        $wb.Close()
        $tempPath = "$env:TEMP\_apurisk_vba_temp.xlam"
        if (Test-Path $tempPath) { Remove-Item $tempPath }
        $tempWb = $excel.Workbooks.Add()
        $tempWb.VBProject.VBComponents.Add(1) | Out-Null
        $tempWb.SaveAs($tempPath, 55)
        $tempWb.Close()
        Add-VbaProjectToXlam -sourceXlam $tempPath -targetXlam $xlamPath
        Remove-Item $tempPath -Force
        $wb = $excel.Workbooks.Open($xlamPath)
        $vba = $wb.VBProject
        if (-not $vba -or -not $vba.VBComponents) {
            throw "No se pudo crear el proyecto VBA en el xlam."
        }
        Write-Host "Proyecto VBA creado exitosamente."
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
        $wasVisible = $excel.Visible
        $excel.Visible = $true
        try {
            $dummyForm = $vba.VBComponents.Add(3)
            $vba.VBComponents.Remove($dummyForm)
            Start-Sleep -Milliseconds 200
        }
        catch {
            Write-Warning "No se pudo inicializar el disenador de formularios."
        }
        try {
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
        finally {
            $excel.Visible = $wasVisible
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
