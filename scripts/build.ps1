param(
    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Debug"
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$csc = "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
$framework = "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319"
$out = Join-Path $root "build\$Configuration"
$extensibility = "C:\Program Files (x86)\Common Files\Microsoft Shared\MSEnv\PublicAssemblies\extensibility.dll"
$office = "C:\Windows\assembly\GAC_MSIL\office\15.0.0.0__71e9bce111e9429c\OFFICE.DLL"

if (-not (Test-Path $csc)) {
    throw "No se encontro el compilador C# de .NET Framework en $csc"
}

if (-not (Test-Path $extensibility)) {
    throw "No se encontro Extensibility.dll en $extensibility"
}

if (-not (Test-Path $office)) {
    throw "No se encontro OFFICE.dll en $office"
}

New-Item -ItemType Directory -Path $out -Force | Out-Null

function Invoke-Csc {
    param(
        [string]$Output,
        [string[]]$Sources,
        [string[]]$References
    )

    $args = @(
        "/nologo",
        "/target:library",
        "/debug+",
        "/optimize-",
        "/out:$Output"
    )

    foreach ($reference in $References) {
        $args += "/reference:$reference"
    }

    $args += $Sources
    & $csc $args

    if ($LASTEXITCODE -ne 0) {
        throw "Fallo compilando $Output"
    }
}

$commonReferences = @(
    "$framework\System.dll",
    "$framework\System.Core.dll"
)

$coreSources = Get-ChildItem "$root\src\Apurisk.Core" -Recurse -Filter *.cs | ForEach-Object { $_.FullName }
Invoke-Csc `
    -Output "$out\Apurisk.Core.dll" `
    -Sources $coreSources `
    -References $commonReferences

$applicationSources = Get-ChildItem "$root\src\Apurisk.Application" -Recurse -Filter *.cs | ForEach-Object { $_.FullName }
Invoke-Csc `
    -Output "$out\Apurisk.Application.dll" `
    -Sources $applicationSources `
    -References ($commonReferences + "$out\Apurisk.Core.dll")

$addinSources = Get-ChildItem "$root\src\Apurisk.ExcelAddIn" -Recurse -Filter *.cs | ForEach-Object { $_.FullName }
Invoke-Csc `
    -Output "$out\Apurisk.ExcelAddIn.dll" `
    -Sources $addinSources `
    -References ($commonReferences + @(
        "$framework\Microsoft.CSharp.dll",
        "$framework\System.Windows.Forms.dll",
        "$framework\System.Xml.dll",
        $extensibility,
        $office
    ))

Write-Host "Build completado en $out"
