# Test-SplitPDFs.ps1
# Erzeugt eigene PDFs, teilt sie und prüft Dateinamen, Seitenzahl und belegte Namen.

$ErrorActionPreference = "Stop"

$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$splitter = Join-Path $scriptDir "SplitPDFs.ps1"
$dllPath = [System.IO.Path]::GetFullPath((Join-Path $scriptDir "..\CombinePDFs\lib\PdfSharp.dll"))
$tempDir = Join-Path $env:TEMP ("splitpdf-test-" + [Guid]::NewGuid().ToString("N"))
$psExe = Join-Path $env:SystemRoot "System32\WindowsPowerShell\v1.0\powershell.exe"

function Write-Failure {
    param([string]$Message)
    Write-Host "FEHLER: $Message" -ForegroundColor Red
}

function Get-PdfPageCount {
    param([string]$PdfPath)
    $pdf = $null
    try {
        $pdf = [PdfSharp.Pdf.IO.PdfReader]::Open($PdfPath, [PdfSharp.Pdf.IO.PdfDocumentOpenMode]::Import)
        return $pdf.PageCount
    } finally {
        if ($pdf) { $pdf.Dispose() }
    }
}

function New-BlankPdf {
    param([string]$PdfPath, [int]$PageCount)
    $pdf = New-Object PdfSharp.Pdf.PdfDocument
    try {
        for ($i = 0; $i -lt $PageCount; $i++) {
            [void]$pdf.AddPage()
        }
        $pdf.Save($PdfPath)
    } finally {
        $pdf.Dispose()
    }
}

function Invoke-Splitter {
    param([string]$PdfPath)
    $argumentList = "-NoProfile -ExecutionPolicy Bypass -File `"$splitter`" -Path `"$PdfPath`" -NoUi"
    $process = Start-Process -FilePath $psExe -ArgumentList $argumentList -Wait -PassThru -NoNewWindow
    return $process.ExitCode
}

function Assert-PageCount {
    param([string]$PdfPath, [int]$Expected)
    if (-not (Test-Path -LiteralPath $PdfPath)) {
        throw "Datei fehlt: $PdfPath"
    }
    $actual = Get-PdfPageCount -PdfPath $PdfPath
    if ($actual -ne $Expected) {
        throw "$PdfPath hat $actual Seiten, erwartet $Expected."
    }
}

try {
    if (-not (Test-Path -LiteralPath $dllPath)) {
        Write-Failure "PdfSharp.dll fehlt: $dllPath. Es wird nichts heruntergeladen."
        exit 1
    }
    Add-Type -Path $dllPath

    if (-not (Test-Path -LiteralPath $splitter)) {
        Write-Failure "SplitPDFs.ps1 fehlt: $splitter"
        exit 1
    }
    if (-not (Test-Path -LiteralPath $psExe)) {
        Write-Failure "Windows PowerShell wurde nicht gefunden: $psExe"
        exit 1
    }

    New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
    $probe = Join-Path $tempDir "Probe.pdf"
    $einzeln = Join-Path $tempDir "Einzeln.pdf"
    New-BlankPdf -PdfPath $probe -PageCount 3

    $code = Invoke-Splitter -PdfPath $probe
    if ($code -ne 0) { throw "Erster Lauf endete mit Exitcode $code." }
    Assert-PageCount -PdfPath $probe -Expected 3
    foreach ($name in @("Probe 01.pdf", "Probe 02.pdf", "Probe 03.pdf")) {
        Assert-PageCount -PdfPath (Join-Path $tempDir $name) -Expected 1
    }

    $hashes = @{}
    foreach ($name in @("Probe 01.pdf", "Probe 02.pdf", "Probe 03.pdf")) {
        $full = Join-Path $tempDir $name
        $hashes[$name] = (Get-FileHash -LiteralPath $full -Algorithm SHA256).Hash
    }

    $code = Invoke-Splitter -PdfPath $probe
    if ($code -ne 0) { throw "Zweiter Lauf endete mit Exitcode $code." }
    foreach ($name in $hashes.Keys) {
        $full = Join-Path $tempDir $name
        $again = (Get-FileHash -LiteralPath $full -Algorithm SHA256).Hash
        if ($again -ne $hashes[$name]) { throw "$name wurde beim zweiten Lauf verändert." }
    }
    foreach ($name in @("Probe 01 1.pdf", "Probe 02 1.pdf", "Probe 03 1.pdf")) {
        Assert-PageCount -PdfPath (Join-Path $tempDir $name) -Expected 1
    }

    New-BlankPdf -PdfPath $einzeln -PageCount 1
    $code = Invoke-Splitter -PdfPath $einzeln
    if ($code -ne 0) { throw "Lauf für Einzeln.pdf endete mit Exitcode $code." }
    Assert-PageCount -PdfPath (Join-Path $tempDir "Einzeln 01.pdf") -Expected 1

    Write-Host "Test erfolgreich." -ForegroundColor Green
    exit 0
} catch {
    Write-Failure $_.Exception.Message
    exit 1
} finally {
    if (Test-Path -LiteralPath $tempDir) {
        Remove-Item -LiteralPath $tempDir -Recurse -Force -ErrorAction SilentlyContinue
    }
}
