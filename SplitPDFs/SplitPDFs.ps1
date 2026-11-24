# SplitPDFs.ps1
# Teilt eine PDF in einzelne Seiten.
# Aufruf: SplitPDFs.ps1 -Path "C:\Ordner\Rechnung.pdf" [-NoUi]
# Die Ausgangsdatei bleibt unverändert. Bei Erfolg erscheint kein Fenster.

param(
    [Parameter(Position = 0)]
    [string]$Path,

    [switch]$NoUi
)

$ErrorActionPreference = "Stop"
$script:scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$script:LogPath = Join-Path $script:scriptDir "splitpdf.log"
$script:Utf8NoBom = New-Object System.Text.UTF8Encoding $false

function Write-Log {
    param([string]$Message)
    $line = "[{0:yyyy-MM-dd HH:mm:ss}] {1}" -f (Get-Date), $Message
    try {
        [System.IO.File]::AppendAllText($script:LogPath, $line + "`r`n", $script:Utf8NoBom)
    } catch { }
}

function Show-Error {
    param([string]$Message)
    Write-Log "FEHLER: $Message"
    if ($NoUi) {
        Write-Host $Message
        return
    }
    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
        [void][System.Windows.Forms.MessageBox]::Show(
            $Message,
            "PDF aufteilen",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
    } catch {
        Write-Host $Message
    }
}

function Get-PdfSharpPath {
    return [System.IO.Path]::GetFullPath((Join-Path $script:scriptDir "..\CombinePDFs\lib\PdfSharp.dll"))
}

function Get-PageOutputPath {
    param(
        [string]$Directory,
        [string]$BaseName,
        [int]$PageNumber,
        [int]$PadWidth,
        [string]$SourcePath
    )

    $number = $PageNumber.ToString(("D" + $PadWidth))
    $sourceFull = [System.IO.Path]::GetFullPath($SourcePath)
    for ($n = 0; $n -le 999; $n++) {
        if ($n -eq 0) {
            $fileName = "{0} {1}.pdf" -f $BaseName, $number
        } else {
            $fileName = "{0} {1} {2}.pdf" -f $BaseName, $number, $n
        }
        $candidate = [System.IO.Path]::GetFullPath((Join-Path $Directory $fileName))
        if ([string]::Equals($candidate, $sourceFull, [System.StringComparison]::OrdinalIgnoreCase)) {
            continue
        }
        if (-not (Test-Path -LiteralPath $candidate)) {
            return $candidate
        }
    }
    return $null
}

Write-Log "SplitPDFs gestartet. Path=$Path | NoUi=$NoUi"

try {
    if (-not $Path) {
        Show-Error "Es wurde kein PDF-Pfad übergeben."
        exit 1
    }

    try {
        $full = (Resolve-Path -LiteralPath $Path -ErrorAction Stop).Path
    } catch {
        Show-Error "Die Datei wurde nicht gefunden: $Path"
        exit 1
    }

    if ([System.IO.Path]::GetExtension($full) -ne ".pdf") {
        Show-Error "Nur PDF-Dateien können aufgeteilt werden: $full"
        exit 1
    }

    $dllPath = Get-PdfSharpPath
    if (-not (Test-Path -LiteralPath $dllPath)) {
        Show-Error "PdfSharp.dll fehlt. Erwartet unter: $dllPath"
        exit 1
    }
    try {
        Add-Type -Path $dllPath
    } catch {
        Show-Error "PdfSharp.dll konnte nicht geladen werden: $dllPath. $($_.Exception.Message)"
        exit 1
    }

    $inputPdf = $null
    try {
        try {
            $inputPdf = [PdfSharp.Pdf.IO.PdfReader]::Open($full, [PdfSharp.Pdf.IO.PdfDocumentOpenMode]::Import)
        } catch {
            Show-Error "Die PDF konnte nicht gelesen werden: $full. $($_.Exception.Message)"
            exit 1
        }

        $pageCount = $inputPdf.PageCount
        Write-Log "Geöffnet: $full ($pageCount Seiten)"
        if ($pageCount -eq 0) {
            Show-Error "Die PDF hat keine Seiten: $full"
            exit 1
        }

        $padWidth = [Math]::Max(2, $pageCount.ToString().Length)
        $directory = [System.IO.Path]::GetDirectoryName($full)
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($full)

        for ($page = 1; $page -le $pageCount; $page++) {
            $outputPath = Get-PageOutputPath -Directory $directory -BaseName $baseName -PageNumber $page -PadWidth $padWidth -SourcePath $full
            if (-not $outputPath) {
                Show-Error "Für Seite $page wurde kein freier Dateiname gefunden."
                exit 1
            }

            $outputPdf = $null
            try {
                $outputPdf = New-Object PdfSharp.Pdf.PdfDocument
                [void]$outputPdf.AddPage($inputPdf.Pages[$page - 1])
                $outputPdf.Save($outputPath)
                Write-Log "Seite $page gespeichert: $outputPath"
            } catch {
                Show-Error "Seite $page von $pageCount konnte nicht gespeichert werden: $outputPath. $($_.Exception.Message) Bereits geschriebene Seiten bleiben liegen."
                exit 1
            } finally {
                if ($outputPdf) { $outputPdf.Dispose() }
            }
        }
    } finally {
        if ($inputPdf) { $inputPdf.Dispose() }
    }

    Write-Log "SplitPDFs beendet: $pageCount Seiten aus $full"
    exit 0
} catch {
    Show-Error "Beim Aufteilen ist ein Fehler aufgetreten: $($_.Exception.Message)"
    exit 1
}
