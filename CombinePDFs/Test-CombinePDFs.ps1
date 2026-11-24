# Test: CombinePDFs mit festen Pfaden (1.pdf und 2.pdf in C:\Temp)
# Ausgabe 11.pdf landet in C:\Temp. PowerShell: cd C:\Projects\CombinePDFs dann .\Test-CombinePDFs.ps1

$scriptDir = $PSScriptRoot
$pdf1 = "C:\Temp\1.pdf"
$pdf2 = "C:\Temp\2.pdf"

if (-not (Test-Path $pdf1)) { Write-Host "FEHLER: $pdf1 fehlt. Bitte 1.pdf in C:\Temp anlegen." -ForegroundColor Red; exit 1 }
if (-not (Test-Path $pdf2)) { Write-Host "FEHLER: $pdf2 fehlt. Bitte 2.pdf in C:\Temp anlegen." -ForegroundColor Red; exit 1 }

# Schreibbarkeit des Zielordners testen (gleicher Ordner wie erste PDF)
$targetDir = [System.IO.Path]::GetDirectoryName($pdf1)
$testFile = Join-Path $targetDir "._combinepdf_write_test_$([Guid]::NewGuid().ToString('N')).tmp"
try {
    [System.IO.File]::WriteAllText($testFile, "test")
    Remove-Item -LiteralPath $testFile -Force
    Write-Host "Schreibtest OK: $targetDir" -ForegroundColor Green
} catch {
    Write-Host "FEHLER: In $targetDir kann nicht geschrieben werden: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

Write-Host "Test: Kombiniere $pdf1 und $pdf2 ..." -ForegroundColor Cyan
& (Join-Path $scriptDir "CombinePDFs.ps1") $pdf1 $pdf2
Write-Host "Fertig. Pruefe C:\Temp auf neue PDF (z. B. 11.pdf)." -ForegroundColor Green
