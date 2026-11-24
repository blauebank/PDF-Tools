# CombinePDFs.ps1
# Kombiniert mehrere PDFs zu einer Datei.
# Aufruf:
#   - Mehrere Pfade als Argumente (z. B. "Senden an"): kombiniert diese.
#   - -Collect "Pfad": Sammelmodus für Kontextmenü (pro Aufruf eine Queue-Datei, ein Worker mit Ruhezeit).
# Ausgabename: Name der ersten Datei + " - combined.pdf" (falls belegt: " - combined 1.pdf", " - combined 2.pdf" …). Originale werden nicht gelöscht.
# Hinweis: Pro Ordner nur eine Kombination gleichzeitig; bei zwei schnellen getrennten Auswahlen im selben Ordner können sie zusammenlaufen.

param(
    [Parameter(Position = 0, ValueFromRemainingArguments = $true)]
    [string[]]$Path,

    [switch]$Collect,

    [string]$QueueDir,
    [switch]$Worker,

    [switch]$DebugMode
)

$ErrorActionPreference = "Stop"
$script:scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$script:LogPath = if ($script:scriptDir) { Join-Path $script:scriptDir "combinepdf.log" } else { Join-Path $env:TEMP "combinepdf.log" }
$script:Utf8NoBom = New-Object System.Text.UTF8Encoding $false

# --------------- Logging ---------------
function Write-Log {
    param([string]$Message)
    $line = "[{0:yyyy-MM-dd HH:mm:ss}] {1}" -f (Get-Date), $Message
    try {
        [System.IO.File]::AppendAllText($script:LogPath, $line + "`r`n", $script:Utf8NoBom)
    } catch { }
}

function Write-LogException {
    param([string]$Kontext, [System.Management.Automation.ErrorRecord]$Err)
    if (-not $Err) { return }
    Write-Log "EXCEPTION [$Kontext]: $($Err.Exception.Message)"
    if ($Err.Exception.InnerException) {
        Write-Log "EXCEPTION [$Kontext] Inner: $($Err.Exception.InnerException.Message)"
    }
    if ($Err.ScriptStackTrace) {
        Write-Log "EXCEPTION [$Kontext] StackTrace: $($Err.ScriptStackTrace)"
    }
}

# Erste Log-Zeile: bestaetigt Start und Log-Pfad
$modus = if ($Worker -and $QueueDir) { "Worker" } elseif ($Collect) { "Collect" } else { "Direkt" }
Write-Log "CombinePDFs gestartet. LogPath=$script:LogPath | Modus=$modus | Path.Count=$($Path.Count) | QueueDir=$QueueDir | DebugMode=$DebugMode"

# --------------- Fehler anzeigen (sichtbar bei Kontextmenü) ---------------
function Show-Error {
    param([string]$Message, [string]$Title = "Combine PDFs")
    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
        [System.Windows.Forms.MessageBox]::Show($Message, $Title, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
    } catch {
        Write-Host $Message
    }
    Write-Log "FEHLER: $Message"
}

# --------------- Eingaben validieren: nur existierende PDFs, normalisiert, Duplikate raus ---------------
function Get-ValidatedPdfPaths {
    param([string[]]$RawPaths)
    Write-Log "Get-ValidatedPdfPaths: Eingabe $($RawPaths.Count) Pfad(e)"
    if ($null -eq $RawPaths -or $RawPaths.Count -eq 0) { Write-Log "Get-ValidatedPdfPaths: leer, return @()"; return @() }
    $seen = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $result = @()
    foreach ($p in $RawPaths) {
        $t = $p.Trim()
        if (-not $t) { continue }
        try {
            $full = (Resolve-Path -LiteralPath $t -ErrorAction Stop).Path
            if (-not (Test-Path -LiteralPath $full)) { Write-Log "Get-ValidatedPdfPaths: uebersprungen (nicht gefunden): $t"; continue }
            if ([System.IO.Path]::GetExtension($full) -ne '.pdf') { Write-Log "Get-ValidatedPdfPaths: uebersprungen (keine .pdf): $t"; continue }
            if ($seen.Add($full)) { $result += $full }
        } catch { Write-Log "Get-ValidatedPdfPaths: Resolve-Path fehlgeschlagen: $t - $($_.Exception.Message)"; Write-LogException "Get-ValidatedPdfPaths" $_ }
    }
    Write-Log "Get-ValidatedPdfPaths: Ergebnis $($result.Count) gueltige PDF(s)"
    return $result
}

# --------------- Zieldatei: Basis + " - combined.pdf", falls belegt " - combined 1.pdf", 2, 3, ... ---------------
function Get-OutputPath {
    param([string]$FirstPdfPath, [int]$StartFrom = 0)
    $dir = [System.IO.Path]::GetDirectoryName($FirstPdfPath)
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($FirstPdfPath)
    $suffix = " - combined"
    $n = $StartFrom
    while ($true) {
        if ($n -eq 0) {
            $candidate = Join-Path $dir ($baseName + $suffix + ".pdf")
        } else {
            $candidate = Join-Path $dir ($baseName + $suffix + " " + $n + ".pdf")
        }
        if (-not (Test-Path -LiteralPath $candidate)) {
            Write-Log "Get-OutputPath: gewaehlt $candidate"
            return $candidate
        }
        $n++
    }
}

# --------------- PdfSharp (nur lib\PdfSharp.dll) ---------------
function Test-PdfSharpAvailable {
    $dllPath = Join-Path $script:scriptDir "lib\PdfSharp.dll"
    if (-not (Test-Path -LiteralPath $dllPath)) {
        Write-Log "PdfSharp.dll nicht gefunden: $dllPath"
        return $false
    }
    try {
        Add-Type -Path $dllPath
        return $true
    } catch {
        Write-Log "PdfSharp.dll konnte nicht geladen werden: $($_.Exception.Message)"
        return $false
    }
}

function Merge-PdfsWithPdfSharp {
    param([string[]]$InputPaths, [string]$OutputPath)
    Write-Log "Merge-PdfsWithPdfSharp: Start mit $($InputPaths.Count) Dateien -> $OutputPath"
    $outputDocument = New-Object PdfSharp.Pdf.PdfDocument
    try {
        foreach ($inputFile in $InputPaths) {
            if (-not (Test-Path -LiteralPath $inputFile)) { Write-Log "Merge-PdfsWithPdfSharp: uebersprungen (nicht gefunden): $inputFile"; continue }
            try {
                $inputPdf = [PdfSharp.Pdf.IO.PdfReader]::Open($inputFile, [PdfSharp.Pdf.IO.PdfDocumentOpenMode]::Import)
                Write-Log "Merge-PdfsWithPdfSharp: geoeffnet $inputFile ($($inputPdf.PageCount) Seiten)"
            } catch {
                Write-Log "Datei gesperrt oder ungueltig: $inputFile - $($_.Exception.Message)"
                Write-LogException "Merge-PdfsWithPdfSharp Open" $_
                throw "Datei nicht lesbar (evtl. gesperrt): $inputFile"
            }
            try {
                for ($i = 0; $i -lt $inputPdf.PageCount; $i++) {
                    $outputDocument.AddPage($inputPdf.Pages[$i]) | Out-Null
                }
            } finally {
                $inputPdf.Dispose()
            }
        }
        if ($outputDocument.PageCount -eq 0) { Write-Log "Merge-PdfsWithPdfSharp: keine Seiten, return null"; return $null }
        try {
            Write-Log "Merge-PdfsWithPdfSharp: Save nach $OutputPath"
            $outputDocument.Save($OutputPath)
            Write-Log "Merge-PdfsWithPdfSharp: Save erfolgreich"
        } catch {
            $fullError = $_.Exception.Message
            if ($_.Exception.InnerException) { $fullError += " | " + $_.Exception.InnerException.Message }
            Write-Log "Zieldatei nicht schreibbar: $OutputPath - $fullError"
            Write-LogException "Merge-PdfsWithPdfSharp Save" $_
            try {
                $tempOut = Join-Path $env:TEMP ("combinepdf_" + [Guid]::NewGuid().ToString("N") + ".pdf")
                Write-Log "Merge-PdfsWithPdfSharp: Fallback Save nach Temp: $tempOut"
                $outputDocument.Save($tempOut)
                Write-Log "Merge-PdfsWithPdfSharp: Copy nach $OutputPath"
                [System.IO.File]::Copy($tempOut, $OutputPath, $true)
                Remove-Item -LiteralPath $tempOut -Force -ErrorAction SilentlyContinue
                return $OutputPath
            } catch {
                $fullError2 = $_.Exception.Message
                if ($_.Exception.InnerException) { $fullError2 += " | " + $_.Exception.InnerException.Message }
                Write-Log "Auch Temp-Speicher fehlgeschlagen: $fullError2"
                Write-LogException "Merge-PdfsWithPdfSharp Fallback" $_
                throw "Zieldatei nicht schreibbar: $OutputPath - $fullError"
            }
        }
        return $OutputPath
    } finally {
        $outputDocument.Dispose()
    }
}

# --------------- Merge ausführen (Validierung + Zielname + Merge) ---------------
function Invoke-Merge {
    param([string[]]$InputPaths)
    Write-Log "Invoke-Merge: aufgerufen mit $($InputPaths.Count) Pfad(en)"
    $paths = Get-ValidatedPdfPaths -RawPaths $InputPaths
    if ($paths.Count -lt 2) {
        Write-Log "Invoke-Merge: Abbruch, weniger als 2 gueltige PDFs ($($paths.Count))"
        Show-Error "Mindestens 2 gueltige PDF-Dateien noetig. Gefunden: $($paths.Count)."
        return
    }
    Write-Log "Invoke-Merge: Test-PdfSharpAvailable aufrufen"
    if (-not (Test-PdfSharpAvailable)) {
        Write-Log "Invoke-Merge: PdfSharp nicht verfuegbar"
        Show-Error "PdfSharp.dll fehlt. Bitte DLL aus vertrauenswuerdiger Quelle (z. B. GitHub PDFsharp) in den Ordner lib\ legen. Siehe README."
        return
    }
    $outputPath = Get-OutputPath -FirstPdfPath $paths[0]
    Write-Log "Invoke-Merge: Merge starten - $($paths.Count) Dateien -> $outputPath"
    $n = 1
    while ($true) {
        try {
            $null = Merge-PdfsWithPdfSharp -InputPaths $paths -OutputPath $outputPath
            Write-Log "Erfolg: $outputPath"
            break
        } catch {
            if ($_.Exception.Message -match "nicht schreibbar|gesperrt|Access|Zugriff" -and $n -lt 20) {
                $outputPath = Get-OutputPath -FirstPdfPath $paths[0] -StartFrom $n
                Write-Log "Invoke-Merge: Zieldatei gesperrt, naechster Versuch: $outputPath"
                $n++
            } else {
                Write-Log "Invoke-Merge: Fehler (kein Retry): $($_.Exception.Message)"
                Write-LogException "Invoke-Merge" $_
                Show-Error $_.Exception.Message
                break
            }
        }
    }
}

# --------------- Queue-Verzeichnis: Hash aus Ordnerpfad ---------------
function Get-QueueDirForPath {
    param([string]$FilePath)
    Write-Log "Get-QueueDirForPath: Start FilePath=$FilePath"
    try {
    $dir = [System.IO.Path]::GetDirectoryName($FilePath)
    $hash = [BitConverter]::ToString(
        [System.Security.Cryptography.SHA1]::Create().ComputeHash(
            [Text.Encoding]::UTF8.GetBytes($dir)
        )
    ).Replace("-", "").Substring(0, 12)
    $queueDir = [System.IO.Path]::Combine($env:TEMP, "combinepdf", $hash)
    Write-Log "Get-QueueDirForPath: $dir -> $queueDir"
    return $queueDir
    } catch {
        Write-LogException "Get-QueueDirForPath" $_
        throw
    }
}

# --------------- Collector: eine Datei pro Pfad schreiben, ggf. Worker starten ---------------
function Add-ToQueueAndMaybeStartWorker {
    param([string]$SinglePath)
    Write-Log "Collect: Add-ToQueueAndMaybeStartWorker Start SinglePath=$SinglePath"
    try {
    Write-Log "Collect gestartet: $SinglePath"
    $resolvedPath = $null
    try {
        $resolvedPath = (Resolve-Path -LiteralPath $SinglePath -ErrorAction Stop).Path
        Write-Log "Pfad aufgeloest: $resolvedPath"
    } catch {
        Write-LogException "Collect Resolve-Path" $_
        Write-Log "Resolve-Path fehlgeschlagen: $SinglePath - $($_.Exception.Message)"
        return
    }
    Write-Log "Collect: Pruefe .pdf und Test-Path"
    if ([System.IO.Path]::GetExtension($resolvedPath) -ne '.pdf') {
        Write-Log "Uebersprungen (keine .pdf): $resolvedPath"
        return
    }
    if (-not (Test-Path -LiteralPath $resolvedPath)) {
        Write-Log "Uebersprungen (Datei nicht gefunden): $resolvedPath"
        return
    }
    Write-Log "Collect: rufe Get-QueueDirForPath auf"
    $queueDir = Get-QueueDirForPath -FilePath $resolvedPath
    Write-Log "Collect: QueueDir=$queueDir"
    if (-not (Test-Path -LiteralPath $queueDir)) {
        Write-Log "Collect: erstelle Queue-Verzeichnis"
        New-Item -ItemType Directory -Path $queueDir -Force | Out-Null
    } else {
        $dirInfo = Get-Item -LiteralPath $queueDir -ErrorAction SilentlyContinue
        if ($dirInfo -and $dirInfo.CreationTime -lt (Get-Date).AddMinutes(-10)) {
            Write-Log "Collect: Stale Queue (>10 Min), bereinige"
            try {
                Get-ChildItem -LiteralPath $queueDir -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
                Remove-Item -LiteralPath $queueDir -Force -Recurse -ErrorAction SilentlyContinue
            } catch { }
            New-Item -ItemType Directory -Path $queueDir -Force | Out-Null
        }
    }
    $uniqueFile = Join-Path $queueDir ([Guid]::NewGuid().ToString("N") + ".txt")
    try {
        [System.IO.File]::WriteAllText($uniqueFile, $resolvedPath, $script:Utf8NoBom)
        Write-Log "Queue geschrieben: $uniqueFile -> $resolvedPath"
    } catch {
        Write-Log "Queue-Datei nicht schreibbar: $uniqueFile - $($_.Exception.Message)"
        Write-LogException "Collect Queue Write" $_
        return
    }

    $mutexName = "Global\CombinePDFs_" + [System.IO.Path]::GetFileName($queueDir)
    Write-Log "Collect: Mutex erstellen: $mutexName"
    $mutex = $null
    try {
        $mutex = [System.Threading.Mutex]::new($false, $mutexName)
        $owned = $mutex.WaitOne(0, $false)
        Write-Log "Collect: Mutex WaitOne: owned=$owned"
        if ($owned) {
            $scriptPath = if ($script:scriptDir) { Join-Path $script:scriptDir "CombinePDFs.ps1" } else { $MyInvocation.MyCommand.Path }
            $workerArgs = "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`" -Worker -QueueDir `"$queueDir`""
            if ($DebugMode) {
                $workerArgs = "-NoExit -NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`" -Worker -QueueDir `"$queueDir`" -DebugMode"
            }
            $windowStyle = if ($DebugMode) { "Normal" } else { "Hidden" }
            $psExe = Join-Path $PSHOME "powershell.exe"
            Write-Log "Starte Worker: $psExe (ArgumentList Laenge: $($workerArgs.Length))"
            Start-Process -FilePath $psExe -ArgumentList $workerArgs -WindowStyle $windowStyle
            Write-Log "Collect: Start-Process ausgefuehrt"
        } else {
            Write-Log "Mutex nicht erlangt - Worker wurde bereits von anderem Collector gestartet"
        }
    } catch {
        Write-Log "Mutex/Worker Exception: $($_.Exception.Message)"
        Write-LogException "Collect Mutex/Worker" $_
    } finally {
        if ($mutex) {
            try { $mutex.ReleaseMutex() } catch { }
            $mutex.Dispose()
        }
    }
    if ($DebugMode) {
        Write-Host "Collect fertig (Queue geschrieben)."
    }
    Write-Log "Collect: Add-ToQueueAndMaybeStartWorker Ende (erfolgreich)"
    } catch {
        Write-LogException "Collect Add-ToQueueAndMaybeStartWorker" $_
        Write-Log "Collect: Add-ToQueueAndMaybeStartWorker Ende (Fehler): $($_.Exception.Message)"
        throw
    }
}

# --------------- Worker: Ruhezeit-Schleife, dann lesen/sortieren/mergen/aufräumen ---------------
function Invoke-WorkerLoop {
    param([string]$QueueDirPath)
    Write-Log "Worker gestartet: $QueueDirPath"
    if ($DebugMode) { Write-Host "Worker gestartet. QueueDir: $QueueDirPath" }
    $quietSeconds = 2
    $intervalSeconds = 1
    $lastCount = -1
    $quietUntil = [DateTime]::MinValue

    Write-Log "Worker: Ruhezeit-Schleife starten (quietSeconds=$quietSeconds, intervalSeconds=$intervalSeconds)"
    try {
    while ($true) {
        Start-Sleep -Seconds $intervalSeconds
        $files = @()
        try {
            if (Test-Path -LiteralPath $QueueDirPath) {
                $files = Get-ChildItem -LiteralPath $QueueDirPath -Filter "*.txt" -ErrorAction SilentlyContinue
            }
        } catch { Write-Log "Worker: Fehler beim Lesen QueueDir: $($_.Exception.Message)"; Write-LogException "Worker QueueDir lesen" $_ }
        $count = $files.Count
        if ($count -ne $lastCount) {
            $lastCount = $count
            $quietUntil = (Get-Date).AddSeconds($quietSeconds)
            Write-Log "Worker: Queue-Anzahl geaendert: count=$count, quietUntil=$quietUntil"
        }
        if ($count -gt 0 -and (Get-Date) -ge $quietUntil) {
            Write-Log "Worker: Ruhezeit erreicht, verlasse Schleife (count=$count)"
            break
        }
    }

    $files = @()
    if (Test-Path -LiteralPath $QueueDirPath) {
        $files = Get-ChildItem -LiteralPath $QueueDirPath -Filter "*.txt" -ErrorAction SilentlyContinue
    }
    Write-Log "Worker: Lese $($files.Count) Queue-Dateien"
    $paths = @()
    try {
        foreach ($f in $files) {
            $content = [System.IO.File]::ReadAllText($f.FullName, [System.Text.Encoding]::UTF8).Trim()
            if ($content) { $paths += $content }
        }
    } catch {
        Write-Log "Worker Lesefehler: $($_.Exception.Message)"
        Write-LogException "Worker Queue-Dateien lesen" $_
    }

    Write-Log "Worker gelesene Pfade aus Queue: $($paths -join ' | ')"
    $paths = $paths | Sort-Object -Unique
    $paths = Get-ValidatedPdfPaths -RawPaths $paths
    Write-Log "Worker gesammelt: $($paths.Count) gueltige PDFs (Reihenfolge: alphabetisch nach Pfad)"
    if ($DebugMode) { Write-Host "Gefunden: $($paths.Count) gueltige PDF(s). Pfade: $($paths -join ', ')" }

    try {
        if ($paths.Count -ge 2) {
            Write-Log "Worker: rufe Invoke-Merge auf mit $($paths.Count) Pfaden"
            if ($DebugMode) { Write-Host "Starte Merge..." }
            Invoke-Merge -InputPaths $paths
            if ($DebugMode) { Write-Host "Merge abgeschlossen." }
        } else {
            Write-Log "Worker: weniger als 2 gueltige PDFs, Abbruch."
            if ($DebugMode) { Write-Host "Abbruch: weniger als 2 gueltige PDFs." }
        }
    } finally {
        try {
            if (Test-Path -LiteralPath $QueueDirPath) {
                Write-Log "Worker: Cleanup QueueDir $QueueDirPath"
                Get-ChildItem -LiteralPath $QueueDirPath -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
                Remove-Item -LiteralPath $QueueDirPath -Force -Recurse -ErrorAction SilentlyContinue
            }
        } catch {
            Write-Log "Worker Cleanup: $($_.Exception.Message)"
            Write-LogException "Worker Cleanup" $_
        }
        Write-Log "Worker beendet."
    }
        } catch {
            Write-LogException "Worker Invoke-WorkerLoop" $_
            Write-Log "Worker Ausnahme: $($_.Exception.Message)"
            if ($_.Exception.InnerException) { Write-Log "Worker InnerException: $($_.Exception.InnerException.Message)" }
            Show-Error "Beim Kombinieren ist ein Fehler aufgetreten. Siehe Log: $script:LogPath"
        }
}

# =============== Modi ===============

try {

# Worker-Modus
if ($Worker -and $QueueDir) {
    Write-Log "Worker-Modus gestartet. QueueDir=$QueueDir"
    if ($DebugMode) { Write-Host "Combine PDFs (Debug) - Worker startet..."; Write-Host "" }
    try {
        Invoke-WorkerLoop -QueueDirPath $QueueDir
    } finally {
        if ($DebugMode) {
            Write-Host ""
            Write-Host "Log: $script:LogPath"
            Read-Host "Druecken Sie Enter zum Schliessen"
        }
    }
    exit 0
}

# Sammelmodus (Kontextmenü)
if ($Collect -and $Path -and $Path.Count -eq 1) {
    if ($DebugMode) { Write-Host "Combine PDF (Debug): Datei wird gesammelt..."; Write-Host $Path[0] }
    try {
        Add-ToQueueAndMaybeStartWorker -SinglePath $Path[0]
    } catch {
        Write-LogException "Sammelmodus Add-ToQueue" $_
        Write-Log "Sammelmodus: Fehler beim Collect - $($_.Exception.Message)"
        throw
    } finally {
        if ($DebugMode) {
            Write-Host "Collect beendet. Druecken Sie Enter zum Schliessen."
            Read-Host
        }
    }
    Write-Log "Sammelmodus: Ende exit 0"
    exit 0
}

# Direktaufruf (z. B. Senden an)
Write-Log "Direktaufruf: Path.Count=$($Path.Count), args.Count=$($args.Count)"
$paths = @()
if ($Path) { $paths = @($Path) }
elseif ($args -and $args.Count -gt 0) { $paths = @($args) }
$paths = Get-ValidatedPdfPaths -RawPaths $paths
if ($paths.Count -lt 2) {
    Write-Log "Direktaufruf: Abbruch, weniger als 2 gueltige PDFs"
    Show-Error "Mindestens 2 gueltige PDF-Dateien noetig. Uebergeben: $($Path.Count + $args.Count), gueltig: $($paths.Count)."
    exit 1
}
Write-Log "Direktaufruf: rufe Invoke-Merge auf"
try {
Invoke-Merge -InputPaths $paths
Write-Log "CombinePDFs Skript Ende (Direktaufruf)"
} catch {
    Write-LogException "Direktaufruf Invoke-Merge" $_
    Write-Log "Direktaufruf: Fehler - $($_.Exception.Message)"
    throw
}

} catch {
    Write-LogException "SKRIPT TOP-LEVEL" $_
    Write-Log "SKRIPT ABBRUCH: Unbehandelte Ausnahme - $($_.Exception.Message)"
    throw
}
