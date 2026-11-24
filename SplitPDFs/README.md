# PDF aufteilen

Eine PDF per Rechtsklick in einzelne Seiten teilen.
Jede Seite wird im selben Ordner gespeichert. Die Ausgangsdatei bleibt unverändert.

Aus `Rechnung.pdf` mit 3 Seiten werden `Rechnung 01.pdf`, `Rechnung 02.pdf` und `Rechnung 03.pdf`.
Ist ein Name schon vergeben, wird eine weitere Zahl angehängt, zum Beispiel `Rechnung 01 1.pdf`.

## PdfSharp.dll

Das Skript lädt die vorhandene Datei

`C:\Projects\SandrasHobbyTools\PDF Tools\CombinePDFs\lib\PdfSharp.dll`

Es wird nichts heruntergeladen. Fehlt die Datei, erscheint eine Meldung mit diesem Pfad.

## Kontextmenü

1. Pfad in `SplitPDFs-contextmenu.reg` prüfen. Wenn der Ordner verschoben wird, den Skriptpfad in der Befehlszeile anpassen.
2. `SplitPDFs-contextmenu.reg` per Doppelklick importieren. Das kann Administratorrechte brauchen.
3. Eine PDF im Explorer mit der rechten Maustaste anklicken und **PDF aufteilen** wählen.

Bei Erfolg erscheint kein Fenster. Die neuen Dateien liegen neben der Ausgangsdatei.
Bei einem Fehler erscheint eine Meldung. Zusätzlich schreibt das Skript nach `SplitPDFs\splitpdf.log`.

## Test

```powershell
& ".\Test-SplitPDFs.ps1"
```

Der Test erzeugt eigene PDFs in einem temporären Ordner und löscht diesen Ordner danach wieder.
