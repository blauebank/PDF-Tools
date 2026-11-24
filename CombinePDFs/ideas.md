## Ideen & Erweiterungsplan

- **DOCX-Unterstuetzung via lokal installiertem Word**
  - DOCX-Dateien im Kontextmenue mit abdecken (zusaetzlicher Kontextmenue-Eintrag oder Erweiterung des bestehenden Eintrags fuer `.docx`).
  - Im Skript erkennen, ob Eingabedateien `.docx` sind; diese zunaechst nach `%TEMP%` als PDFs exportieren, danach wie normale PDFs behandeln.
  - Word ausschliesslich lokal per COM Automation ansprechen (`New-Object -ComObject Word.Application`) – kein Download, keine externen Module / Online-Dienste.
  - Word unsichtbar starten (`Visible = $false`) und fuer jeden DOCX:
    - Dokument oeffnen.
    - Als PDF nach `%TEMP%` exportieren (`ExportAsFixedFormat` / `wdExportFormatPDF`).
    - Dokument wieder schliessen.
  - Nach erfolgreichem Merge alle temporaeren DOCX-PDFs wieder loeschen (Cleanup in `finally`).
  - Reihenfolge definieren (z. B. erst alle Original-PDFs, dann alle konvertierten DOCX – oder gemischt nach urspruenglicher Auswahlreihenfolge).
  - Fehlerbehandlung:
    - Falls Word nicht installiert / COM-Objekt nicht erzeugbar: klare Fehlermeldung (MessageBox + Log) und Merge abbrechen oder nur reine PDFs verarbeiten.
    - Falls einzelne DOCX konvertiert werden koennen, andere nicht: wahlweise ueberspringen mit Warnung oder gesamten Vorgang abbrechen.
  - Ressourcenverwaltung:
    - `Quit()` fuer die Word-Instanz auch bei Fehlern sicherstellen (try/finally).
    - COM-Objekte explizit freigeben (`[System.Runtime.InteropServices.Marshal]::ReleaseComObject`), um haengende WINWORD.EXE-Prozesse zu vermeiden.
  - Testszenarien:
    - Nur DOCX, nur PDFs, Mischfaelle, gesperrte Dateien, defekte DOCX, fehlende Schreibrechte im Zielordner.
  - Keine Implementierung in diesem Schritt – nur Planung; Einbau erfolgt spaeter explizit und schrittweise.

