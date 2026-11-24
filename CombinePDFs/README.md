# Combine PDFs – Kontextmenü & Senden an

Mehrere PDFs per Rechtsklick oder „Senden an“ zu einer PDF zusammenführen.  
Ausgabename: **Name der ersten Datei + „ - combined.pdf“** (falls belegt: „ - combined 1.pdf“, „ - combined 2.pdf“ …).  
Die ursprünglichen Dateien werden **nicht gelöscht**.  
**Reihenfolge** im Kontextmenü: alphabetisch nach Pfad (determiniert). **Nur eine Kombination pro Ordner gleichzeitig** – zwei schnelle getrennte Auswahlen im selben Ordner können in einer Datei landen.

---

## Supply-Chain / Sicherheit

- **ImageToPdf** (deine Vorlage) erstellt PDFs **nur mit Windows-Bordmitteln**: .NET `System.Drawing` und der Druckertreiber **„Microsoft Print to PDF“**. Kein Download, keine Drittanbieter-Bibliotheken.
- **CombinePDFs** kann PDFs **nicht** nur mit Windows-Bordmitteln zusammenführen – Windows hat keine eingebaute API zum Mergen von PDFs. Das Skript verwendet deshalb **nur** die Datei **PdfSharp.dll** im Ordner `lib\`. Es werden **keine** Module aus dem Internet geladen (`Install-Module` ist nicht mehr vorgesehen). Du lädst die DLL **einmal selbst** von einer Quelle deiner Wahl und legst sie in `lib\` – so behältst du die Kontrolle über die einzige externe Abhängigkeit (z. B. offizielles Repo prüfen, Hash vergleichen).

---

## 1. PdfSharp.dll (einmalig)

- Ordner **`lib`** im Skriptordner anlegen.
- **PdfSharp.dll** in den Ordner `lib` legen.  
  Quelle z. B. offizielles Projekt [PDFsharp auf GitHub](https://github.com/empira/PDFsharp) (Build anlegen oder Release nutzen) oder [NuGet PdfSharp](https://www.nuget.org/packages/PdfSharp/) – du entscheidest die Quelle.

---

## 2. Kontextmenü (Rechtsklick auf PDFs)

1. **Skriptpfad in der Registry-Datei prüfen**  
   In `CombinePDFs-contextmenu.reg` den Pfad zu `CombinePDFs.ps1` anpassen (z. B. `C:\Projects\CombinePDFs` oder `C:\Scripts`).

2. **Registry-Eintrag setzen**  
   Doppelklick auf `CombinePDFs-contextmenu.reg` → Einträge übernehmen.

3. **Verwendung**  
   Eine oder mehrere PDFs im Explorer auswählen → Rechtsklick → **„Combine PDF“**.  
   Nach kurzer Ruhezeit (ca. 2 Sek ohne neue Dateien) erscheint im selben Ordner die kombinierte Datei (Name der ersten Datei + „ - combined.pdf“, bei Bedarf „ - combined 1.pdf“, „ - combined 2.pdf“ …). Bei Fehlern erscheint eine MessageBox; Log: **`combinepdf.log`** im Skriptordner.

---

## 3. Senden an (mehrere PDFs auf einmal)

1. SendTo-Ordner öffnen:  
   `Win + R` → `shell:sendto` → Enter.

2. **Verknüpfung** anlegen:  
   Rechtsklick → „Verknüpfung erstellen“.  
   Ziel: `C:\Projects\CombinePDFs\CombinePDFs-SendTo.cmd`  
   (Pfad ggf. anpassen.)  
   Name z. B. „Combine PDFs“.

3. **Verwendung**  
   Mehrere PDFs im Explorer auswählen → Rechtsklick → **„Senden an“** → **„Combine PDFs“**.  
   Die kombinierte PDF liegt im Ordner der ersten gewählten Datei.

---

## Dateien im Projekt

| Datei | Beschreibung |
|-------|--------------|
| `CombinePDFs.ps1` | Hauptskript (Kombinieren + Sammelmodus für Kontextmenü). |
| `CombinePDFs-contextmenu.reg` | Registry für Kontextmenü-Eintrag „Combine PDF“. |
| `CombinePDFs-SendTo.cmd` | Wrapper für „Senden an“ (ruft das Skript mit allen gewählten Dateien auf). |
| `lib\` | Hier **PdfSharp.dll** ablegen (einmalig von einer vertrauenswürdigen Quelle). |

**Logging / Fehler:** Fehler werden per **MessageBox** angezeigt (sichtbar bei Kontextmenü). Zusätzlich wird in **`combinepdf.log`** im Skriptordner protokolliert (Worker-Start, Anzahl Dateien, Zieldatei, Fehler).

---

## Nur PDFs

Aktuell werden nur **PDF-Dateien** unterstützt. Word-Dokumente würden eine zusätzliche Konvertierung (z. B. Word-COM oder anderes Tool) erfordern und sind hier nicht umgesetzt.
