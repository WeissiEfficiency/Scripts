# Data-Vergleich – Plan

PowerShell-Script, das zwei Ordner per **SHA256-Hash** vergleicht und Dubletten
**schrittweise und kontrolliert** aus dem Zielordner löscht.

Script: `Compare-And-Remove-Duplicates.ps1`

## Entscheidungen

| Thema              | Entscheidung                                                                                 |
|--------------------|----------------------------------------------------------------------------------------------|
| Rollen der Ordner  | **Vergleichsordner** wird vorab festgelegt und nie verändert. **Zielordner** wird nach Dubletten durchsucht, nur dort wird gelöscht. |
| Dublette           | Gleiche Größe **und** gleicher SHA256-Hash. Name, Unterordner und Zeitstempel egal (Zeitstempel ändern sich beim Kopieren/Hochladen). |
| Ordnerstruktur     | Datei darf im Zielordner in einem anderen Unterordner / mit anderem Namen liegen.            |
| Löschen            | Papierkorb (nur lokale Laufwerke). Optional `-NoRecycleBin`.                                 |
| Pakete             | Max. 10 GB pro Paket, danach exakt `CONTINUE` eintippen.                                     |
| Leere Ordner       | Werden nach dem Löschen im Zielordner entfernt (Zielordner selbst bleibt).                   |
| DryRun             | `-DryRun` simuliert den kompletten Ablauf ohne zu löschen.                                   |

## Ablauf

1. **Validierung** – beide Ordner existieren, sind nicht identisch und nicht ineinander verschachtelt.
2. **Einlesen** – alle Dateien beider Ordner rekursiv (inkl. versteckter Dateien).
3. **Vergleich** – Vorfilter über Dateigröße (nur Dateien mit passender Größe werden gehasht),
   danach SHA256. 0-Byte-Dateien werden standardmäßig ignoriert (`-IncludeEmptyFiles`).
4. **Report** – Zusammenfassung in der Konsole + CSV `Logs\Vergleich_<Zeit>.csv`, dann Frage J/N.
5. **Doppelte Bestätigung** – (1) J/N ob aus dem Zielordner gelöscht werden soll,
   (2) Name des Zielordners eintippen. Stimmt er nicht, wird abgebrochen.
6. **Pakete** – Dubletten in Pakete ≤ 10 GB aufteilen (Datei > 10 GB = eigenes Paket).
7. **Löschen** – vor jeder Datei wird geprüft, ob Referenz- und Zieldatei noch existieren und seit dem
   Scan unverändert sind (Größe + Änderungszeit), sonst Überspringen. Nach jedem Paket: `CONTINUE`.
8. **Fehler** – beim ersten Fehler Stopp; weiter nur mit `CONTINUE`.
9. **Aufräumen** – leer gewordene Unterordner im Zielordner entfernen, Zusammenfassung, CSV-Log
   `Logs\Loeschung_<Zeit>.csv` mit jeder Aktion.

## Status

1. [x] Plan + Ordner `Data-Vergleich`
2. [x] Validierung, Einlesen, Hash-Vergleich, Report
3. [x] Doppelte Bestätigung
4. [x] Paket-Löschung mit 10-GB-Grenze, CONTINUE, Fehler-Stopp
5. [x] Leere Ordner entfernen
6. [x] Test mit Testdaten (DryRun, Abbruch, falsche Bestätigung, Paketgrenze, verschachtelte Ordner)
7. [x] DryRun auf Windows mit echten Daten (v1)
8. [x] v2: SHA256 beider Dateien explizit vergleichen + erneute Hash-Prüfung direkt vor dem Löschen
9. [ ] v2 auf Windows testen – erst `-DryRun`, dann kleiner Teilordner
