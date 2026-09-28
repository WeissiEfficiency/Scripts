# Data-Vergleich – Plan

PowerShell-Script, das zwei Ordner anhand von Metadaten vergleicht und doppelte Dateien
**schrittweise und kontrolliert** aus einem der beiden Ordner löscht.

Geplanter Dateiname: `Data-Vergleich/Compare-And-Remove-Duplicates.ps1`

---

## 1. Parameter

| Parameter        | Typ      | Standard        | Beschreibung                                                        |
|------------------|----------|-----------------|---------------------------------------------------------------------|
| `-FolderA`       | string   | (Pflicht)       | Erster Ordner                                                       |
| `-FolderB`       | string   | (Pflicht)       | Zweiter Ordner                                                      |
| `-MatchMode`     | enum     | `RelativePath`  | `RelativePath` = gleicher Unterpfad + Name, `NameOnly` = nur Dateiname |
| `-VerifyHash`    | switch   | aus             | Zusätzlich SHA256-Hash vergleichen (langsamer, aber sicherer)       |
| `-MaxBatchGB`    | int      | `10`            | Maximale Datenmenge pro Lösch-Durchgang                             |
| `-UseRecycleBin` | switch   | an              | In den Papierkorb statt endgültig löschen                           |
| `-DryRun`        | switch   | aus             | Nur anzeigen, was gelöscht würde – nichts löschen                   |
| `-LogPath`       | string   | `.\Logs`        | Ablage für Vergleichs-Report und Lösch-Log (CSV)                    |

## 2. Ablauf (Schritt für Schritt)

### Schritt 1 – Validierung
- Beide Ordner müssen existieren.
- Die Ordner dürfen **nicht identisch** sein und **nicht ineinander verschachtelt** (sonst würde man
  Dateien mit sich selbst vergleichen und evtl. das "Original" löschen).
- Log-Ordner anlegen, Log-Datei mit Zeitstempel starten (`Start-Transcript` + eigenes CSV).

### Schritt 2 – Metadaten einlesen
- `Get-ChildItem -Recurse -File` für beide Ordner.
- Pro Datei erfassen: `Name`, `RelativePath`, `Length` (Größe), `LastWriteTime`, `CreationTime`, `Extension`.
- Fortschrittsanzeige mit `Write-Progress` (bei großen Datenmengen wichtig).

### Schritt 3 – Vergleich
Eine Datei gilt als **doppelt**, wenn übereinstimmen:
1. Schlüssel (je nach `MatchMode`: relativer Pfad **oder** Dateiname)
2. Größe (`Length`) – exakt
3. Änderungsdatum (`LastWriteTime`) – mit Toleranz von 2 Sekunden (FAT/NTFS/Netzlaufwerk-Unterschiede)
4. Optional (`-VerifyHash`): SHA256-Hash identisch

Umsetzung über eine Hashtable (Schlüssel -> Datei) für schnellen Lookup statt verschachtelter Schleifen.

### Schritt 4 – Report
- Übersicht in der Konsole: Anzahl Dateien A / B, Anzahl Duplikate, Gesamtgröße der Duplikate.
- Vollständige Liste als CSV in `Logs\Vergleich_<Datum>.csv` (zum Prüfen vor dem Löschen).
- Frage: "Möchtest du mit dem Löschen fortfahren? (J/N)" – bei N Ende.

### Schritt 5 – Ordnerauswahl (doppelte Abfrage)
1. **Erste Abfrage:** "Aus welchem Ordner soll gelöscht werden? [A] <Pfad A> / [B] <Pfad B>"
2. **Zweite Abfrage:** Bestätigung, bei der der Ordner **erneut** gewählt werden muss
   (z. B. vollständigen Pfad oder Buchstaben nochmal eintippen).
3. Stimmen beide Antworten nicht überein -> Abbruch, nichts wird gelöscht.

### Schritt 6 – Löschen in Batches (max. 10 GB)
- Duplikate des gewählten Ordners werden in Pakete aufgeteilt, jedes Paket **≤ `MaxBatchGB`**.
  (Eine einzelne Datei > 10 GB bildet ein eigenes Paket und wird extra angekündigt.)
- Vor jedem Paket: Anzeige "Paket 1 von N – X Dateien – Y GB".
- **Vor jeder einzelnen Löschung** Sicherheitsprüfungen:
  - Gegenstück im anderen Ordner existiert noch? (Sonst wäre es keine Kopie mehr -> überspringen)
  - Größe/Datum unverändert seit dem Scan? (Sonst überspringen)
- Löschen per Papierkorb (`Microsoft.VisualBasic.FileIO.FileSystem.DeleteFile(..., SendToRecycleBin)`)
  oder bei `-UseRecycleBin:$false` per `Remove-Item`.
- Jede Aktion (gelöscht / übersprungen / Fehler) wird ins Lösch-Log (CSV) geschrieben.

### Schritt 7 – Continue-Einwilligung nach jedem Paket
- Nach jedem Paket (≤ 10 GB): Zusammenfassung (gelöscht, übersprungen, Fehler, bisher gelöschte GB).
- Nächstes Paket startet **nur**, wenn der Benutzer exakt `CONTINUE` eintippt.
  Jede andere Eingabe -> sauberer Abbruch.

### Schritt 8 – Fehlerbehandlung
- `try/catch` um jede Löschung.
- **Beim ersten Fehler wird das aktuelle Paket gestoppt** und der Benutzer gefragt, ob weitergemacht
  werden soll (standardmäßig Abbruch). So wird bei einem Problem nie viel auf einmal gelöscht.
- Abschluss-Zusammenfassung + Pfad zu den Log-Dateien.

## 3. Sicherheitsnetz (Zusammenfassung)

- `-DryRun` für einen kompletten Testlauf ohne Löschen
- Papierkorb statt endgültiges Löschen (Standard)
- Doppelte Ordnerauswahl
- Max. 10 GB pro Durchgang + `CONTINUE`-Einwilligung
- Prüfung vor jeder Löschung, ob das Gegenstück noch existiert
- Stopp beim ersten Fehler
- Vollständiges CSV-Log jeder Aktion

## 4. Umsetzungsschritte

1. [x] Plan + Ordner `Data-Vergleich` anlegen
2. [ ] Parameter, Validierung, Logging (Schritt 1)
3. [ ] Metadaten einlesen + Vergleich + CSV-Report (Schritte 2–4)
4. [ ] Doppelte Ordnerauswahl (Schritt 5)
5. [ ] Batch-Löschung mit 10-GB-Grenze, CONTINUE und Fehler-Stopp (Schritte 6–8)
6. [ ] Test mit `-DryRun` an Testordnern, danach echter Test mit kleinen Daten
7. [ ] README mit Beispielaufrufen

## 5. Offene Fragen

- **Vergleichsmodus:** Sollen Dateien nur als doppelt gelten, wenn sie im **gleichen Unterordner** liegen
  (`RelativePath`), oder reicht der **gleiche Dateiname** irgendwo im Ordner (`NameOnly`)?
- **Hash-Prüfung:** Standardmäßig an (sicherer, langsamer) oder aus (schneller)?
- **Papierkorb:** Bei Netzlaufwerken gibt es keinen Papierkorb – dort endgültig löschen oder abbrechen?
- **Leere Ordner:** Sollen nach dem Löschen leer gewordene Unterordner ebenfalls entfernt werden?
