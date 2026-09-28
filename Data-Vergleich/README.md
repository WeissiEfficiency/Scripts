# Data-Vergleich

Findet Dateien im **Zielordner**, die (per SHA256-Hash) bereits im **Vergleichsordner** existieren,
und löscht sie schrittweise – max. 10 GB pro Durchgang, danach ist `CONTINUE` nötig.
Aus dem Vergleichsordner wird **nie** gelöscht.

## Verwendung

```powershell
# 1. Immer zuerst testen – es wird nichts gelöscht
.\Compare-And-Remove-Duplicates.ps1 -ReferenceFolder "D:\Original" -TargetFolder "E:\Kopie" -DryRun

# 2. Echter Lauf (Papierkorb)
.\Compare-And-Remove-Duplicates.ps1 -ReferenceFolder "D:\Original" -TargetFolder "E:\Kopie"
```

Ohne Parameter fragt das Script die Pfade ab. Falls Scripts blockiert sind:
`powershell -ExecutionPolicy Bypass -File .\Compare-And-Remove-Duplicates.ps1 ...`

## Parameter

| Parameter            | Standard  | Beschreibung                                          |
|----------------------|-----------|-------------------------------------------------------|
| `-ReferenceFolder`   | –         | Vergleichsordner, wird nie verändert                  |
| `-TargetFolder`      | –         | Hier wird nach Dubletten gesucht und gelöscht         |
| `-MaxBatchGB`        | `10`      | Max. Datenmenge pro Paket                             |
| `-DryRun`            | aus       | Nur simulieren                                        |
| `-NoRecycleBin`      | aus       | Endgültig löschen statt Papierkorb                    |
| `-KeepEmptyFolders`  | aus       | Leer gewordene Ordner nicht entfernen                 |
| `-IncludeEmptyFiles` | aus       | Auch 0-Byte-Dateien als Dubletten behandeln           |
| `-LogPath`           | `.\Logs`  | Ablage der CSV-Reports                                |

Hinweis: Dateien im Papierkorb belegen weiter Speicherplatz, bis der Papierkorb geleert wird.
Details zum Ablauf: siehe [PLAN.md](PLAN.md).

## Version 2 (`Compare-And-Remove-Duplicates-v2.ps1`)

Gleicher Ablauf wie Version 1, aber mit expliziter Hash-Prüfung beider Dateien:

- Der SHA256 wird für **beide** Dateien gebildet (Zieldatei und Vergleichsdatei) und direkt miteinander
  verglichen. Der Vergleichs-Report enthält beide Werte: `SHA256_Ziel`, `SHA256_Referenz`, `HashGleich`,
  dazu `GroesseZiel` / `GroesseReferenz`.
- **Unmittelbar vor jeder Löschung** werden beide Dateien erneut gehasht. Nur wenn beide frischen Hashes
  identisch sind (und zum Scan passen), wird gelöscht. Sonst: übersprungen + Eintrag im Lösch-Log.
- Das Lösch-Log enthält ebenfalls beide Hashes je Datei.
- Gibt es mehrere identische Vergleichsdateien, wird bevorzugt die mit gleichem Dateinamen als Gegenstück genommen.
- Logs heißen `Vergleich_v2_<Zeit>.csv` / `Loeschung_v2_<Zeit>.csv`.

Zusätzlicher Parameter: `-SkipRehash` schaltet das erneute Hashen vor dem Löschen ab (schneller, aber weniger sicher).
Hinweis: Durch das erneute Hashen werden beide Dateien vor dem Löschen noch einmal komplett gelesen –
bei 10 GB pro Paket also ca. 20 GB Lesezugriff.

```powershell
.\Compare-And-Remove-Duplicates-v2.ps1 -ReferenceFolder "E:\OpenCloud" -TargetFolder "F:\Daten" -DryRun
```

## Dubletten innerhalb von Ordnern (`Find-Duplicates.ps1`)

Sucht in **einem oder mehreren Ordnern** (inkl. aller Unterordner) nach Dateien mit gleichem Inhalt
(Größe + SHA256). Pro Dubletten-Gruppe wird **eine** Datei behalten, alle anderen werden gelöscht.
Ablauf und Sicherheitsnetz wie bei v2: Report, doppelte Bestätigung, Pakete ≤ 10 GB mit `CONTINUE`,
erneute Hash-Prüfung direkt vor dem Löschen, Stopp beim ersten Fehler, Papierkorb, `-DryRun`,
leere Ordner entfernen.

```powershell
# Ein Ordner
.\Find-Duplicates.ps1 -Path "F:\Daten" -DryRun

# Mehrere Ordner zusammen (Dubletten auch ordnerübergreifend)
.\Find-Duplicates.ps1 -Path "F:\Daten", "E:\Fotos" -DryRun
.\Find-Duplicates.ps1 -Path "F:\Daten;E:\Fotos" -DryRun
```

| Parameter        | Standard       | Beschreibung                                                             |
|------------------|----------------|--------------------------------------------------------------------------|
| `-Path`          | –              | Ein oder mehrere Ordner (Liste oder mit `;` getrennt)                     |
| `-KeepRule`      | `ShortestPath` | Welche Datei pro Gruppe bleibt: `ShortestPath` (kürzester Pfad, z. B. `Bilder` statt `Bilder - Kopie`), `Oldest`, `Newest` (nach Änderungsdatum), `Alphabetical` |
| `-ExcludeFolder` | `.git`         | Ordnernamen, die komplett übersprungen werden. `-ExcludeFolder @()` = nichts ausschließen |
| `-MaxBatchGB`, `-DryRun`, `-NoRecycleBin`, `-KeepEmptyFolders`, `-IncludeEmptyFiles`, `-SkipRehash`, `-LogPath` | | wie bei v2 |

**Bestätigung:** (1) J/N, (2) die angezeigte Anzahl der zu löschenden Dateien eintippen.

**Report** `Logs\Dubletten_<Zeit>.csv`: Spalte `Gruppe` (gleiche Nummer = gleicher Inhalt),
`Aktion` (`Behalten` / `Loeschen`), `SHA256`, `BehalteneDatei`, `SHA256_Behalten`.
Vor dem echten Lauf prüfen, ob pro Gruppe die richtige Datei auf `Behalten` steht – sonst `-KeepRule` ändern.
