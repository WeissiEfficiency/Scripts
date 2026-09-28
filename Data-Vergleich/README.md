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
