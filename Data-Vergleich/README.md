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
