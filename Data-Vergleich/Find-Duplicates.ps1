<#
.SYNOPSIS
    Sucht in einem oder mehreren Ordnern nach Dubletten (SHA256) und löscht sie schrittweise.

.DESCRIPTION
    Alle Dateien der angegebenen Ordner (inkl. aller Unterordner) werden verglichen. Dateien mit gleicher
    Größe UND gleichem SHA256-Hash bilden eine Dubletten-Gruppe. Name, Ordner und Zeitstempel sind egal.

    Pro Gruppe wird genau EINE Datei behalten (Regel: -KeepRule), alle anderen werden gelöscht.

    Sicherheitsnetz (wie Compare-And-Remove-Duplicates-v2.ps1):
    - Vergleichs-Report als CSV mit Gruppe, Aktion (Behalten/Loeschen) und SHA256 jeder Datei
    - Doppelte Bestätigung vor dem Löschen
    - Löschen in Paketen von max. -MaxBatchGB (Standard 10 GB), danach "CONTINUE" nötig
    - Vor jeder Löschung: behaltene Datei existiert noch, beide Dateien unverändert und beide werden
      NEU gehasht - nur bei identischem Hash wird gelöscht (abschaltbar mit -SkipRehash)
    - Stopp beim ersten Fehler
    - Standardmäßig Papierkorb statt endgültigem Löschen
    - -DryRun: kompletter Ablauf ohne Löschen
    - Leer gewordene Unterordner werden entfernt
    - Ordner namens ".git" werden standardmäßig ausgeschlossen (-ExcludeFolder)

.EXAMPLE
    .\Find-Duplicates.ps1 -Path "F:\Daten" -DryRun

.EXAMPLE
    .\Find-Duplicates.ps1 -Path "F:\Daten", "E:\Fotos" -KeepRule Oldest

.EXAMPLE
    .\Find-Duplicates.ps1 -Path "F:\Daten;E:\Fotos"
#>
[CmdletBinding()]
param(
    [string[]]$Path,
    [ValidateSet('ShortestPath', 'Oldest', 'Newest', 'Alphabetical')]
    [string]$KeepRule = 'ShortestPath',
    [string[]]$ExcludeFolder = @('.git'),
    [ValidateRange(1, 10000)]
    [int]$MaxBatchGB = 10,
    [switch]$DryRun,
    [switch]$NoRecycleBin,
    [switch]$KeepEmptyFolders,
    [switch]$IncludeEmptyFiles,
    [switch]$SkipRehash,
    [string]$LogPath = (Join-Path $PSScriptRoot 'Logs')
)

$ErrorActionPreference = 'Stop'
$MaxBatchBytes = [int64]$MaxBatchGB * 1GB
$Sep = [IO.Path]::DirectorySeparatorChar

# ---------------------------------------------------------------------------
# Hilfsfunktionen
# ---------------------------------------------------------------------------

function Format-Size([int64]$Bytes) {
    if ($Bytes -ge 1GB) { return '{0:N2} GB' -f ($Bytes / 1GB) }
    if ($Bytes -ge 1MB) { return '{0:N2} MB' -f ($Bytes / 1MB) }
    if ($Bytes -ge 1KB) { return '{0:N2} KB' -f ($Bytes / 1KB) }
    return "$Bytes B"
}

function Resolve-FolderPath([string]$Folder) {
    $Folder = $Folder.Trim().Trim('"')
    if (-not (Test-Path -LiteralPath $Folder -PathType Container)) {
        throw "Ordner '$Folder' existiert nicht oder ist kein Ordner."
    }
    $full = (Get-Item -LiteralPath $Folder).FullName
    return $full.TrimEnd('\', '/') + $Sep
}

function Test-IsInside([string]$Child, [string]$Parent) {
    return $Child.StartsWith($Parent, [StringComparison]::OrdinalIgnoreCase)
}

function Test-Excluded([string]$FullName, [string]$Root) {
    if ($ExcludeFolder.Count -eq 0) { return $false }
    $parts = $FullName.Substring($Root.Length).Split([char[]]@('\', '/'), [StringSplitOptions]::RemoveEmptyEntries)
    # Letzter Teil ist der Dateiname, nur Ordnerteile prüfen
    for ($p = 0; $p -lt $parts.Count - 1; $p++) {
        foreach ($ex in $ExcludeFolder) {
            if ($parts[$p] -ieq $ex) { return $true }
        }
    }
    return $false
}

function Get-FileSnapshot([string]$Root, $Files) {
    Write-Host "Lese Dateien ein: $Root ..."
    $count = 0; $excluded = 0
    Get-ChildItem -LiteralPath $Root -Recurse -File -Force -ErrorAction SilentlyContinue -ErrorVariable scanErrors |
        ForEach-Object {
            if (Test-Excluded $_.FullName $Root) { $excluded++; return }
            $count++
            if ($count % 500 -eq 0) { Write-Progress -Activity "Einlesen: $Root" -Status "$count Dateien" }
            $Files.Add([pscustomobject]@{
                FullName     = $_.FullName
                Root         = $Root
                Name         = $_.Name
                Length       = [int64]$_.Length
                LastWriteUtc = $_.LastWriteTimeUtc.Ticks
                LastWrite    = $_.LastWriteTime
                Hash         = $null
            })
        }
    Write-Progress -Activity "Einlesen: $Root" -Completed
    foreach ($e in $scanErrors) { Write-Warning "Nicht lesbar: $($e.TargetObject) - $($e.Exception.Message)" }
    Write-Host ("  {0} Dateien eingelesen, {1} ausgeschlossen ({2})" -f $count, $excluded, ($ExcludeFolder -join ', '))
}

function Get-Sha256([string]$File) {
    return (Get-FileHash -LiteralPath $File -Algorithm SHA256).Hash
}

function Add-Hashes($Files) {
    $total = $Files.Count
    $done = 0
    foreach ($f in $Files) {
        $done++
        Write-Progress -Activity 'SHA256 berechnen' -Status "$done / $total - $($f.Name)" -PercentComplete (($done / [math]::Max($total, 1)) * 100)
        try {
            $f.Hash = Get-Sha256 $f.FullName
        } catch {
            Write-Warning "Hash fehlgeschlagen: $($f.FullName) - $($_.Exception.Message)"
        }
    }
    Write-Progress -Activity 'SHA256 berechnen' -Completed
}

function Select-KeepFile($Group) {
    switch ($KeepRule) {
        'Oldest'       { return ($Group | Sort-Object LastWriteUtc, { $_.FullName.Length }, FullName | Select-Object -First 1) }
        'Newest'       { return ($Group | Sort-Object @{ Expression = 'LastWriteUtc'; Descending = $true }, @{ Expression = { $_.FullName.Length } }, @{ Expression = 'FullName' } | Select-Object -First 1) }
        'Alphabetical' { return ($Group | Sort-Object FullName | Select-Object -First 1) }
        default        { return ($Group | Sort-Object { $_.FullName.Length }, FullName | Select-Object -First 1) }
    }
}

function Test-Unchanged($Snapshot) {
    if (-not (Test-Path -LiteralPath $Snapshot.FullName -PathType Leaf)) { return $false }
    $item = Get-Item -LiteralPath $Snapshot.FullName -Force
    return ($item.Length -eq $Snapshot.Length -and $item.LastWriteTimeUtc.Ticks -eq $Snapshot.LastWriteUtc)
}

function Remove-SingleFile([string]$File) {
    if ($NoRecycleBin) {
        Remove-Item -LiteralPath $File -Force
    } else {
        [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile(
            $File,
            [Microsoft.VisualBasic.FileIO.UIOption]::OnlyErrorDialogs,
            [Microsoft.VisualBasic.FileIO.RecycleOption]::SendToRecycleBin)
    }
}

function Write-Log([string]$Action, [string]$File, [string]$Kept, [int64]$Size, [string]$Message,
                   [string]$FileHash = '', [string]$KeptHash = '') {
    [pscustomobject]@{
        Zeit                = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        Aktion              = $Action
        Datei               = $File
        BehalteneDatei      = $Kept
        GroesseByte         = $Size
        SHA256_Datei        = $FileHash
        SHA256_Behalten     = $KeptHash
        Info                = $Message
    } | Export-Csv -LiteralPath $script:DeleteLog -Append -NoTypeInformation -Delimiter ';' -Encoding UTF8
}

function Remove-EmptyFolders($Folders, $Roots) {
    $removed = 0
    $candidates = New-Object System.Collections.Generic.HashSet[string]([StringComparer]::OrdinalIgnoreCase)
    foreach ($dir in $Folders) {
        $root = $Roots | Where-Object { Test-IsInside ($dir.TrimEnd('\', '/') + $Sep) $_ } | Select-Object -First 1
        if (-not $root) { continue }
        # Alle Elternordner bis (exklusive) zum Root-Ordner als Kandidaten aufnehmen
        $d = $dir
        while ($d -and ($d.TrimEnd('\', '/') + $Sep) -ne $root -and (Test-IsInside ($d.TrimEnd('\', '/') + $Sep) $root)) {
            [void]$candidates.Add($d)
            $d = [IO.Path]::GetDirectoryName($d)
        }
    }
    foreach ($dir in ($candidates | Sort-Object Length -Descending)) {
        if (-not (Test-Path -LiteralPath $dir -PathType Container)) { continue }
        if (@(Get-ChildItem -LiteralPath $dir -Force).Count -eq 0) {
            try {
                Remove-Item -LiteralPath $dir -Force
                $removed++
                Write-Log -Action 'OrdnerEntfernt' -File $dir -Kept '' -Size 0 -Message 'Leerer Ordner'
            } catch {
                Write-Warning "Leerer Ordner konnte nicht entfernt werden: $dir - $($_.Exception.Message)"
            }
        }
    }
    return $removed
}

# ---------------------------------------------------------------------------
# Schritt 1 - Validierung
# ---------------------------------------------------------------------------

Write-Host ''
Write-Host '=== Dubletten-Suche per SHA256 ===' -ForegroundColor Cyan
if ($DryRun) { Write-Host '*** DRY RUN - es wird nichts gelöscht ***' -ForegroundColor Yellow }
Write-Host ''

if (-not $Path -or $Path.Count -eq 0) {
    $inputPaths = Read-Host 'Ordner eingeben (mehrere mit ; trennen)'
    $Path = $inputPaths.Split(';') | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
}
if (-not $Path -or @($Path).Count -eq 0) { throw 'Kein Ordner angegeben.' }

# Mehrere Ordner können auch als ein Text mit ; getrennt übergeben werden (z. B. beim Aufruf per -File)
$Path = @($Path | ForEach-Object { $_.Split(';') } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
$Roots = @($Path | ForEach-Object { Resolve-FolderPath $_ } | Sort-Object -Unique)
foreach ($a in $Roots) {
    foreach ($b in $Roots) {
        if ($a -ne $b -and (Test-IsInside $a $b)) {
            throw "Die Ordner '$a' und '$b' sind ineinander verschachtelt. Bitte nur den übergeordneten Ordner angeben."
        }
    }
}

if (-not $NoRecycleBin -and -not $DryRun) {
    Add-Type -AssemblyName Microsoft.VisualBasic
}

New-Item -ItemType Directory -Path $LogPath -Force | Out-Null
$stamp            = Get-Date -Format 'yyyyMMdd_HHmmss'
$Report           = Join-Path $LogPath "Dubletten_$stamp.csv"
$script:DeleteLog = Join-Path $LogPath "Dubletten_Loeschung_$stamp.csv"

Write-Host 'Durchsuchte Ordner:' -ForegroundColor Yellow
$Roots | ForEach-Object { Write-Host "  $_" }
Write-Host "Regel zum Behalten: $KeepRule"
Write-Host ''

# ---------------------------------------------------------------------------
# Schritt 2 - Einlesen
# ---------------------------------------------------------------------------

$allFiles = New-Object System.Collections.Generic.List[object]
foreach ($r in $Roots) { Get-FileSnapshot $r $allFiles }

# ---------------------------------------------------------------------------
# Schritt 3 - Vorfilter Größe, dann SHA256
# ---------------------------------------------------------------------------

$sizeGroups = $allFiles |
    Where-Object { $IncludeEmptyFiles -or $_.Length -gt 0 } |
    Group-Object Length |
    Where-Object { $_.Count -gt 1 }
$candidates = @($sizeGroups | ForEach-Object { $_.Group })

Write-Host ''
Write-Host ("Vorfilter nach Größe: {0} von {1} Dateien haben eine gleich große Datei." -f $candidates.Count, $allFiles.Count)
Write-Host 'Berechne SHA256-Hashes (kann je nach Datenmenge dauern) ...'
Add-Hashes $candidates

$hashGroups = @($candidates |
    Where-Object { $_.Hash } |
    Group-Object Length, Hash |
    Where-Object { $_.Count -gt 1 })

# Pro Gruppe eine Datei behalten, Rest zum Löschen vormerken
$toDelete = New-Object System.Collections.Generic.List[object]
$reportRows = New-Object System.Collections.Generic.List[object]
$groupNo = 0
foreach ($g in ($hashGroups | Sort-Object { (Select-KeepFile $_.Group).FullName })) {
    $groupNo++
    $keep = Select-KeepFile $g.Group
    foreach ($f in ($g.Group | Sort-Object FullName)) {
        $isKeep = [object]::ReferenceEquals($f, $keep)
        $action = 'Loeschen'
        if ($isKeep) { $action = 'Behalten' }
        $reportRows.Add([pscustomobject]@{
            Gruppe          = $groupNo
            Aktion          = $action
            Datei           = $f.FullName
            GroesseByte     = $f.Length
            SHA256          = $f.Hash
            Geaendert       = $f.LastWrite.ToString('yyyy-MM-dd HH:mm:ss')
            BehalteneDatei  = $keep.FullName
            SHA256_Behalten = $keep.Hash
        })
        if (-not $isKeep) {
            $toDelete.Add([pscustomobject]@{ File = $f; Keep = $keep; Group = $groupNo })
        }
    }
}

# ---------------------------------------------------------------------------
# Schritt 4 - Report
# ---------------------------------------------------------------------------

$reportRows | Export-Csv -LiteralPath $Report -NoTypeInformation -Delimiter ';' -Encoding UTF8
$delBytes = [int64](($toDelete | ForEach-Object { $_.File.Length } | Measure-Object -Sum).Sum)

Write-Host ''
Write-Host '=== Ergebnis ===' -ForegroundColor Cyan
Write-Host ("Dateien gesamt:          {0}" -f $allFiles.Count)
Write-Host ("Dubletten-Gruppen:       {0}" -f $hashGroups.Count)
Write-Host ("Zu löschende Dubletten:  {0} ({1})" -f $toDelete.Count, (Format-Size $delBytes)) -ForegroundColor Yellow
Write-Host "Report: $Report"
Write-Host '  (Spalte "Aktion": Behalten / Loeschen, Spalte "Gruppe": gleiche Nummer = gleicher Inhalt)'
Write-Host ''

if ($toDelete.Count -eq 0) {
    Write-Host 'Keine Dubletten gefunden. Nichts zu tun.' -ForegroundColor Green
    return
}

$answer = Read-Host 'Report prüfen. Mit dem Löschen fortfahren? (J/N)'
if ($answer -notmatch '^(j|ja|y|yes)$') {
    Write-Host 'Abgebrochen. Es wurde nichts gelöscht.'
    return
}

# ---------------------------------------------------------------------------
# Schritt 5 - Doppelte Bestätigung
# ---------------------------------------------------------------------------

Write-Host ''
Write-Host ("Es werden {0} Dateien ({1}) aus folgenden Ordnern gelöscht; pro Gruppe bleibt 1 Datei erhalten:" -f $toDelete.Count, (Format-Size $delBytes)) -ForegroundColor Yellow
$Roots | ForEach-Object { Write-Host "  $_" -ForegroundColor Yellow }
$confirm1 = Read-Host 'Bestätigung 1/2: Dubletten löschen? (J/N)'
if ($confirm1 -notmatch '^(j|ja|y|yes)$') {
    Write-Host 'Abgebrochen. Es wurde nichts gelöscht.'
    return
}
$confirm2 = Read-Host ("Bestätigung 2/2: Anzahl der zu löschenden Dateien eintippen ({0})" -f $toDelete.Count)
if ($confirm2.Trim() -ne [string]$toDelete.Count) {
    Write-Host 'Eingabe stimmt nicht. Abgebrochen, es wurde nichts gelöscht.' -ForegroundColor Red
    return
}

# ---------------------------------------------------------------------------
# Schritt 6 - Pakete bilden (max. MaxBatchGB)
# ---------------------------------------------------------------------------

$batches = New-Object System.Collections.Generic.List[object]
$current = New-Object System.Collections.Generic.List[object]
$currentBytes = [int64]0
foreach ($d in ($toDelete | Sort-Object { $_.File.FullName })) {
    if ($current.Count -gt 0 -and ($currentBytes + $d.File.Length) -gt $MaxBatchBytes) {
        $batches.Add($current)
        $current = New-Object System.Collections.Generic.List[object]
        $currentBytes = 0
    }
    $current.Add($d)
    $currentBytes += $d.File.Length
}
if ($current.Count -gt 0) { $batches.Add($current) }

Write-Host ''
Write-Host ("{0} Dubletten werden in {1} Paket(en) zu max. {2} GB gelöscht." -f $toDelete.Count, $batches.Count, $MaxBatchGB)
if ($DryRun) {
    Write-Host 'Modus: DRY RUN (nur Simulation).' -ForegroundColor Yellow
} elseif ($NoRecycleBin) {
    Write-Host 'Modus: ENDGÜLTIG löschen (kein Papierkorb).' -ForegroundColor Red
} else {
    Write-Host 'Modus: Papierkorb (Hinweis: auf USB-Sticks/Wechseldatenträgern löscht Windows ggf. endgültig).'
}
if ($SkipRehash) {
    Write-Host 'Hash-Prüfung vor dem Löschen: AUS (-SkipRehash)' -ForegroundColor Yellow
} else {
    Write-Host 'Hash-Prüfung vor dem Löschen: Dublette und behaltene Datei werden neu gehasht und verglichen.'
}

# ---------------------------------------------------------------------------
# Schritt 7/8 - Löschen mit CONTINUE-Abfrage und Fehler-Stopp
# ---------------------------------------------------------------------------

$totalDeleted = 0; $totalSkipped = 0; $totalErrors = 0; $totalBytes = [int64]0
$touchedFolders = New-Object System.Collections.Generic.HashSet[string]([StringComparer]::OrdinalIgnoreCase)
$keptHashCache = @{}
$aborted = $false

for ($b = 0; $b -lt $batches.Count -and -not $aborted; $b++) {
    $batch = $batches[$b]
    $batchBytes = [int64](($batch | ForEach-Object { $_.File.Length } | Measure-Object -Sum).Sum)

    if ($b -gt 0) {
        Write-Host ''
        $cont = Read-Host ("Nächstes Paket {0} von {1} ({2} Dateien, {3}). Zum Fortfahren exakt CONTINUE eintippen" -f ($b + 1), $batches.Count, $batch.Count, (Format-Size $batchBytes))
        if ($cont -cne 'CONTINUE') {
            Write-Host 'Abgebrochen durch Benutzer.' -ForegroundColor Yellow
            $aborted = $true
            break
        }
    }

    Write-Host ''
    Write-Host ("--- Paket {0} von {1}: {2} Dateien, {3} ---" -f ($b + 1), $batches.Count, $batch.Count, (Format-Size $batchBytes)) -ForegroundColor Cyan

    $bDeleted = 0; $bSkipped = 0; $bErrors = 0; $bBytes = [int64]0
    $i = 0
    foreach ($d in $batch) {
        $i++
        $f = $d.File; $k = $d.Keep
        Write-Progress -Activity ("Paket {0}/{1}" -f ($b + 1), $batches.Count) -Status $f.FullName -PercentComplete (($i / $batch.Count) * 100)

        if (-not (Test-Unchanged $k)) {
            $bSkipped++
            Write-Log 'Uebersprungen' $f.FullName $k.FullName $f.Length 'Behaltene Datei fehlt oder wurde seit dem Scan verändert' $f.Hash $k.Hash
            Write-Warning "Übersprungen (behaltene Datei fehlt/verändert): $($f.FullName)"
            continue
        }
        if (-not (Test-Unchanged $f)) {
            $bSkipped++
            Write-Log 'Uebersprungen' $f.FullName $k.FullName $f.Length 'Datei fehlt oder wurde seit dem Scan verändert' $f.Hash $k.Hash
            Write-Warning "Übersprungen (Datei fehlt/verändert): $($f.FullName)"
            continue
        }

        # SHA256 beider Dateien unmittelbar vor dem Löschen neu bilden und vergleichen
        $fHashNow = $f.Hash; $kHashNow = $k.Hash
        if (-not $SkipRehash) {
            try {
                # Behaltene Datei nur einmal pro Lauf neu hashen (kann Gegenstück vieler Dubletten sein)
                if (-not $keptHashCache.ContainsKey($k.FullName)) { $keptHashCache[$k.FullName] = Get-Sha256 $k.FullName }
                $kHashNow = $keptHashCache[$k.FullName]
                $fHashNow = Get-Sha256 $f.FullName
            } catch {
                $bErrors++
                Write-Log 'Fehler' $f.FullName $k.FullName $f.Length ("Hash vor dem Löschen fehlgeschlagen: " + $_.Exception.Message) $f.Hash $k.Hash
                Write-Host "  FEHLER beim Hashen von $($f.FullName): $($_.Exception.Message)" -ForegroundColor Red
                $errAnswer = Read-Host 'Es ist ein Fehler aufgetreten. Trotzdem weitermachen? Dafür exakt CONTINUE eintippen, sonst Abbruch'
                if ($errAnswer -cne 'CONTINUE') { $aborted = $true; break }
                continue
            }
            if ($fHashNow -ne $kHashNow -or $fHashNow -ne $f.Hash) {
                $bSkipped++
                Write-Log 'Uebersprungen' $f.FullName $k.FullName $f.Length 'SHA256 von Dublette und behaltener Datei stimmt vor dem Löschen nicht (mehr) überein' $fHashNow $kHashNow
                Write-Warning "Übersprungen (Hash ungleich): $($f.FullName)"
                continue
            }
        }

        if ($DryRun) {
            $bDeleted++; $bBytes += $f.Length
            Write-Log 'DryRun' $f.FullName $k.FullName $f.Length 'Würde gelöscht werden' $fHashNow $kHashNow
            Write-Host "  [DryRun] $($f.FullName)"
            continue
        }

        try {
            Remove-SingleFile $f.FullName
            $bDeleted++; $bBytes += $f.Length
            [void]$touchedFolders.Add([IO.Path]::GetDirectoryName($f.FullName))
            Write-Log 'Geloescht' $f.FullName $k.FullName $f.Length '' $fHashNow $kHashNow
            Write-Host "  Gelöscht: $($f.FullName)"
        } catch {
            $bErrors++
            Write-Log 'Fehler' $f.FullName $k.FullName $f.Length $_.Exception.Message $fHashNow $kHashNow
            Write-Host "  FEHLER bei $($f.FullName): $($_.Exception.Message)" -ForegroundColor Red
            Write-Progress -Activity ("Paket {0}/{1}" -f ($b + 1), $batches.Count) -Completed
            $errAnswer = Read-Host 'Es ist ein Fehler aufgetreten. Trotzdem weitermachen? Dafür exakt CONTINUE eintippen, sonst Abbruch'
            if ($errAnswer -cne 'CONTINUE') {
                $aborted = $true
                break
            }
        }
    }
    Write-Progress -Activity ("Paket {0}/{1}" -f ($b + 1), $batches.Count) -Completed

    $totalDeleted += $bDeleted; $totalSkipped += $bSkipped; $totalErrors += $bErrors; $totalBytes += $bBytes
    $verb = 'Gelöscht'
    if ($DryRun) { $verb = 'Würde löschen' }
    Write-Host ("Paket {0} fertig: {1} {2} ({3}), übersprungen {4}, Fehler {5}. Gesamt bisher: {6}" -f `
        ($b + 1), $verb, $bDeleted, (Format-Size $bBytes), $bSkipped, $bErrors, (Format-Size $totalBytes))
}

# ---------------------------------------------------------------------------
# Leere Ordner entfernen + Zusammenfassung
# ---------------------------------------------------------------------------

$removedFolders = 0
if (-not $DryRun -and -not $KeepEmptyFolders -and $touchedFolders.Count -gt 0) {
    $removedFolders = Remove-EmptyFolders $touchedFolders $Roots
}

Write-Host ''
Write-Host '=== Zusammenfassung ===' -ForegroundColor Cyan
if ($DryRun) { Write-Host '(DRY RUN - es wurde nichts gelöscht)' -ForegroundColor Yellow }
if ($aborted) { Write-Host 'Vorgang wurde vorzeitig abgebrochen.' -ForegroundColor Yellow }
Write-Host ("Gelöscht:            {0} Dateien ({1})" -f $totalDeleted, (Format-Size $totalBytes))
Write-Host ("Übersprungen:        {0}" -f $totalSkipped)
Write-Host ("Fehler:              {0}" -f $totalErrors)
Write-Host ("Leere Ordner entf.:  {0}" -f $removedFolders)
Write-Host "Report:              $Report"
Write-Host "Lösch-Log:           $($script:DeleteLog)"
