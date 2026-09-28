<#
.SYNOPSIS
    Version 2: Vergleicht zwei Ordner per SHA256-Hash und löscht Dubletten schrittweise aus dem Zielordner.

.DESCRIPTION
    - ReferenceFolder: Vergleichsordner. Aus diesem Ordner wird NIE gelöscht.
    - TargetFolder:    In diesem Ordner wird nach Dubletten gesucht und gelöscht.

    Eine Datei im Zielordner gilt als Dublette, wenn im Vergleichsordner (egal in welchem
    Unterordner, egal unter welchem Namen) eine Datei mit gleicher Größe UND gleichem SHA256-Hash
    existiert. Zeitstempel werden bewusst nicht verglichen, da sie sich beim Kopieren/Hochladen ändern.

    Neu in Version 2:
    - Der SHA256 wird für BEIDE Dateien (Ziel- und Vergleichsdatei) gebildet und explizit verglichen.
      Beide Hashes stehen im Vergleichs-Report und im Lösch-Log (SHA256_Ziel / SHA256_Referenz).
    - Unmittelbar vor jeder Löschung werden beide Dateien NEU gehasht und nochmals verglichen.
      Nur wenn beide frischen Hashes identisch sind, wird gelöscht (abschaltbar mit -SkipRehash).

    Sicherheitsnetz:
    - Doppelte Bestätigung des Ordners, aus dem gelöscht wird
    - Löschen in Paketen von max. -MaxBatchGB (Standard 10 GB), danach "CONTINUE" nötig
    - Vor jeder Löschung: Referenzdatei existiert noch und beide Dateien sind seit dem Scan unverändert
    - Stopp beim ersten Fehler
    - Standardmäßig Papierkorb statt endgültigem Löschen
    - -DryRun: kompletter Ablauf ohne Löschen
    - CSV-Logs von Vergleich und jeder Lösch-Aktion

.EXAMPLE
    .\Compare-And-Remove-Duplicates-v2.ps1 -ReferenceFolder "D:\Original" -TargetFolder "E:\Backup" -DryRun

.EXAMPLE
    .\Compare-And-Remove-Duplicates-v2.ps1 -ReferenceFolder "D:\Original" -TargetFolder "E:\Backup"
#>
[CmdletBinding()]
param(
    [string]$ReferenceFolder,
    [string]$TargetFolder,
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

# ---------------------------------------------------------------------------
# Hilfsfunktionen
# ---------------------------------------------------------------------------

function Format-Size([int64]$Bytes) {
    if ($Bytes -ge 1GB) { return '{0:N2} GB' -f ($Bytes / 1GB) }
    if ($Bytes -ge 1MB) { return '{0:N2} MB' -f ($Bytes / 1MB) }
    if ($Bytes -ge 1KB) { return '{0:N2} KB' -f ($Bytes / 1KB) }
    return "$Bytes B"
}

function Resolve-FolderPath([string]$Path, [string]$Label) {
    while ([string]::IsNullOrWhiteSpace($Path)) {
        $Path = Read-Host "Pfad zum $Label eingeben"
    }
    $Path = $Path.Trim().Trim('"')
    if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
        throw "$Label '$Path' existiert nicht oder ist kein Ordner."
    }
    $full = (Get-Item -LiteralPath $Path).FullName
    return $full.TrimEnd('\', '/') + [IO.Path]::DirectorySeparatorChar
}

function Test-IsInside([string]$Child, [string]$Parent) {
    return $Child.StartsWith($Parent, [StringComparison]::OrdinalIgnoreCase)
}

function Get-FileSnapshot([string]$Root, [string]$Label) {
    Write-Host "Lese Dateien ein: $Label ($Root) ..."
    $files = New-Object System.Collections.Generic.List[object]
    $i = 0
    Get-ChildItem -LiteralPath $Root -Recurse -File -Force -ErrorAction SilentlyContinue -ErrorVariable scanErrors |
        ForEach-Object {
            $i++
            if ($i % 500 -eq 0) { Write-Progress -Activity "Einlesen: $Label" -Status "$i Dateien" }
            $files.Add([pscustomobject]@{
                FullName      = $_.FullName
                RelativePath  = $_.FullName.Substring($Root.Length)
                Name          = $_.Name
                Length        = [int64]$_.Length
                LastWriteUtc  = $_.LastWriteTimeUtc.Ticks
                Hash          = $null
            })
        }
    Write-Progress -Activity "Einlesen: $Label" -Completed
    foreach ($e in $scanErrors) { Write-Warning "Nicht lesbar: $($e.TargetObject) - $($e.Exception.Message)" }
    Write-Host ("  {0} Dateien, {1}" -f $files.Count, (Format-Size (($files | Measure-Object Length -Sum).Sum)))
    return , $files
}

function Add-Hashes($Files, [string]$Label) {
    $total = $Files.Count
    $done = 0
    foreach ($f in $Files) {
        $done++
        Write-Progress -Activity "SHA256 berechnen: $Label" -Status "$done / $total - $($f.Name)" -PercentComplete (($done / [math]::Max($total, 1)) * 100)
        try {
            $f.Hash = Get-Sha256 $f.FullName
        } catch {
            Write-Warning "Hash fehlgeschlagen: $($f.FullName) - $($_.Exception.Message)"
        }
    }
    Write-Progress -Activity "SHA256 berechnen: $Label" -Completed
}

function Get-Sha256([string]$Path) {
    return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash
}

function Test-Unchanged($Snapshot) {
    if (-not (Test-Path -LiteralPath $Snapshot.FullName -PathType Leaf)) { return $false }
    $item = Get-Item -LiteralPath $Snapshot.FullName -Force
    return ($item.Length -eq $Snapshot.Length -and $item.LastWriteTimeUtc.Ticks -eq $Snapshot.LastWriteUtc)
}

function Remove-SingleFile([string]$Path) {
    if ($NoRecycleBin) {
        Remove-Item -LiteralPath $Path -Force
    } else {
        [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile(
            $Path,
            [Microsoft.VisualBasic.FileIO.UIOption]::OnlyErrorDialogs,
            [Microsoft.VisualBasic.FileIO.RecycleOption]::SendToRecycleBin)
    }
}

function Remove-EmptyFolders($Folders, [string]$Root) {
    $removed = 0
    $candidates = New-Object System.Collections.Generic.HashSet[string]([StringComparer]::OrdinalIgnoreCase)
    foreach ($dir in $Folders) {
        # Alle Elternordner bis (exklusive) zum Zielordner-Root als Kandidaten aufnehmen
        $d = $dir
        while ($d -and (Test-IsInside ($d.TrimEnd('\', '/') + [IO.Path]::DirectorySeparatorChar) $Root) -and
               ($d.TrimEnd('\', '/') + [IO.Path]::DirectorySeparatorChar) -ne $Root) {
            [void]$candidates.Add($d)
            $d = [IO.Path]::GetDirectoryName($d)
        }
    }
    # Tiefste Ordner zuerst
    foreach ($dir in ($candidates | Sort-Object Length -Descending)) {
        if (-not (Test-Path -LiteralPath $dir -PathType Container)) { continue }
        if (@(Get-ChildItem -LiteralPath $dir -Force).Count -eq 0) {
            try {
                Remove-Item -LiteralPath $dir -Force
                $removed++
                Write-Log -Action 'OrdnerEntfernt' -Target $dir -Reference '' -Size 0 -Message 'Leerer Ordner'
            } catch {
                Write-Warning "Leerer Ordner konnte nicht entfernt werden: $dir - $($_.Exception.Message)"
            }
        }
    }
    return $removed
}

function Write-Log([string]$Action, [string]$Target, [string]$Reference, [int64]$Size, [string]$Message,
                   [string]$TargetHash = '', [string]$ReferenceHash = '') {
    [pscustomobject]@{
        Zeit            = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        Aktion          = $Action
        Datei           = $Target
        Referenz        = $Reference
        GroesseByte     = $Size
        SHA256_Ziel     = $TargetHash
        SHA256_Referenz = $ReferenceHash
        Info            = $Message
    } | Export-Csv -LiteralPath $script:DeleteLog -Append -NoTypeInformation -Delimiter ';' -Encoding UTF8
}

# ---------------------------------------------------------------------------
# Schritt 1 - Validierung
# ---------------------------------------------------------------------------

Write-Host ''
Write-Host '=== Data-Vergleich v2: Dubletten per SHA256 (beide Dateien) finden und löschen ===' -ForegroundColor Cyan
if ($DryRun) { Write-Host '*** DRY RUN - es wird nichts gelöscht ***' -ForegroundColor Yellow }
Write-Host ''

$ReferenceFolder = Resolve-FolderPath $ReferenceFolder 'Vergleichsordner (wird NICHT verändert)'
$TargetFolder    = Resolve-FolderPath $TargetFolder    'Zielordner (hier wird nach Dubletten gesucht und gelöscht)'

if ($ReferenceFolder -eq $TargetFolder) {
    throw 'Vergleichsordner und Zielordner sind identisch.'
}
if ((Test-IsInside $ReferenceFolder $TargetFolder) -or (Test-IsInside $TargetFolder $ReferenceFolder)) {
    throw 'Die Ordner sind ineinander verschachtelt. Das ist nicht erlaubt, da sonst Dateien mit sich selbst verglichen würden.'
}

if (-not $NoRecycleBin -and -not $DryRun) {
    Add-Type -AssemblyName Microsoft.VisualBasic
}

New-Item -ItemType Directory -Path $LogPath -Force | Out-Null
$stamp            = Get-Date -Format 'yyyyMMdd_HHmmss'
$CompareReport    = Join-Path $LogPath "Vergleich_v2_$stamp.csv"
$script:DeleteLog = Join-Path $LogPath "Loeschung_v2_$stamp.csv"

Write-Host "Vergleichsordner (bleibt unverändert): $ReferenceFolder" -ForegroundColor Green
Write-Host "Zielordner (hier wird gelöscht):       $TargetFolder" -ForegroundColor Yellow
Write-Host ''

# ---------------------------------------------------------------------------
# Schritt 2 - Metadaten einlesen
# ---------------------------------------------------------------------------

$refFiles    = Get-FileSnapshot $ReferenceFolder 'Vergleichsordner'
$targetFiles = Get-FileSnapshot $TargetFolder 'Zielordner'

# ---------------------------------------------------------------------------
# Schritt 3 - Vergleich (Größe als Vorfilter, dann SHA256)
# ---------------------------------------------------------------------------

$refSizes = New-Object System.Collections.Generic.HashSet[int64]
foreach ($f in $refFiles) { [void]$refSizes.Add($f.Length) }

$targetCandidates = @($targetFiles | Where-Object { $refSizes.Contains($_.Length) -and ($IncludeEmptyFiles -or $_.Length -gt 0) })
$candidateSizes = New-Object System.Collections.Generic.HashSet[int64]
foreach ($f in $targetCandidates) { [void]$candidateSizes.Add($f.Length) }
$refCandidates = @($refFiles | Where-Object { $candidateSizes.Contains($_.Length) })

Write-Host ''
Write-Host ("Vorfilter nach Größe: {0} Kandidaten im Zielordner, {1} im Vergleichsordner." -f $targetCandidates.Count, $refCandidates.Count)
Write-Host 'Berechne SHA256-Hashes (kann je nach Datenmenge dauern) ...'

Add-Hashes $refCandidates 'Vergleichsordner'
Add-Hashes $targetCandidates 'Zielordner'

# Vergleichsdateien nach Größe gruppieren. Für jede Zieldatei werden nur Vergleichsdateien gleicher
# Größe herangezogen und die SHA256-Werte BEIDER Dateien explizit miteinander verglichen.
$refBySize = @{}
foreach ($f in $refCandidates) {
    if (-not $f.Hash) { continue }
    if (-not $refBySize.ContainsKey($f.Length)) { $refBySize[$f.Length] = New-Object System.Collections.Generic.List[object] }
    $refBySize[$f.Length].Add($f)
}

$duplicates = New-Object System.Collections.Generic.List[object]
foreach ($t in $targetCandidates) {
    if (-not $t.Hash -or -not $refBySize.ContainsKey($t.Length)) { continue }
    $r = $null
    foreach ($candidate in $refBySize[$t.Length]) {
        if ($candidate.Hash -eq $t.Hash) {
            # Bevorzugt die Vergleichsdatei mit gleichem Namen als Gegenstück nehmen
            if ($null -eq $r -or ($candidate.Name -eq $t.Name -and $r.Name -ne $t.Name)) { $r = $candidate }
        }
    }
    if ($null -ne $r) {
        $duplicates.Add([pscustomobject]@{
            Target        = $t
            Reference     = $r
            SameName      = ($t.Name -eq $r.Name)
            SamePath      = ($t.RelativePath -eq $r.RelativePath)
        })
    }
}

# ---------------------------------------------------------------------------
# Schritt 4 - Report
# ---------------------------------------------------------------------------

$dupBytes = [int64](($duplicates | ForEach-Object { $_.Target.Length } | Measure-Object -Sum).Sum)

$duplicates | ForEach-Object {
    [pscustomobject]@{
        Zieldatei       = $_.Target.FullName
        Referenzdatei   = $_.Reference.FullName
        GroesseZiel     = $_.Target.Length
        GroesseReferenz = $_.Reference.Length
        SHA256_Ziel     = $_.Target.Hash
        SHA256_Referenz = $_.Reference.Hash
        HashGleich      = ($_.Target.Hash -eq $_.Reference.Hash)
        GleicherName    = $_.SameName
        GleicherPfad    = $_.SamePath
    }
} | Export-Csv -LiteralPath $CompareReport -NoTypeInformation -Delimiter ';' -Encoding UTF8

Write-Host ''
Write-Host '=== Ergebnis ===' -ForegroundColor Cyan
Write-Host ("Dateien Vergleichsordner: {0}" -f $refFiles.Count)
Write-Host ("Dateien Zielordner:       {0}" -f $targetFiles.Count)
Write-Host ("Dubletten im Zielordner:  {0} ({1})" -f $duplicates.Count, (Format-Size $dupBytes)) -ForegroundColor Yellow
if ($duplicates.Count -gt 0) {
    Write-Host ("  davon mit anderem Namen:  {0}" -f @($duplicates | Where-Object { -not $_.SameName }).Count)
    Write-Host ("  davon in anderem Ordner:  {0}" -f @($duplicates | Where-Object { -not $_.SamePath }).Count)
}
Write-Host "Vergleichs-Report: $CompareReport"
Write-Host ''

if ($duplicates.Count -eq 0) {
    Write-Host 'Keine Dubletten gefunden. Nichts zu tun.' -ForegroundColor Green
    return
}

$answer = Read-Host 'Report prüfen. Mit dem Löschen fortfahren? (J/N)'
if ($answer -notmatch '^(j|ja|y|yes)$') {
    Write-Host 'Abgebrochen. Es wurde nichts gelöscht.'
    return
}

# ---------------------------------------------------------------------------
# Schritt 5 - Doppelte Bestätigung des Ordners, aus dem gelöscht wird
# ---------------------------------------------------------------------------

Write-Host ''
Write-Host "Gelöscht wird AUSSCHLIESSLICH aus: $TargetFolder" -ForegroundColor Yellow
Write-Host "Der Vergleichsordner bleibt unverändert: $ReferenceFolder" -ForegroundColor Green
$confirm1 = Read-Host "Bestätigung 1/2: Aus dem Zielordner löschen? (J/N)"
if ($confirm1 -notmatch '^(j|ja|y|yes)$') {
    Write-Host 'Abgebrochen. Es wurde nichts gelöscht.'
    return
}

$targetLeaf = Split-Path -Leaf $TargetFolder.TrimEnd('\', '/')
if ([string]::IsNullOrEmpty($targetLeaf)) { $targetLeaf = $TargetFolder.TrimEnd('\', '/') }
$confirm2 = Read-Host "Bestätigung 2/2: Namen des Zielordners zur Bestätigung eintippen ('$targetLeaf')"
if ($confirm2.Trim() -ne $targetLeaf) {
    Write-Host 'Eingabe stimmt nicht mit dem Zielordner überein. Abgebrochen, es wurde nichts gelöscht.' -ForegroundColor Red
    return
}

# ---------------------------------------------------------------------------
# Schritt 6 - Pakete bilden (max. MaxBatchGB)
# ---------------------------------------------------------------------------

$batches = New-Object System.Collections.Generic.List[object]
$current = New-Object System.Collections.Generic.List[object]
$currentBytes = [int64]0
foreach ($d in ($duplicates | Sort-Object { $_.Target.FullName })) {
    if ($current.Count -gt 0 -and ($currentBytes + $d.Target.Length) -gt $MaxBatchBytes) {
        $batches.Add($current)
        $current = New-Object System.Collections.Generic.List[object]
        $currentBytes = 0
    }
    $current.Add($d)
    $currentBytes += $d.Target.Length
}
if ($current.Count -gt 0) { $batches.Add($current) }

Write-Host ''
Write-Host ("{0} Dubletten werden in {1} Paket(en) zu max. {2} GB gelöscht." -f $duplicates.Count, $batches.Count, $MaxBatchGB)
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
    Write-Host 'Hash-Prüfung vor dem Löschen: beide Dateien werden neu gehasht und verglichen.'
}

# ---------------------------------------------------------------------------
# Schritt 7/8 - Löschen mit CONTINUE-Abfrage und Fehler-Stopp
# ---------------------------------------------------------------------------

$totalDeleted = 0; $totalSkipped = 0; $totalErrors = 0; $totalBytes = [int64]0
$touchedFolders = New-Object System.Collections.Generic.HashSet[string]([StringComparer]::OrdinalIgnoreCase)
$aborted = $false

for ($b = 0; $b -lt $batches.Count -and -not $aborted; $b++) {
    $batch = $batches[$b]
    $batchBytes = [int64](($batch | ForEach-Object { $_.Target.Length } | Measure-Object -Sum).Sum)

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
    if (@($batch).Count -eq 1 -and $batchBytes -gt $MaxBatchBytes) {
        Write-Host ("Hinweis: Einzelne Datei größer als {0} GB: {1}" -f $MaxBatchGB, $batch[0].Target.FullName) -ForegroundColor Yellow
    }

    $bDeleted = 0; $bSkipped = 0; $bErrors = 0; $bBytes = [int64]0
    $i = 0
    foreach ($d in $batch) {
        $i++
        $t = $d.Target; $r = $d.Reference
        Write-Progress -Activity ("Paket {0}/{1}" -f ($b + 1), $batches.Count) -Status $t.RelativePath -PercentComplete (($i / $batch.Count) * 100)

        # Sicherheitsprüfungen vor jeder Löschung
        if (-not (Test-Unchanged $r)) {
            $bSkipped++
            Write-Log 'Uebersprungen' $t.FullName $r.FullName $t.Length 'Referenzdatei fehlt oder wurde seit dem Scan verändert' $t.Hash $r.Hash
            Write-Warning "Übersprungen (Referenz fehlt/verändert): $($t.FullName)"
            continue
        }
        if (-not (Test-Unchanged $t)) {
            $bSkipped++
            Write-Log 'Uebersprungen' $t.FullName $r.FullName $t.Length 'Zieldatei fehlt oder wurde seit dem Scan verändert' $t.Hash $r.Hash
            Write-Warning "Übersprungen (Zieldatei fehlt/verändert): $($t.FullName)"
            continue
        }

        # SHA256 beider Dateien unmittelbar vor dem Löschen neu bilden und vergleichen
        $tHashNow = $t.Hash; $rHashNow = $r.Hash
        if (-not $SkipRehash) {
            try {
                $rHashNow = Get-Sha256 $r.FullName
                $tHashNow = Get-Sha256 $t.FullName
            } catch {
                $bErrors++
                Write-Log 'Fehler' $t.FullName $r.FullName $t.Length ("Hash vor dem Löschen fehlgeschlagen: " + $_.Exception.Message) $t.Hash $r.Hash
                Write-Host "  FEHLER beim Hashen von $($t.FullName): $($_.Exception.Message)" -ForegroundColor Red
                $errAnswer = Read-Host 'Es ist ein Fehler aufgetreten. Trotzdem weitermachen? Dafür exakt CONTINUE eintippen, sonst Abbruch'
                if ($errAnswer -cne 'CONTINUE') { $aborted = $true; break }
                continue
            }
            if ($tHashNow -ne $rHashNow -or $tHashNow -ne $t.Hash) {
                $bSkipped++
                Write-Log 'Uebersprungen' $t.FullName $r.FullName $t.Length 'SHA256 von Ziel- und Vergleichsdatei stimmt vor dem Löschen nicht (mehr) überein' $tHashNow $rHashNow
                Write-Warning "Übersprungen (Hash ungleich): $($t.FullName)"
                continue
            }
        }

        if ($DryRun) {
            $bDeleted++; $bBytes += $t.Length
            Write-Log 'DryRun' $t.FullName $r.FullName $t.Length 'Würde gelöscht werden' $tHashNow $rHashNow
            Write-Host "  [DryRun] $($t.RelativePath)"
            continue
        }

        try {
            Remove-SingleFile $t.FullName
            $bDeleted++; $bBytes += $t.Length
            [void]$touchedFolders.Add([IO.Path]::GetDirectoryName($t.FullName))
            Write-Log 'Geloescht' $t.FullName $r.FullName $t.Length '' $tHashNow $rHashNow
            Write-Host "  Gelöscht: $($t.RelativePath)"
        } catch {
            $bErrors++
            Write-Log 'Fehler' $t.FullName $r.FullName $t.Length $_.Exception.Message $tHashNow $rHashNow
            Write-Host "  FEHLER bei $($t.FullName): $($_.Exception.Message)" -ForegroundColor Red
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
    $removedFolders = Remove-EmptyFolders $touchedFolders $TargetFolder
}

Write-Host ''
Write-Host '=== Zusammenfassung ===' -ForegroundColor Cyan
if ($DryRun) { Write-Host '(DRY RUN - es wurde nichts gelöscht)' -ForegroundColor Yellow }
if ($aborted) { Write-Host 'Vorgang wurde vorzeitig abgebrochen.' -ForegroundColor Yellow }
Write-Host ("Gelöscht:            {0} Dateien ({1})" -f $totalDeleted, (Format-Size $totalBytes))
Write-Host ("Übersprungen:        {0}" -f $totalSkipped)
Write-Host ("Fehler:              {0}" -f $totalErrors)
Write-Host ("Leere Ordner entf.:  {0}" -f $removedFolders)
Write-Host "Vergleichs-Report:   $CompareReport"
Write-Host "Lösch-Log:           $($script:DeleteLog)"
