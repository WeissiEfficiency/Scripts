<#.SYNOPSIS
.DESCRIPTION
- Liest UserPrincipalName und Country aus der angegebenen CSV-Datei.
- Verwendet das AzureAD-Modul (erfordert vorheriges Connect-AzureAD).
- Exportiert Entra-Benutzer- und Manager-Daten als XML in das übergeordnete Verzeichnis der CSV.
- Findet den lokalen AD-Benutzer ausschließlich über den UserPrincipalName.
- Fordert nur die minimal notwendigen AD-Attribute (objectGUID, SamAccountName) an.
- Berechnet die ImmutableID für Entra ID basierend auf der AD objectGUID.
- Bereitet (auskommentierte) Updates für AD-Attribute vor:
    - Land (countryCode) wird aus der CSV gelesen und konvertiert (Netherlands->NL, United Kingdom->GB).
    - Telefon, Adresse, Titel, Abteilung, Ort, Firma werden aus dem Entra-Benutzerobjekt übernommen (sofern vorhanden).
    - extensionAttribute10 wird auf "365" gesetzt.
- Fragt am Ende optional die Auslösung eines Delta Syncs ab (Befehl auskommentiert).
.PARAMETER CsvPfad
Der Pfad zur CSV-Datei. Diese **muss** die Spalten 'UserPrincipalName' und 'Country' enthalten. Die XML-Exportdateien werden im übergeordneten Verzeichnis dieses Pfads gespeichert. Standard ist '.\user_upns.csv'.
.EXAMPLE
.\EntraAdSyncRefactored.ps1 -CsvPfad "C:\Skripte\Benutzerlisten\user_data.csv"
# Stellt sicher, dass C:\Skripte\Benutzerlisten existiert!

.EXAMPLE
.\EntraAdSyncRefactored.ps1 -CsvPfad ".\BenutzerMitLand.csv" -Verbose -WhatIf
# Stellt sicher, dass das aktuelle Verzeichnis existiert.

.NOTES
VERSION: 1.0
CSV-ANFORDERUNG: Die Datei muss die Spalten 'UserPrincipalName' und 'Country' enthalten.
VERZEICHNIS-ANFORDERUNG: Das übergeordnete Verzeichnis des CsvPfad MUSS existieren, bevor das Skript gestartet wird!
MODUL-ANFORDERUNG: AzureAD, ActiveDirectory (RSAT). Führen Sie 'Connect-AzureAD' vor dem Start aus.
BERECHTIGUNGEN: Stellen Sie sicher, dass das ausführende Konto Lese-/Schreibrechte in Entra ID und AD sowie Schreibrechte im XML-Exportverzeichnis hat.
#>
[CmdletBinding(SupportsShouldProcess=$true)] # Ermöglicht -WhatIf und -Confirm für das GESAMTE Skript
param(
    [Parameter(Mandatory=$false)]
    [string]$CsvPfad = ".\user_upns.csv" # Standardwert, Datei muss UPN und Country enthalten
)

#region Hilfsfunktionen

# Stellt sicher, dass die notwendigen Module vorhanden und verbunden sind
function Test-Prerequisites {
    Write-Verbose "Prüfe Voraussetzungen..."
    if (-not (Get-Module -ListAvailable -Name AzureAD)) {
        Write-Error "Das AzureAD-Modul ist nicht installiert. Bitte installieren Sie es mit 'Install-Module -Name AzureAD'."
        return $false
    }
    try {
        Get-AzureADCurrentUser -ErrorAction Stop | Out-Null
        Write-Verbose "Verbindung zu AzureAD besteht."
    } catch {
        Write-Error "Es besteht keine Verbindung zu AzureAD. Bitte führen Sie 'Connect-AzureAD' aus."
        return $false
    }
    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        Write-Warning "Das ActiveDirectory-Modul ist nicht verfügbar. AD-Operationen werden fehlschlagen. Installieren Sie die RSAT-AD-Tools."
    }
    Write-Verbose "Voraussetzungen erfolgreich geprüft."
    return $true
}

# Liest Benutzerdaten (als Objekte) aus einer CSV-Datei (erwartet 'UserPrincipalName', 'Country')
function Get-UserDataFromCsv {
    param(
        [Parameter(Mandatory=$true)]
        [string]$FilePath
    )
    try {
        Write-Verbose "Lese Benutzerdaten (UPN, Country) aus '$FilePath'."
        $userData = Import-Csv -Path $FilePath -ErrorAction Stop
        if ($userData -and ($userData[0].PSObject.Properties.Name -contains 'UserPrincipalName') -and ($userData[0].PSObject.Properties.Name -contains 'Country')) {
             Write-Verbose "$($userData.Count) Datensätze aus CSV gelesen."
             return $userData
        } elseif (-not $userData) {
             Write-Warning "Die CSV-Datei '$FilePath' ist leer oder konnte nicht gelesen werden."
             return $null
        } else {
            Write-Error "Die CSV-Datei '$FilePath' enthält nicht die erwarteten Spalten 'UserPrincipalName' UND 'Country'."
            return $null
        }
    } catch {
        Write-Error "Fehler beim Lesen der CSV-Datei '$FilePath': $($_.Exception.Message)"
        return $null
    }
}

# Ruft einen Entra ID Benutzer anhand seines UPN ab
function Get-EntraUserByUpn {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Upn
    )
    Write-Verbose "Suche Entra ID Benutzer für UPN: $Upn"
    try {
        $user = Get-AzureADUser -Filter "UserPrincipalName eq '$Upn'" -ErrorAction Stop
        if ($user) {
            Write-Verbose "Entra ID Benutzer gefunden: $($user.ObjectId)"
            return $user
        } else {
            Write-Warning "Benutzer mit UPN '$Upn' nicht in Entra ID gefunden."
            return $null
        }
    } catch {
        Write-Error "Fehler beim Abrufen des Entra ID Benutzers für UPN '$Upn': $($_.Exception.Message)"
        return $null
    }
}

# Exportiert ein Objekt in eine XML-Datei (CLIXML)
function Export-ObjectToXml {
    param(
        [Parameter(Mandatory=$true)] $InputObject,
        [Parameter(Mandatory=$true)] [string]$BasePath, # Muss existierendes Verzeichnis sein!
        [Parameter(Mandatory=$true)] [string]$FileName
    )
    $FullXmlPath = Join-Path -Path $BasePath -ChildPath $FileName
    Write-Verbose "Exportiere Objekt nach '$FullXmlPath'..."
    try {
        # Export-Clixml schlägt fehl, wenn BasePath nicht existiert!
        $InputObject | Export-Clixml -Path $FullXmlPath -ErrorAction Stop
        Write-Host "Daten wurden nach '$FullXmlPath' exportiert."
    } catch {
        # Fehler abfangen und spezifischere Meldung ausgeben, wenn möglich
        if ($_.Exception.Message -like "*Could not find a part of the path*") {
             Write-Error "Fehler beim Exportieren nach '$FullXmlPath': Das Verzeichnis '$BasePath' wurde nicht gefunden. $($_.Exception.Message)"
        } else {
             Write-Error "Fehler beim Exportieren nach '$FullXmlPath': $($_.Exception.Message)"
        }

    }
}

# Ruft die Managerdaten eines Entra ID Benutzers ab und formatiert sie
function Get-EntraUserManagerData {
    param(
        [Parameter(Mandatory=$true)] $EntraUser # Objekt vom Typ [Microsoft.Open.MSGraph.Model.User]
    )
    Write-Verbose "Suche Manager für Entra ID Benutzer: $($EntraUser.UserPrincipalName)"
    try {
        $manager = Get-AzureADUserManager -ObjectId $EntraUser.ObjectId -ErrorAction SilentlyContinue
        if ($manager) {
            Write-Verbose "Manager gefunden: $($manager.UserPrincipalName)"
            return [PSCustomObject]@{
                ObjectId          = $manager.ObjectId
                UserPrincipalName = $manager.UserPrincipalName
            }
        } else {
            Write-Warning "Für Benutzer '$($EntraUser.UserPrincipalName)' wurde kein Manager in Entra ID gefunden."
            return $null
        }
    } catch {
         Write-Error "Fehler beim Abrufen des Managers für '$($EntraUser.UserPrincipalName)': $($_.Exception.Message)"
        return $null
    }
}

# Ruft einen lokalen AD Benutzer anhand seines UPN ab (optimierte Properties)
function Get-LocalAdUserByUpn {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Upn
    )
    Write-Verbose "Suche lokalen AD Benutzer für UPN: $Upn"
    try {
        $adUser = Get-ADUser -Filter "UserPrincipalName -eq '$Upn'" -Properties objectGUID, SamAccountName -ErrorAction Stop
        if ($adUser) {
            Write-Verbose "Lokalen AD Benutzer gefunden: $($adUser.SamAccountName)"
            return $adUser
        } else {
            Write-Warning "Kein passender lokaler AD-Benutzer für UPN '$Upn' gefunden."
            return $null
        }
    } catch {
        Write-Error "Fehler beim Abrufen des lokalen AD Benutzers für UPN '$Upn': $($_.Exception.Message)"
        return $null
    }
}

# Berechnet die ImmutableID aus der ObjectGUID eines AD Benutzers
function Get-ImmutableIdFromAdUser {
    param(
        [Parameter(Mandatory=$true)] $AdUser # Objekt vom Typ [Microsoft.ActiveDirectory.Management.ADUser]
    )
    if (-not $AdUser.objectGUID) {
         Write-Error "objectGUID nicht im ADUser-Objekt für '$($AdUser.SamAccountName)' gefunden."
         return $null
    }
    try {
        $immutableIdBytes = [System.Guid]::Parse($AdUser.objectGUID).ToByteArray()
        $immutableId = [System.Convert]::ToBase64String($immutableIdBytes)
        Write-Verbose "Berechnete ImmutableID für '$($AdUser.SamAccountName)': $immutableId"
        return $immutableId
    } catch {
        Write-Error "Fehler beim Berechnen der ImmutableID für '$($AdUser.SamAccountName)': $($_.Exception.Message)"
        return $null
    }
}

# (Vorbereitet) Aktualisiert Attribute eines lokalen AD Benutzers
function Update-LocalAdUserAttributes {
    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
        [Parameter(Mandatory=$true)] $AdUser,
        [Parameter(Mandatory=$true)] [string]$CountryFromCsv,
        [Parameter(Mandatory=$true)] $EntraUser
    )
    Write-Verbose "Bereite Attribut-Updates für AD Benutzer '$($AdUser.SamAccountName)' vor..."

    # ---- Land aus CSV konvertieren ----
    $Land = $null
    if (-not [string]::IsNullOrWhiteSpace($CountryFromCsv)) {
        Write-Verbose "Konvertiere Länderangabe aus CSV: '$CountryFromCsv'"
        switch -Wildcard ($CountryFromCsv) {
             { $_ -ieq "Netherlands" }    { $Land = "NL"; Write-Verbose "Länderkennzeichen zu '$Land' konvertiert."; break }
             { $_ -ieq "United Kingdom" } { $Land = "GB"; Write-Verbose "Länderkennzeichen zu '$Land' konvertiert."; break }
             default { $Land = $CountryFromCsv; Write-Verbose "Keine Konvertierung, verwende '$Land' aus CSV." }
        }
    } else {
         Write-Warning "Keine Länderangabe in CSV für '$($AdUser.SamAccountName)' gefunden. Land wird nicht gesetzt."
    }

    # ---- Attribute aus EntraUser vorbereiten ----
    $setParams = @{ Identity = $AdUser; Replace = @{extensionAttribute10 = "365"} }
    if (-not [string]::IsNullOrWhiteSpace($EntraUser.TelephoneNumber)) { $setParams.Add('OfficePhone', $EntraUser.TelephoneNumber) }
    if (-not [string]::IsNullOrWhiteSpace($EntraUser.StreetAddress)) { $setParams.Add('StreetAddress', $EntraUser.StreetAddress) }
    if (-not [string]::IsNullOrWhiteSpace($EntraUser.PostalCode)) { $setParams.Add('PostalCode', $EntraUser.PostalCode) }
    if (-not [string]::IsNullOrWhiteSpace($EntraUser.JobTitle)) { $setParams.Add('Title', $EntraUser.JobTitle) }
    if (-not [string]::IsNullOrWhiteSpace($EntraUser.Department)) { $setParams.Add('Department', $EntraUser.Department) }
    if (-not [string]::IsNullOrWhiteSpace($EntraUser.City)) { $setParams.Add('City', $EntraUser.City) }
    if (-not [string]::IsNullOrWhiteSpace($EntraUser.CompanyName)) { $setParams.Add('Company', $EntraUser.CompanyName) }
    if ($Land) { $setParams.Add('countryCode', $Land) } # countryCode (Attribut c)

    # ---- Beschreibung für ShouldProcess ----
    $attrUpdates = [System.Collections.Generic.List[string]]::new()
    if ($setParams.ContainsKey('OfficePhone')) { $attrUpdates.Add("OfficePhone") }
    # ... (restliche Attribute wie in V6) ...
    if ($setParams.ContainsKey('Company')) { $attrUpdates.Add("Company") }
    $attrUpdates.Add("extensionAttribute10='365'")
    if ($setParams.ContainsKey('countryCode')) { $attrUpdates.Add("countryCode='$($setParams.countryCode)'") }
    if($attrUpdates.Count -gt 0) { $updateActionDescription = "Attribute aktualisieren ($($attrUpdates -join ', '))" }
    else { $updateActionDescription = "Attribute aktualisieren (nur via -Replace)" }

    # ---- Ausführung (Vorbereitet) ----
    if ($PSCmdlet.ShouldProcess($AdUser.DistinguishedName, $updateActionDescription)) {
        Write-Host "[Vorbereitet] Führe Set-ADUser für '$($AdUser.SamAccountName)' aus..."
        Set-ADUser @setParams # <-- HIER EINKOMMENTIEREN ZUM AKTIVIEREN
    }
}

# (Vorbereitet) Setzt den Manager für einen lokalen AD Benutzer
function Set-LocalAdUserManager {
    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
        [Parameter(Mandatory=$true)] $AdUser,
        [Parameter(Mandatory=$true)] $EntraManagerData
    )
    Write-Verbose "Versuche Manager für AD Benutzer '$($AdUser.SamAccountName)' zu setzen..."
    $ManagerSamAccountName = $EntraManagerData.UserPrincipalName.Split('@')[0] # Annahme!
    Write-Verbose "Suche lokalen AD Manager mit SAMAccountName: $ManagerSamAccountName"
    try {
        $AdManager = Get-ADUser -Filter "SamAccountName -eq '$ManagerSamAccountName'" -ErrorAction Stop
        if ($AdManager) {
             if ($PSCmdlet.ShouldProcess($AdUser.DistinguishedName, "Manager setzen auf '$($AdManager.Name)'")) {
                 Write-Host "[Vorbereitet] Setze Manager für '$($AdUser.SamAccountName)' auf '$($AdManager.SamAccountName)'..."
                 Set-ADUser -Identity $AdUser -Manager $AdManager # <-- HIER EINKOMMENTIEREN ZUM AKTIVIEREN
             }
        } else {
            Write-Warning "Der Manager mit dem SAMAccountName '$ManagerSamAccountName' (abgeleitet von '$($EntraManagerData.UserPrincipalName)') wurde im lokalen AD nicht gefunden."
        }
    } catch {
         Write-Error "Fehler beim Suchen des lokalen AD Managers '$ManagerSamAccountName': $($_.Exception.Message)"
    }
}

# (Vorbereitet) Setzt die ImmutableID für einen Entra ID Benutzer
function Set-EntraUserImmutableId {
    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
        [Parameter(Mandatory=$true)] $EntraUser,
        [Parameter(Mandatory=$true)] [string]$ImmutableId
    )
    Write-Verbose "Bereite Setzen der ImmutableID für Entra Benutzer '$($EntraUser.UserPrincipalName)' vor..."
    if ($PSCmdlet.ShouldProcess($EntraUser.ObjectId, "ImmutableID setzen auf '$ImmutableId'")) {
        Write-Host "[Vorbereitet] Setze ImmutableID für '$($EntraUser.UserPrincipalName)'..."
        Set-AzureADUser -ObjectId $EntraUser.ObjectId -ImmutableId $ImmutableId # <-- HIER EINKOMMENTIEREN ZUM AKTIVIEREN
    }
}

# Fragt den Benutzer und löst (vorbereitet) einen Delta Sync aus
function Invoke-EntraDeltaSync {
    [CmdletBinding(SupportsShouldProcess=$true)]
    param()
    $Antwort = Read-Host "Soll ein Entra ID Delta Sync ausgelöst werden? (ja/nein)"
    if ($Antwort -ieq 'ja') {
        if ($PSCmdlet.ShouldProcess("Entra ID / Azure AD Connect", "Delta Sync starten")) {
             Write-Host "[Vorbereitet] Entra ID Delta Sync wird ausgelöst (Befehl muss ggf. angepasst werden)..."
             Start-ADSyncSyncCycle -PolicyType Delta # <-- HIER EINKOMMENTIEREN UND ANPASSEN ZUM AKTIVIEREN
        }
    } else {
        Write-Host "Kein Entra ID Delta Sync ausgelöst."
    }
}

#endregion Hilfsfunktionen

#region Hauptskript

# --- Initialisierung ---
Write-Host "Starte Entra ID / AD Sync Skript (Version 7)..."
if (-not (Test-Prerequisites)) { exit 1 }

Write-Verbose "Prüfe CSV-Pfad: $CsvPfad"
if (-not (Test-Path -Path $CsvPfad -PathType Leaf)) {
    Write-Error "Die CSV-Datei '$CsvPfad' wurde nicht gefunden."
    exit 1
}

# --- Pfade ermitteln ---
$AbsoluteCsvPath = $null
$XmlExportPath = $null
try {
     $AbsoluteCsvPath = Resolve-Path -Path $CsvPfad -ErrorAction Stop
     $XmlExportPath = Split-Path -Path $AbsoluteCsvPath -Parent -ErrorAction Stop
     Write-Host "XML-Exportpfad: '$XmlExportPath' (Dieses Verzeichnis muss existieren!)"
} catch {
     Write-Error "Fehler beim Ermitteln der Pfade für '$CsvPfad': $($_.Exception.Message)"
     exit 1
}

# --- Daten einlesen ---
$UserEntriesFromCsv = Get-UserDataFromCsv -FilePath $AbsoluteCsvPath
if ($null -eq $UserEntriesFromCsv) { exit 1 }

# --- Hauptverarbeitung ---
Write-Host "Starte die Verarbeitung von $($UserEntriesFromCsv.Count) Benutzereinträgen aus '$AbsoluteCsvPath'..."
foreach ($UserData in $UserEntriesFromCsv) {
    $CurrentUpn = $UserData.UserPrincipalName
    $CurrentCountry = $UserData.Country

    if ([string]::IsNullOrWhiteSpace($CurrentUpn)) {
        Write-Warning "Überspringe Eintrag in CSV, da UserPrincipalName fehlt oder leer ist."
        continue
    }

    Write-Host "--------------------------------------------------"
    Write-Host "Verarbeite: $CurrentUpn (Land CSV: '$CurrentCountry')"

    $EntraUser = Get-EntraUserByUpn -Upn $CurrentUpn
    if (-not $EntraUser) { continue }

    # XML Export
    $safeUpnPart = $EntraUser.UserPrincipalName -replace '[^a-zA-Z0-9._-]','_'
    Export-ObjectToXml -InputObject $EntraUser -BasePath $XmlExportPath -FileName "$($safeUpnPart).user.xml"
    $EntraManagerData = Get-EntraUserManagerData -EntraUser $EntraUser
    if ($EntraManagerData) {
        Export-ObjectToXml -InputObject $EntraManagerData -BasePath $XmlExportPath -FileName "$($safeUpnPart).manager.xml"
    }

    # AD Operationen vorbereiten
    $AdUser = Get-LocalAdUserByUpn -Upn $CurrentUpn
    if ($AdUser) {
        $immutableId = Get-ImmutableIdFromAdUser -AdUser $AdUser
        if ($immutableId) {
            Set-EntraUserImmutableId -EntraUser $EntraUser -ImmutableId $immutableId
        }
        Update-LocalAdUserAttributes -AdUser $AdUser -CountryFromCsv $CurrentCountry -EntraUser $EntraUser
        if ($EntraManagerData) {
            Set-LocalAdUserManager -AdUser $AdUser -EntraManagerData $EntraManagerData
        } else {
            Write-Host "Keine Manager-Daten aus Entra ID vorhanden, lokaler AD-Manager wird nicht gesetzt/geprüft."
        }
    } else {
        Write-Warning "Keine lokalen AD-Operationen für '$CurrentUpn' möglich, da AD-Benutzer nicht gefunden."
    }
} # Ende foreach UserData

# --- Abschluss ---
Write-Host "--------------------------------------------------"
Write-Host "$($UserEntriesFromCsv.Count) Benutzereinträge wurden verarbeitet."

Invoke-EntraDeltaSync

Write-Host "`nSkriptausführung abgeschlossen."
#endregion Hauptskript
