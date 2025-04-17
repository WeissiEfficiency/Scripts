<#
.SYNOPSIS
    Synchronisiert ausgewählte Attribute von Azure AD (Entra ID) Benutzern zu lokalen AD Benutzern basierend auf dem UserPrincipalName.

.DESCRIPTION
    Dieses Skript liest eine CSV-Datei mit UserPrincipalNames und Länderinformationen.
    Für jeden Benutzer werden Informationen aus Azure AD abgerufen und entsprechende Attribute
    (Telefon, Adresse, Titel, Abteilung, Firma, Land, Manager) im lokalen Active Directory aktualisiert.
    Zusätzlich wird die ImmutableID (als msDS-ConsistencyGuid im lokalen AD) basierend auf der ObjectGUID des lokalen Benutzers gesetzt.

.PARAMETER CsvPath
    Pfad zur CSV-Datei. Die CSV muss mindestens die Spalten 'UserPrincipalName' und 'Country' enthalten.

.EXAMPLE
    .\Sync-AzureADToLocalAD_Improved.ps1 -CsvPath "C:\Temp\users.csv"

.NOTES
    Voraussetzungen:
    - Das PowerShell-Modul 'AzureAD' muss installiert sein (Install-Module AzureAD). Beachten Sie, dass dieses Modul veraltet ist.
    - RSAT ADDS PowerShell Tools müssen installiert sein (ActiveDirectory Modul).
    - Eine Verbindung zu Azure AD muss vor Ausführung des Skripts mit Connect-AzureAD hergestellt werden.
    - Ausreichende Berechtigungen in Azure AD (Lesen von Benutzerdaten und Manager) und im lokalen AD (Schreiben von Benutzerattributen, inkl. msDS-ConsistencyGuid) sind erforderlich.
    - Das Attribut 'msDS-ConsistencyGuid' muss im lokalen AD-Schema vorhanden sein.

.AUTHOR
    Verbessert basierend auf Vorlage
    Datum: 17. April 2025
#>

# Sicherstellen, dass das ActiveDirectory-Modul verfügbar ist
if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
    Write-Error "Das ActiveDirectory PowerShell Modul wurde nicht gefunden. Bitte installieren Sie die RSAT ADDS Tools."
    exit
}
Import-Module ActiveDirectory -ErrorAction Stop

param(
    [Parameter(Mandatory=$true)]
    [ValidateScript({Test-Path $_ -PathType Leaf})]
    [string]$CsvPath
)

#Region Functions

# Funktion zum Abrufen der AzureAD-Benutzerinformationen
function Get-EntraUserInfo {
    param(
        [Parameter(Mandatory=$true)]
        [string]$UserPrincipalName
    )
    try {
        # Benötigte Attribute auswählen für Effizienz
        $azureADUser = Get-AzureADUser -Filter "userPrincipalName eq '$UserPrincipalName'" | Select-Object ObjectId, UserPrincipalName, DisplayName, TelephoneNumber, StreetAddress, PostalCode, City, JobTitle, Department, CompanyName
        if ($azureADUser) {
            Write-Verbose "Azure AD Benutzer '$UserPrincipalName' gefunden (ObjectId: $($azureADUser.ObjectId))."
            return $azureADUser
        } else {
            Write-Warning "Benutzer '$UserPrincipalName' wurde nicht in Azure AD gefunden."
            return $null
        }
    } catch {
        Write-Error "Fehler beim Abrufen des Azure AD Benutzers '$UserPrincipalName': $_"
        return $null
    }
}

# Funktion zum Abrufen des Azure AD Managers
function Get-EntraUserManagerInfo {
    param(
        [Parameter(Mandatory=$true)]
        $AzureADUserObjectId
    )
    try {
        $manager = Get-AzureADUserManager -ObjectId $AzureADUserObjectId | Select-Object ObjectId, UserPrincipalName, DisplayName
        if ($manager) {
            Write-Verbose "Manager für $($AzureADUserObjectId) gefunden: $($manager.UserPrincipalName)."
            return $manager
        } else {
            Write-Verbose "Kein Manager für Azure AD Benutzer $($AzureADUserObjectId) gefunden."
            return $null
        }
    } catch {
        # Fehler abfangen, wenn kein Manager gefunden wird (ist kein echter Fehler in diesem Kontext)
        if ($_.Exception.Message -like "*Resource*'$($AzureADUserObjectId'*does not exist or one of its queried reference-property objects are not present*") {
             Write-Verbose "Kein Manager für Azure AD Benutzer $($AzureADUserObjectId) gefunden (API-Antwort)."
             return $null
        } else {
            Write-Warning "Fehler beim Abrufen des Managers für Azure AD Benutzer $($AzureADUserObjectId): $_"
            return $null
        }
    }
}

# Funktion zum optionalen Exportieren der Benutzer- und Manager-Informationen in XML-Dateien (für Backup/Logging)
function Export-UserAndManagerInfoToXml {
    param(
        [Parameter(Mandatory=$true)]
        $AzureADUser,
        [Parameter(Mandatory=$false)]
        $AzureADManager,
        [Parameter(Mandatory=$true)]
        [string]$OutputPath
    )
    try {
        # Sicherstellen, dass der Output-Pfad existiert
        if (-not (Test-Path -Path $OutputPath -PathType Container)) {
            New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
        }

        # Bereinigen des UPN für den Dateinamen (ersetzt ungültige Zeichen)
        $safeFileName = $AzureADUser.UserPrincipalName -replace '[\\/:"*?<>|]', '_'

        # Exportieren der Benutzerinformationen
        $xmlUserPath = Join-Path -Path $OutputPath -ChildPath "$($safeFileName).user.xml"
        $AzureADUser | Export-Clixml -Path $xmlUserPath
        Write-Verbose "Azure AD Benutzerdaten exportiert nach '$xmlUserPath'."

        # Exportieren der Manager-Informationen, falls vorhanden
        if ($AzureADManager) {
            $xmlManagerPath = Join-Path -Path $OutputPath -ChildPath "$($safeFileName).manager.xml"
            $AzureADManager | Export-Clixml -Path $xmlManagerPath
             Write-Verbose "Azure AD Managerdaten exportiert nach '$xmlManagerPath'."
        }
    } catch {
        Write-Warning "Fehler beim Exportieren der XML-Dateien für $($AzureADUser.UserPrincipalName): $_"
    }
}

# Funktion zum Aktualisieren des lokalen AD-Benutzers mit AzureAD-Attributen und Setzen der ImmutableID
function Set-LocalADUserAttributes {
    param(
        [Parameter(Mandatory=$true)]
        $AzureADUser,
        [Parameter(Mandatory=$false)]
        $AzureADManager,
        [Parameter(Mandatory=$true)]
        [string]$Country # Aus der CSV
    )
    $localADUser = $null # Sicherstellen, dass die Variable definiert ist
    try {
        # Benötigte Attribute für Get-ADUser definieren
        $adProperties = @(
            'ObjectGUID',
            'DistinguishedName',
            'SamAccountName',
            'UserPrincipalName',
            'OfficePhone',
            'StreetAddress',
            'PostalCode',
            'l', # City
            'Title',
            'Department',
            'Company',
            'co', # Country Name
            'c',  # Country Code (2-letter)
            'Manager',
            'msDS-ConsistencyGuid'
        )

        # Lokalen AD-Benutzer anhand des UPN finden
        Write-Verbose "Suche lokalen AD Benutzer mit UPN '$($AzureADUser.UserPrincipalName)'..."
        $localADUser = Get-ADUser -Filter "UserPrincipalName -eq '$($AzureADUser.UserPrincipalName)'" -Properties $adProperties

        if ($localADUser) {
            Write-Host "  Lokaler AD Benutzer '$($localADUser.SamAccountName)' gefunden." -ForegroundColor Green

            # --- ImmutableID vorbereiten ---
            # Konvertiere die ObjectGUID des *lokalen* Benutzers in den Base64-String für die ImmutableID
            $immutableID = [System.Convert]::ToBase64String($localADUser.ObjectGUID.ToByteArray())
            Write-Verbose "  Berechnete ImmutableID (msDS-ConsistencyGuid): $immutableID"

            # --- Hashtable für zu aktualisierende AD-Attribute erstellen ---
            $UpdateAttributes = @{}

            # Attribute aus AzureAD übernehmen, wenn sie vorhanden sind und sich vom lokalen Wert unterscheiden (optional, aber gute Praxis)
            if ($AzureADUser.TelephoneNumber -and $AzureADUser.TelephoneNumber -ne $localADUser.OfficePhone) { $UpdateAttributes["OfficePhone"] = $AzureADUser.TelephoneNumber }
            if ($AzureADUser.StreetAddress -and $AzureADUser.StreetAddress -ne $localADUser.StreetAddress) { $UpdateAttributes["StreetAddress"] = $AzureADUser.StreetAddress }
            if ($AzureADUser.PostalCode -and $AzureADUser.PostalCode -ne $localADUser.PostalCode) { $UpdateAttributes["PostalCode"] = $AzureADUser.PostalCode }
            if ($AzureADUser.JobTitle -and $AzureADUser.JobTitle -ne $localADUser.Title) { $UpdateAttributes["Title"] = $AzureADUser.JobTitle }
            if ($AzureADUser.Department -and $AzureADUser.Department -ne $localADUser.Department) { $UpdateAttributes["Department"] = $AzureADUser.Department }
            if ($AzureADUser.City -and $AzureADUser.City -ne $localADUser.l) { $UpdateAttributes["l"] = $AzureADUser.City } # 'l' ist das AD-Attribut für City
            if ($AzureADUser.CompanyName -and $AzureADUser.CompanyName -ne $localADUser.Company) { $UpdateAttributes["Company"] = $AzureADUser.CompanyName }

            # Ländercode konvertieren und zum Hashtable hinzufügen
            # Beachte: AD hat oft 'co' (Country Name) und 'c' (2-Letter Code)
            $countryCode = $null
            $countryName = $Country # Standardmäßig den Namen aus der CSV übernehmen
            switch ($Country) {
                "Netherlands"    { $countryCode = "NL"; $countryName = "Netherlands" }
                "United Kingdom" { $countryCode = "GB"; $countryName = "United Kingdom" }
                # Füge hier weitere Länder hinzu
                default {
                     Write-Warning "  Kein spezifischer Ländercode für '$Country' definiert. Verwende '$Country' als Namen."
                     # Optional: Versuche, einen Code zu erraten oder leer zu lassen
                     # $countryCode = $Country # Oder $null
                }
            }
            if ($countryCode -and $countryCode -ne $localADUser.c) { $UpdateAttributes["c"] = $countryCode }
            if ($countryName -and $countryName -ne $localADUser.co) { $UpdateAttributes["co"] = $countryName }

             # --- Manager setzen, falls ein Azure AD Manager vorhanden ist ---
            $managerDistinguishedName = $null
            if ($AzureADManager) {
                Write-Verbose "  Suche lokalen AD Benutzer für Manager '$($AzureADManager.UserPrincipalName)'..."
                $managerLocalAD = Get-ADUser -Filter "UserPrincipalName -eq '$($AzureADManager.UserPrincipalName)'" -Properties DistinguishedName
                if ($managerLocalAD) {
                    Write-Verbose "    Lokaler Manager gefunden: $($managerLocalAD.DistinguishedName)"
                    # Nur setzen, wenn sich der Manager geändert hat
                    if ($managerLocalAD.DistinguishedName -ne $localADUser.Manager) {
                         $managerDistinguishedName = $managerLocalAD.DistinguishedName # Für spätere gemeinsame Aktualisierung
                         Write-Host "    Manager wird auf '$($managerLocalAD.DistinguishedName)' gesetzt." -ForegroundColor Cyan
                    } else {
                        Write-Verbose "    Manager ist bereits korrekt gesetzt."
                    }
                } else {
                    Write-Warning "  Manager mit UPN '$($AzureADManager.UserPrincipalName)' wurde nicht im lokalen AD gefunden. Manager wird nicht gesetzt."
                }
            } else {
                 # Optional: Wenn kein Manager in Azure AD, lokalen Manager entfernen?
                 # if ($localADUser.Manager) { $UpdateAttributes["Manager"] = $null }
                 Write-Verbose "  Kein Manager in Azure AD gefunden. Lokaler Manager wird nicht geändert."
            }

            # Manager zur Hashtable hinzufügen, falls er gesetzt werden soll
            if ($managerDistinguishedName) {
                 $UpdateAttributes["Manager"] = $managerDistinguishedName
            }

            # --- ImmutableID (msDS-ConsistencyGuid) setzen ---
            # Überprüfen, ob das Attribut existiert und ob es sich unterscheidet
            if ($localADUser.PSObject.Properties['msDS-ConsistencyGuid'] -ne $null -and $localADUser.msDS-ConsistencyGuid -ne $immutableID) {
                 $UpdateAttributes['msDS-ConsistencyGuid'] = $immutableID
                 Write-Host "  msDS-ConsistencyGuid (ImmutableID) wird auf '$immutableID' gesetzt." -ForegroundColor Cyan
            } elseif ($localADUser.PSObject.Properties['msDS-ConsistencyGuid'] -eq $null) {
                 # Das Attribut ist möglicherweise nicht im Standard-Property-Set, aber wir versuchen es zu setzen
                 $UpdateAttributes['msDS-ConsistencyGuid'] = $immutableID
                 Write-Host "  msDS-ConsistencyGuid (ImmutableID) wird initial auf '$immutableID' gesetzt." -ForegroundColor Cyan
            } elseif ($localADUser.msDS-ConsistencyGuid -eq $immutableID) {
                 Write-Verbose "  msDS-ConsistencyGuid (ImmutableID) ist bereits korrekt gesetzt."
            }


            # --- Lokalen AD Benutzer aktualisieren, wenn Änderungen vorhanden sind ---
            if ($UpdateAttributes.Count -gt 0) {
                 try {
                    Write-Host "  Aktualisiere Attribute: $($UpdateAttributes.Keys -join ', ')" -ForegroundColor Cyan
                    Set-ADUser -Identity $localADUser.SamAccountName @UpdateAttributes -ErrorAction Stop
                    Write-Host "  Lokaler AD Benutzer '$($localADUser.SamAccountName)' erfolgreich aktualisiert." -ForegroundColor Green
                    return $true
                 } catch {
                    Write-Error "  Fehler beim Aktualisieren des lokalen AD Benutzers '$($localADUser.SamAccountName)': $_"
                    return $false
                 }
            } else {
                Write-Host "  Keine Änderungen für den lokalen AD Benutzer '$($localADUser.SamAccountName)' erforderlich." -ForegroundColor Gray
                return $true # Zählt als Erfolg, da keine Aktion nötig war
            }

        } else {
            Write-Warning "Lokaler AD Benutzer mit UPN '$($AzureADUser.UserPrincipalName)' wurde nicht gefunden."
            return $false
        }
    } catch {
        # Allgemeiner Fehler in der Funktion
        $errMsg = "Unerwarteter Fehler bei der Verarbeitung von '$($AzureADUser.UserPrincipalName)'"
        if($localADUser){ $errMsg += " (Lokaler User: $($localADUser.SamAccountName))"}
        $errMsg += ": $_"
        Write-Error $errMsg
        # Den Fehler weiterleiten, falls spezifischere Behandlung oben fehlschlug
        # throw $_ # Uncomment if you want the script to stop on any error within this function
        return $false
    }
}

#EndRegion Functions

#Region Main Script

# Überprüfen der AzureAD-Verbindung
try {
    Write-Verbose "Überprüfe Azure AD Verbindung..."
    $azureADConnection = Get-AzureADCurrentSessionInfo -ErrorAction Stop
    Write-Host "Verbunden mit Azure AD Tenant: $($azureADConnection.TenantId) ($($azureADConnection.TenantDomain))" -ForegroundColor Green
} catch {
    Write-Error "Nicht mit Azure AD verbunden. Bitte zuerst 'Connect-AzureAD' ausführen."
    exit
}

# CSV-Datei einlesen
try {
    Write-Verbose "Lese CSV-Datei '$CsvPath'..."
    $users = Import-Csv -Path $CsvPath
    $userCount = ($users | Measure-Object).Count
    if ($userCount -eq 0) {
        Write-Error "Die CSV-Datei '$CsvPath' ist leer oder konnte nicht korrekt gelesen werden."
        exit
    }
     # Überprüfen, ob die notwendigen Spalten vorhanden sind
    if (-not $users[0].PSObject.Properties.Match('UserPrincipalName') -or -not $users[0].PSObject.Properties.Match('Country')) {
        Write-Error "Die CSV-Datei muss mindestens die Spalten 'UserPrincipalName' und 'Country' enthalten."
        exit
    }
    Write-Host "[$($userCount)] Benutzer in der CSV-Datei gefunden: '$CsvPath'" -ForegroundColor Cyan
} catch {
    Write-Error "Fehler beim Lesen der CSV-Datei '$CsvPath': $_"
    exit
}

# Ausgabepfad für optionale XML-Dateien (gleicher Ordner wie die CSV)
$outputPath = Join-Path -Path (Split-Path -Path $CsvPath -Parent) -ChildPath "UserSyncExport_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
Write-Verbose "Optionale XML-Exportdateien werden nach '$outputPath' geschrieben (falls aktiviert)."

# Statistik-Variablen
$successCount = 0
$failureCount = 0
$processedCount = 0

# Verarbeitung der Benutzer
Write-Host "`nStarte Synchronisationsprozess..." -ForegroundColor Yellow
$startTime = Get-Date

foreach ($userEntry in $users) {
    $processedCount++
    $upn = $userEntry.UserPrincipalName
    $country = $userEntry.Country

    Write-Host "`n[$($processedCount)/$($userCount)] Verarbeite Benutzer: '$upn'" -ForegroundColor Yellow

    # AzureAD-Benutzer abrufen
    $azureADUser = Get-EntraUserInfo -UserPrincipalName $upn

    if ($azureADUser) {
        # Azure AD Manager abrufen
        $azureADManager = Get-EntraUserManagerInfo -AzureADUserObjectId $azureADUser.ObjectId

        # Optional: XML-Dateien exportieren
        # Export-UserAndManagerInfoToXml -AzureADUser $azureADUser -AzureADManager $azureADManager -OutputPath $outputPath

        # Lokalen AD-Benutzer aktualisieren (inkl. Manager und ImmutableID)
        $updateResult = Set-LocalADUserAttributes -AzureADUser $azureADUser -AzureADManager $azureADManager -Country $country
        if ($updateResult) {
            $successCount++
        } else {
            $failureCount++
        }
    } else {
        # Fehler wurde bereits in Get-EntraUserInfo geloggt
        $failureCount++
    }
}

# Zusammenfassung anzeigen
$endTime = Get-Date
$duration = New-TimeSpan -Start $startTime -End $endTime

Write-Host "`nSynchronisation abgeschlossen." -ForegroundColor Green
Write-Host "--------------------------------------------------"
Write-Host "Verarbeitet:         $processedCount Benutzer"
Write-Host "Erfolgreich:         $successCount Benutzer" -ForegroundColor Green
if ($failureCount -gt 0) {
    Write-Host "Fehlgeschlagen:        $failureCount Benutzer" -ForegroundColor Red
}
Write-Host "Dauer:               $($duration.ToString('g'))"
Write-Host "Optionale XML-Exporte in: '$outputPath'"
Write-Host "--------------------------------------------------"


# Optional: Entra Delta Sync starten
$runSync = Read-Host "Möchten Sie jetzt eine Entra Connect Delta Synchronisation starten? (J/N)"
if ($runSync -eq "J" -or $runSync -eq "j") {
    Write-Host "Versuche Delta Synchronisation zu starten..."
    try {
        # Sicherstellen, dass das ADSync-Modul geladen ist
        if (-not (Get-Module -Name ADSync -ErrorAction SilentlyContinue)) {
             # Standardpfad für das Modul
             $syncModulePath = "C:\Program Files\Microsoft Azure AD Sync\Bin\ADSync\ADSync.psd1"
             if(Test-Path $syncModulePath){
                 Import-Module $syncModulePath -ErrorAction Stop
             } else {
                 throw "ADSync Modul nicht gefunden. Stellen Sie sicher, dass Entra Connect (Azure AD Connect) auf diesem Server installiert ist oder importieren Sie das Modul manuell."
             }
        }
        Start-ADSyncSyncCycle -PolicyType Delta -ErrorAction Stop
        Write-Host "Delta Synchronisation erfolgreich gestartet." -ForegroundColor Green
    } catch {
        Write-Error "Fehler beim Starten der Delta Synchronisation: $_"
        Write-Warning "Stellen Sie sicher, dass das Skript auf dem Entra Connect (Azure AD Connect) Server ausgeführt wird oder das ADSync Modul importiert werden kann."
    }
} else {
    Write-Host "Delta Synchronisation nicht gestartet."
}

#EndRegion Main Script

#EndRegion Main Script
