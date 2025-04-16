<#
.SYNOPSIS
Synchronisiert ausgewählte Benutzerattribute (Land, Manager) von Microsoft Entra ID (Azure AD)
zum lokalen Active Directory basierend auf einer CSV-Datei.

.DESCRIPTION
Dieses Skript liest Benutzer-UPNs und Länderinformationen aus einer CSV-Datei.
Für jeden Benutzer:
1. Ruft das Benutzerobjekt und dessen Manager aus Entra ID ab.
2. Speichert diese Objekte optional als XML-Dateien zur Sicherung/Überprüfung.
3. Sucht den entsprechenden Benutzer im lokalen AD anhand des UPN.
4. Aktualisiert das Land (mit Mapping 'Netherlands'->'NL', 'United Kingdom'->'GB')
   und den Manager (findet den lokalen Manager über dessen UPN) im lokalen AD.
5. Berechnet und zeigt die ImmutableID (ObjectGUID als Base64) des lokalen Benutzers an.
6. Fragt am Ende, ob ein Entra ID Connect (Azure AD Connect) Delta-Synchronisationszyklus gestartet werden soll.

.PARAMETER CsvPath
Pfad zur CSV-Datei. Die Datei MUSS die Spalten 'UserPrincipalName' und 'Country' enthalten.

.EXAMPLE
.\Sync-EntraToLocalAD.ps1 -CsvPath "C:\temp\users_to_sync.csv"

.NOTES
Autor: Ihr Name / Ihre Organisation
Datum: 2025-04-16
Version: 1.0

Voraussetzungen:
- PowerShell-Modul 'AzureAD' muss installiert sein (`Install-Module AzureAD`).
- PowerShell-Modul 'ActiveDirectory' (Teil der RSAT-Tools) muss installiert sein.
  # Install-WindowsFeature RSAT-AD-PowerShell
- Eine aktive Verbindung zu Entra ID muss bestehen (`Connect-AzureAD`).
- Ausreichende Berechtigungen in Entra ID (Lesen von Benutzern/Managern) und lokalem AD (Lesen/Schreiben von Benutzern).
- Wenn der Delta-Sync ausgelöst werden soll: Ausführung auf dem AD Connect Server oder über Remoting mit entsprechenden Berechtigungen.
#>


param(
   
   
    [string]$CsvPath
)

# --- Vorbemerkungen ---
Write-Host "INFO: Skript gestartet."
Write-Host "INFO: Stellen Sie sicher, dass Sie bereits mit Microsoft Entra ID (Azure AD) verbunden sind (Connect-AzureAD)."
Write-Host "INFO: Stellen Sie sicher, dass die RSAT AD PowerShell Tools installiert sind (`Install-WindowsFeature RSAT-AD-PowerShell`)."
Write-Host "INFO: Das AzureAD Modul wird benötigt (`Install-Module AzureAD`)."

# --- Funktionen ---

function Get-AzureAdUserData {
   
    param(
        [Parameter(Mandatory=$true)]
        [string]$UPN
    )

    Write-Verbose "Versuche, Entra ID Benutzerdaten für UPN '$UPN' abzurufen."
    $azureUser = $null
    $azureManager = $null

    try {
        # Benutzerobjekt abrufen
        $azureUser = Get-AzureADUser -ObjectId $UPN -ErrorAction Stop
        if ($azureUser) {
            Write-Verbose "Entra ID Benutzer gefunden: $($azureUser.ObjectId)"
            # Manager des Benutzers abrufen
            try {
                # Hinweis: Get-AzureADUserManager gibt $null zurück, wenn kein Manager zugewiesen ist, kein Fehler.
                $azureManager = Get-AzureADUserManager -ObjectId $azureUser.ObjectId -ErrorAction SilentlyContinue # Fehler hier nicht stoppen lassen
                if ($azureManager) {
                    Write-Verbose "Entra ID Manager gefunden: $($azureManager.UserPrincipalName)"
                } else {
                    Write-Verbose "Kein Manager für Benutzer '$UPN' in Entra ID gefunden."
                }
            } catch {
                # Fehler beim Manager-Abruf spezifisch behandeln (z.B. Berechtigungsproblem)
                Write-Warning "WARNUNG: Fehler beim Abrufen des Managers für '$UPN'. Details: $($_.Exception.Message)"
            }
        } else {
             # Sollte durch -ErrorAction Stop oben nicht erreicht werden, aber zur Sicherheit
             Write-Verbose "Kein Entra ID Benutzer mit UPN '$UPN' gefunden."
        }
    } catch {
        # Allgemeiner Fehler beim Benutzerabruf
        Write-Warning "WARNUNG: Fehler beim Abrufen des Entra ID Benutzers '$UPN'. Details: $($_.Exception.Message)"
    }

    # Benutzer- und Managerobjekt (kann $null sein) zurückgeben
    return @{
        User    = $azureUser
        Manager = $azureManager
    }
}

function Export-UserDataToXml {
   
    param(
        [Parameter(Mandatory=$true)]
        $AzureUser, # Das Entra ID Benutzerobjekt

        $AzureManager, # Das Entra ID Managerobjekt (kann $null sein)

        [Parameter(Mandatory=$true)]
        [string]$CsvPath # Pfad zur CSV-Datei zur Bestimmung des Ausgabeordners
    )

    try {
        # Ausgabeordner aus dem CSV-Pfad ableiten
        $outputDir = Split-Path -Path $CsvPath -Parent
        # Basis-Dateiname aus dem UPN erstellen (Sonderzeichen ersetzen für Robustheit)
        $safeUpnPart = $AzureUser.UserPrincipalName -replace '[^a-zA-Z0-9.-_]', '_'
        $userFileNameBase = $safeUpnPart

        # Pfade für XML-Dateien erstellen
        $userXmlPath = Join-Path -Path $outputDir -ChildPath "$($userFileNameBase).user.xml"
        $managerXmlPath = Join-Path -Path $outputDir -ChildPath "$($userFileNameBase).manager.xml"

        # Benutzerobjekt exportieren
        Write-Verbose "Exportiere Benutzerdaten nach '$userXmlPath'"
        Export-CliXml -Path $userXmlPath -InputObject $AzureUser -Force

        # Managerobjekt exportieren, falls vorhanden
        if ($AzureManager) {
            Write-Verbose "Exportiere Managerdaten nach '$managerXmlPath'"
            Export-CliXml -Path $managerXmlPath -InputObject $AzureManager -Force
        } else {
            Write-Verbose "Kein Managerobjekt zum Exportieren vorhanden für Benutzer '$($AzureUser.UserPrincipalName)'."
            # Optional: Alte Manager-Datei löschen, falls vorhanden und kein Manager mehr da ist
            if (Test-Path $managerXmlPath) {
                Write-Verbose "Entferne alte Manager-XML-Datei '$managerXmlPath'."
                Remove-Item $managerXmlPath -Force -ErrorAction SilentlyContinue
            }
        }
    } catch {
        Write-Warning "WARNUNG: Fehler beim Exportieren der Daten nach XML für Benutzer '$($AzureUser.UserPrincipalName)'. Details: $($_.Exception.Message)"
    }
}

function Update-LocalAdUser {
   
    param(
        [Parameter(Mandatory=$true)]
        [string]$UPN, # UPN zur Identifizierung des lokalen AD Benutzers

        [Parameter(Mandatory=$true)]
        $AzureUser, # Entra ID Benutzerobjekt als Quelle

        $AzureManager, # Entra ID Managerobjekt (kann $null sein)

        [Parameter(Mandatory=$true)]
        [string]$CsvCountry # Länderinformation aus der CSV-Datei
    )

    Write-Verbose "Versuche, lokalen AD Benutzer mit UPN '$UPN' zu aktualisieren."
    $localAdUser = $null
    try {
        # Lokalen AD Benutzer anhand des UPN finden
        # Wichtig: Benötigte Properties explizit anfordern (-Properties)
        $localAdUser = Get-ADUser -Filter "UserPrincipalName -eq '$UPN'" -Properties Country, c, Manager, ObjectGUID, DistinguishedName -ErrorAction Stop
    } catch {
        Write-Warning "WARNUNG: Fehler beim Suchen des lokalen AD Benutzers '$UPN'. Details: $($_.Exception.Message)"
        return # Verarbeitung für diesen Benutzer hier abbrechen
    }

    if (-not $localAdUser) {
        Write-Warning "WARNUNG: Lokaler AD Benutzer mit UPN '$UPN' nicht gefunden. Keine Aktualisierung möglich."
        return # Verarbeitung für diesen Benutzer hier abbrechen
    }

    Write-Verbose "Lokalen AD Benutzer gefunden: $($localAdUser.DistinguishedName)"

    # Hashtable für die zu aktualisierenden AD-Attribute vorbereiten
    $adUserProperties = @{}

    # Länder-Mapping basierend auf der CSV-Information
    $targetCountryCode = $null
    switch ($CsvCountry.Trim().ToLower()) { # Trim() entfernt Leerzeichen, ToLower() für Groß-/Kleinschreibung
        'netherlands'    { $targetCountryCode = 'NL' }
        'united kingdom' { $targetCountryCode = 'GB' }
        # default { Write-Verbose "Kein Mapping für Land '$CsvCountry' definiert." }
    }

    if ($targetCountryCode) {
        # Nur aktualisieren, wenn sich der Wert ändert oder noch nicht gesetzt ist
        if ($localAdUser.Country -ne $targetCountryCode -or $localAdUser.c -ne $targetCountryCode) {
             Write-Verbose "Setze Landeskennzeichen auf '$targetCountryCode' für Benutzer '$UPN'."
             $adUserProperties.Country = $targetCountryCode # Attribut 'co' (Country Name)
             $adUserProperties.c = $targetCountryCode      # Attribut 'c' (ISO 3166 Country Code)
        } else {
             Write-Verbose "Landeskennzeichen '$targetCountryCode' ist bereits korrekt gesetzt für '$UPN'."
        }
    } else {
         Write-Verbose "Kein gültiges Mapping für Land '$CsvCountry' gefunden oder Land nicht in Mapping-Liste. Land wird nicht aktualisiert für '$UPN'."
    }

    # --- HIER WEITERE ATTRIBUT-MAPPINGS EINFÜGEN ---
    # Beispiel: Abteilung und Titel aus Entra ID übernehmen, falls vorhanden
    # if ($AzureUser.Department -and $localAdUser.Department -ne $AzureUser.Department) {
    #    Write-Verbose "Setze Abteilung auf '$($AzureUser.Department)'."
    #    $adUserProperties.Department = $AzureUser.Department
    # }
    # if ($AzureUser.JobTitle -and $localAdUser.Title -ne $AzureUser.JobTitle) {
    #    Write-Verbose "Setze Titel auf '$($AzureUser.JobTitle)'."
    #    $adUserProperties.Title = $AzureUser.JobTitle
    # }
    # -------------------------------------------------

    # Manager aktualisieren
    $localAdManagerDN = $null # Zielwert für das Manager-Attribut (DistinguishedName)
    $updateManager = $false   # Flag, ob Manager-Attribut geändert werden soll

    if ($AzureManager -and $AzureManager.UserPrincipalName) {
        Write-Verbose "Entra ID Manager '$($AzureManager.UserPrincipalName)' vorhanden. Suche entsprechenden lokalen AD Benutzer."
        try {
            # Lokalen AD Manager anhand des UPN des Entra ID Managers finden
            $localAdManager = Get-ADUser -Filter "UserPrincipalName -eq '$($AzureManager.UserPrincipalName)'" -Properties DistinguishedName -ErrorAction Stop
            if ($localAdManager) {
                Write-Verbose "Lokaler AD Manager gefunden: $($localAdManager.DistinguishedName)."
                $localAdManagerDN = $localAdManager.DistinguishedName
                # Prüfen, ob der Manager geändert werden muss
                if ($localAdUser.Manager -ne $localAdManagerDN) {
                    Write-Verbose "Manager muss aktualisiert werden."
                    $updateManager = $true
                } else {
                    Write-Verbose "Manager ist bereits korrekt gesetzt."
                }
            } else {
                Write-Warning "WARNUNG: Lokaler AD Manager mit UPN '$($AzureManager.UserPrincipalName)' nicht gefunden. Manager wird NICHT gesetzt für '$UPN'."
                # Hier wird der Manager NICHT auf $null gesetzt, es sei denn, das ist explizit gewünscht.
            }
        } catch {
            Write-Warning "WARNUNG: Fehler beim Suchen des lokalen AD Managers '$($AzureManager.UserPrincipalName)'. Details: $($_.Exception.Message)"
        }
    } else {
        Write-Verbose "Kein Manager in Entra ID Daten vorhanden oder Manager hat keinen UPN."
        # Prüfen, ob der lokale Manager entfernt werden soll, wenn in Entra keiner (mehr) ist
        if ($localAdUser.Manager) {
             Write-Verbose "Lokaler AD Benutzer '$UPN' hat aktuell einen Manager, aber in Entra ID ist keiner (mehr) definiert. Setze Manager auf $null."
             $localAdManagerDN = $null # Explizit auf $null setzen
             $updateManager = $true
        } else {
             Write-Verbose "Kein Manager in Entra ID und auch lokal keiner gesetzt. Keine Änderung am Manager-Attribut."
        }
    }

    # Manager zur Hashtable hinzufügen, wenn er aktualisiert werden soll
    if ($updateManager) {
        $adUserProperties.Manager = $localAdManagerDN # Kann DN oder $null sein
    }

    # Änderungen im AD anwenden, wenn welche vorhanden sind
    if ($adUserProperties.Keys.Count -gt 0) {
        try {
            Write-Host "INFO: Aktualisiere Attribute für lokalen AD Benutzer '$UPN': $($adUserProperties.Keys -join ', ')"
            Set-ADUser -Identity $localAdUser -Replace $adUserProperties -ErrorAction Stop # -Replace ist oft sicherer als -Set
            Write-Verbose "Attribute erfolgreich aktualisiert."
        } catch {
            Write-Error "FEHLER: Konnte Attribute für lokalen AD Benutzer '$UPN' nicht aktualisieren. Details: $($_.Exception.Message)"
        }
    } else {
        Write-Host "INFO: Keine zu aktualisierenden Attribute für '$UPN' identifiziert."
    }

    # ImmutableID berechnen und ausgeben (nach potenzieller Aktualisierung, basierend auf dem *aktuellen* Objekt)
    try {
        # ObjectGUID neu laden, falls sie sich theoretisch ändern könnte (unwahrscheinlich, aber sicher)
        $refreshedUser = Get-ADUser -Identity $localAdUser.ObjectGUID -Properties ObjectGUID
        $immutableId =::ToBase64String($refreshedUser.ObjectGUID.ToByteArray())
        Write-Host "INFO: Berechnete ImmutableID (basierend auf objectGUID) für '$UPN': $immutableId"
    } catch {
        Write-Warning "WARNUNG: Konnte ImmutableID für '$UPN' nicht berechnen. Details: $($_.Exception.Message)"
    }
}

# --- Hauptskript ---

# CSV-Datei importieren
try {
    # Sicherstellen, dass die erwarteten Spalten vorhanden sind
    $csvHeaders = (Import-Csv -Path $CsvPath -TotalCount 1).PSObject.Properties.Name
    if (-not ($csvHeaders -contains 'UserPrincipalName' -and $csvHeaders -contains 'Country')) {
        throw "CSV-Datei '$CsvPath' muss die Spalten 'UserPrincipalName' und 'Country' enthalten."
    }
    $usersToProcess = Import-Csv -Path $CsvPath -ErrorAction Stop
    Write-Host "INFO: $($usersToProcess.Count) Benutzerdatensätze aus '$CsvPath' eingelesen."
} catch {
    Write-Error "FEHLER: CSV-Datei '$CsvPath' konnte nicht gelesen oder validiert werden. Details: $($_.Exception.Message)"
    exit 1 # Skript beenden
}

# Verarbeitung jedes Benutzers aus der CSV
foreach ($record in $usersToProcess) {
    # Trim() entfernt führende/nachfolgende Leerzeichen aus den CSV-Werten
    $upn = $record.UserPrincipalName.Trim()
    $csvCountry = $record.Country.Trim()

    # Überspringen, wenn UPN leer ist
    if ([string]::IsNullOrWhiteSpace($upn)) {
        Write-Warning "WARNUNG: Leerer UserPrincipalName in Zeile $($usersToProcess.IndexOf($record) + 2) der CSV gefunden. Überspringe."
        continue
    }

    Write-Host "--- Verarbeitung von Benutzer: $upn ---"

    # Schritt 1: Entra ID Daten abrufen
    $azureData = Get-AzureAdUserData -UPN $upn -Verbose:$VerbosePreference
    # Minimalprüfung, ob Benutzer gefunden wurde
    if ($azureData -and $azureData.User) {
        Write-Host "INFO: Entra ID Benutzer '$upn' gefunden (ObjectID: $($azureData.User.ObjectId))."

        # Schritt 2: Entra ID Daten nach XML exportieren
        Export-UserDataToXml -AzureUser $azureData.User -AzureManager $azureData.Manager -CsvPath $CsvPath -Verbose:$VerbosePreference

        # Schritt 3: Lokalen AD Benutzer aktualisieren
        Update-LocalAdUser -UPN $upn -AzureUser $azureData.User -AzureManager $azureData.Manager -CsvCountry $csvCountry -Verbose:$VerbosePreference

    } else {
        # Warnung wurde bereits in Get-AzureAdUserData ausgegeben
        Write-Warning "WARNUNG: Benutzer '$upn' nicht in Entra ID gefunden oder Fehler beim Abruf. Überspringe lokale AD Aktualisierung."
    }
    Write-Host "--- Verarbeitung von Benutzer: $upn abgeschlossen ---`n"
}

Write-Host "INFO: Verarbeitung aller Benutzer aus der CSV-Datei abgeschlossen."

# Nach Abschluss der Schleife, Abfrage für Delta Sync
$triggerSync = Read-Host "Möchten Sie jetzt einen Microsoft Entra Connect (Azure AD Connect) Delta Sync-Zyklus starten? (j/N)"

if ($triggerSync -eq 'j') {
    Write-Host "INFO: Starte Entra Connect Delta Sync..."
    Write-Warning "WICHTIG: Dieser Befehl muss auf dem AD Connect Server ausgeführt werden oder über PowerShell Remoting zu diesem Server."
    try {
        # Prüfen, ob das Modul geladen ist oder geladen werden kann
        if (-not (Get-Module -Name ADSync -ErrorAction SilentlyContinue)) {
            Import-Module ADSync -ErrorAction Stop
        }
        Start-ADSyncSyncCycle -PolicyType Delta -ErrorAction Stop
        Write-Host "INFO: Delta Sync erfolgreich gestartet."
    } catch {
        Write-Error "FEHLER: Konnte den Delta Sync nicht starten. Stellen Sie sicher, dass Sie auf dem AD Connect Server sind, das ADSync-Modul verfügbar ist und Sie über die nötigen Berechtigungen verfügen. Details: $($_.Exception.Message)"
    }
} else {
    Write-Host "INFO: Entra Connect Delta Sync wird übersprungen."
}

Write-Host "INFO: Skript beendet."
