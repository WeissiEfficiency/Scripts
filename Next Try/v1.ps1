param(
    [Parameter(Mandatory=$true)]
    [string]$CsvPath
)

# Überprüfen, ob die CSV-Datei existiert
if (-not (Test-Path -Path $CsvPath)) {
    Write-Error "CSV file not found: $CsvPath"
    exit 1
}

# Ordnerpfad der CSV-Datei ermitteln
$csvFolder = Split-Path -Path $CsvPath -Parent

# CSV-Datei einlesen
try {
    $csvUsers = Import-Csv -Path $CsvPath -ErrorAction Stop
} catch {
    Write-Error "Failed to read CSV file: $($_.Exception.Message)"
    exit 1
}

foreach ($csvUser in $csvUsers) {
    $upn = $csvUser.UserPrincipalName
    if (-not $upn) {
        Write-Verbose "Skipping record with missing UserPrincipalName."
        continue
    }
    Write-Host "Processing user: $upn" -ForegroundColor Cyan

    # Azure AD Benutzer abrufen
    try {
        $azureUser = Get-AzureADUser -Filter "UserPrincipalName eq '$upn'" -ErrorAction Stop
    } catch {
        Write-Verbose "Could not retrieve Azure AD user for $upn"
        continue
    }

    if ($azureUser) {
        # Erzeuge einen sicheren Dateinamen (ungültige Zeichen ersetzen)
        $safeUpn = $azureUser.UserPrincipalName -replace '[\\/:*?"<>|]', '_'
        
        # Export Azure AD user information
        $userXmlFile = Join-Path -Path $csvFolder -ChildPath "$safeUpn.user.xml"
        $azureUser | Export-Clixml -Path $userXmlFile
        Write-Host "Exported user data to: $userXmlFile" -ForegroundColor Green

        # Managerinformationen abrufen
        try {
            $azureManager = Get-AzureADUserManager -ObjectId $azureUser.ObjectId -ErrorAction Stop
        } catch {
            Write-Warning "No manager found for $upn"
            $azureManager = $null
        }

        if ($azureManager) {
            # Managerinformationen exportieren
            $managerXmlFile = Join-Path -Path $csvFolder -ChildPath "$safeUpn.manager.xml"
            $azureManager | Export-Clixml -Path $managerXmlFile
            Write-Host "Exported manager data to: $managerXmlFile" -ForegroundColor Green
        }

        # Lokalen AD Benutzer anhand des UPNs abfragen
        try {
            $localAdUser = Get-ADUser -Filter "UserPrincipalName -eq '$($azureUser.UserPrincipalName)'" -ErrorAction Stop
        } catch {
            Write-Warning "Local AD user not found for $($azureUser.UserPrincipalName)"
            continue
        }
        
        # Lokale AD Attribute aktualisieren – einzelne Set-ADUser Aufrufe
        try {
            Set-ADUser -Identity $localAdUser.DistinguishedName -OfficePhone $azureUser.TelephoneNumber
            Set-ADUser -Identity $localAdUser.DistinguishedName -StreetAddress $azureUser.StreetAddress
            Set-ADUser -Identity $localAdUser.DistinguishedName -PostalCode $azureUser.PostalCode
            Set-ADUser -Identity $localAdUser.DistinguishedName -Title $azureUser.JobTitle
            Set-ADUser -Identity $localAdUser.DistinguishedName -Department $azureUser.Department
            Set-ADUser -Identity $localAdUser.DistinguishedName -City $azureUser.City
            Set-ADUser -Identity $localAdUser.DistinguishedName -Company $azureUser.CompanyName
            Set-ADUser -Identity $localAdUser.DistinguishedName -Country $azureUser.Country
            Write-Host "Updated local AD attributes for $($azureUser.UserPrincipalName)" -ForegroundColor Green
        } catch {
            Write-Warning "Failed to update local AD attributes for $($azureUser.UserPrincipalName): $($_.Exception.Message)"
        }

        # Lokalen AD Manager setzen, falls vorhanden
        if ($azureManager) {
            try {
                 $localManager = Get-ADUser -Filter "UserPrincipalName -eq '$($azureManager.UserPrincipalName)'" -ErrorAction Stop
                 Set-ADUser -Identity $localAdUser.DistinguishedName -Manager $localManager.DistinguishedName
                 Write-Host "Set local AD manager for $($azureUser.UserPrincipalName)" -ForegroundColor Green
            } catch {
                 Write-Warning "Failed to set local AD manager for $($azureUser.UserPrincipalName): $($_.Exception.Message)"
            }
        }

        # ImmutableID berechnen und ausgeben 
        try {
            $refreshedUser = Get-ADUser -Identity $localAdUser.ObjectGUID -Properties ObjectGUID
            $immutableId = [System.Convert]::ToBase64String($refreshedUser.ObjectGUID.ToByteArray())
            Write-Host "INFO: Berechnete ImmutableID (basierend auf ObjectGUID) für '$upn': $immutableId"
        } catch {
            Write-Verbose "WARNUNG: Konnte ImmutableID für '$upn' nicht berechnen. Details: $($_.Exception.Message)"
        }

        # ExtensionAttribute10 auf "365" setzen
        try {
            Set-ADUser -Identity $localAdUser.DistinguishedName -Replace @{extensionAttribute10 = "365"}
            Write-Host "Set extensionAttribute10 to 365 for $($azureUser.UserPrincipalName)" -ForegroundColor Green
        } catch {
            Write-Warning "Failed to set extensionAttribute10 for $($azureUser.UserPrincipalName): $($_.Exception.Message)"
        }
    }
# Ende der foreach-Schleife
}

Write-Host "Führen Sie jetzt einen Entra Sync Delta aus, indem Sie den entsprechenden Befehl starten." -ForegroundColor Cyan
Write-Host "Hinweis: Nach dem Sync Delta sollten die Remote-Mailboxen aktiviert werden." -ForegroundColor Cyan