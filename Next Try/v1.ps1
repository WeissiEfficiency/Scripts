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
        
        # Benutzerinformationen exportieren
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
    }
}