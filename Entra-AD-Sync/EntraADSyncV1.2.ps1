<#
.SYNOPSIS
    Synchronizes attributes from Azure AD users to local AD users based on UserPrincipalName.

.DESCRIPTION
    This script synchronizes attributes from Azure AD (Entra ID) users to local Active Directory users.
    It reads a CSV file containing Azure AD UserPrincipalNames and Country information, 
    retrieves the user information from Azure AD, and updates the corresponding local AD user attributes.
    The script also sets the manager relationship and handles country code conversion.

.PARAMETER CsvPath
    Path to the CSV file containing user information. The CSV must include UserPrincipalName and Country columns.

.EXAMPLE
    .\Sync-AzureADToLocalAD.ps1 -CsvPath "C:\Temp\users.csv"

.NOTES
    Prerequisites:
    - Azure AD PowerShell module must be installed
    - RSAT AD PowerShell tools must be installed
    - Must be connected to Azure AD using Connect-AzureAD before running this script
    - Must have appropriate permissions in both Azure AD and local AD

.AUTHOR
    Created on: April 16, 2025
#>

# Installation der benötigten RSAT-Tools (auskommentiert, bei Bedarf ausführen)
<#
Install-WindowsFeature RSAT-AD-PowerShell
Import-Module ActiveDirectory
#>

param(
    [Parameter(Mandatory=$true)]
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
        $azureADUser = Get-AzureADUser -Filter "userPrincipalName eq '$UserPrincipalName'"
        if ($azureADUser) {
            return $azureADUser
        } else {
            Write-Warning "User $UserPrincipalName not found in Azure AD."
            return $null
        }
    } catch {
        Write-Error "Error retrieving Azure AD user $UserPrincipalName: $_"
        return $null
    }
}

# Funktion zum Exportieren der Benutzer- und Manager-Informationen in XML-Dateien
function Export-UserManagerInfo {
    param(
        [Parameter(Mandatory=$true)]
        $AzureADUser,
        [Parameter(Mandatory=$true)]
        [string]$OutputPath
    )
    
    try {
        # Exportieren der Benutzerinformationen
        $xmlUserPath = Join-Path -Path $OutputPath -ChildPath "$($AzureADUser.UserPrincipalName).user.xml"
        $AzureADUser | Export-Clixml -Path $xmlUserPath
        
        # Abrufen und Exportieren der Manager-Informationen
        try {
            $manager = Get-AzureADUserManager -ObjectId $AzureADUser.ObjectId
            if ($manager) {
                $xmlManagerPath = Join-Path -Path $OutputPath -ChildPath "$($AzureADUser.UserPrincipalName).manager.xml"
                $manager | Export-Clixml -Path $xmlManagerPath
                return $manager
            }
        } catch {
            Write-Warning "  No manager found for user $($AzureADUser.UserPrincipalName)"
        }
        
        return $null
    } catch {
        Write-Error "Error exporting XML files for $($AzureADUser.UserPrincipalName): $_"
        return $null
    }
}

# Funktion zum Aktualisieren des lokalen AD-Benutzers mit AzureAD-Attributen
function Set-LocalADUserAttributes {
    param(
        [Parameter(Mandatory=$true)]
        $AzureADUser,
        [Parameter(Mandatory=$false)]
        $AzureADManager,
        [Parameter(Mandatory=$true)]
        [string]$Country
    )
    
    try {
        # Lokalen AD-Benutzer anhand des UPN finden
        $localADUser = Get-ADUser -Filter "UserPrincipalName -eq '$($AzureADUser.UserPrincipalName)'" -Properties *
        
        if ($localADUser) {
            # Hashtable für zu aktualisierende Attribute erstellen
            $UpdateAttributes = @{}
            
            # Attribute aus AzureAD übernehmen, wenn sie vorhanden sind
            if ($AzureADUser.TelephoneNumber) { $UpdateAttributes["OfficePhone"] = $AzureADUser.TelephoneNumber }
            if ($AzureADUser.StreetAddress) { $UpdateAttributes["StreetAddress"] = $AzureADUser.StreetAddress }
            if ($AzureADUser.PostalCode) { $UpdateAttributes["PostalCode"] = $AzureADUser.PostalCode }
            if ($AzureADUser.JobTitle) { $UpdateAttributes["Title"] = $AzureADUser.JobTitle }
            if ($AzureADUser.Department) { $UpdateAttributes["Department"] = $AzureADUser.Department }
            if ($AzureADUser.City) { $UpdateAttributes["City"] = $AzureADUser.City }
            if ($AzureADUser.CompanyName) { $UpdateAttributes["Company"] = $AzureADUser.CompanyName }
            
            # Ländercode konvertieren und zum Hashtable hinzufügen
            $countryCode = switch ($Country) {
                "Netherlands" { "NL" }
                "United Kingdom" { "GB" }
                default { $Country }
            }
            $UpdateAttributes["Country"] = $countryCode
            
            # Manager setzen, falls vorhanden
            if ($AzureADManager) {
                $managerLocalAD = Get-ADUser -Filter "UserPrincipalName -eq '$($AzureADManager.UserPrincipalName)'"
                if ($managerLocalAD) {
                    Set-ADUser -Identity $localADUser.SamAccountName -Manager $managerLocalAD.DistinguishedName
                }
            }
            
            # Attribute auf einmal aktualisieren
            Set-ADUser -Identity $localADUser.SamAccountName @UpdateAttributes
            
            # ImmutableID setzen (basierend auf ObjectGUID des AD-Benutzers)
            $immutableID = [System.Convert]::ToBase64String($localADUser.ObjectGUID.ToByteArray())
            
            Write-Host "Local AD user $($localADUser.SamAccountName) updated. ImmutableID: $immutableID" -ForegroundColor Green
            Write-Host "Updated attributes: $($UpdateAttributes.Keys -join ', ')" -ForegroundColor Cyan
            
            return $true
        } else {
            Write-Warning "Local AD user with UPN $($AzureADUser.UserPrincipalName) not found."
            return $false
        }
    } catch {
        Write-Error "Error updating local AD user $($AzureADUser.UserPrincipalName): $_"
        return $false
    }
}

#EndRegion Functions

#Region Main Script

# Überprüfen der AzureAD-Verbindung
try {
    $azureADConnection = Get-AzureADCurrentSessionInfo -ErrorAction Stop
    Write-Host "Connected to Azure AD tenant: $($azureADConnection.TenantDomain)" -ForegroundColor Green
} catch {
    Write-Error "Not connected to Azure AD. Please run Connect-AzureAD first."
    exit
}

# Überprüfen, ob die CSV-Datei existiert
if (-not (Test-Path -Path $CsvPath)) {
    Write-Error "The specified CSV file was not found: $CsvPath"
    exit
}

# CSV-Datei einlesen
try {
    $users = Import-Csv -Path $CsvPath
    $userCount = ($users | Measure-Object).Count
    Write-Host "Found $userCount users in CSV file." -ForegroundColor Cyan
} catch {
    Write-Error "Error reading CSV file: $_"
    exit
}

# Ausgabepfad für XML-Dateien (gleicher Ordner wie die CSV)
$outputPath = Split-Path -Path $CsvPath -Parent

# Statistik-Variablen
$successCount = 0
$failureCount = 0

# Verarbeitung der Benutzer
Write-Host "Starting synchronization process..." -ForegroundColor Yellow
foreach ($user in $users) {
    $upn = $user.UserPrincipalName
    $country = $user.Country
    
    Write-Host "`nProcessing user: $upn" -ForegroundColor Yellow
    
    # AzureAD-Benutzer abrufen
    $azureADUser = Get-EntraUserInfo -UserPrincipalName $upn
    
    if ($azureADUser) {
        # XML-Dateien exportieren und Manager-Informationen abrufen
        $manager = Export-UserManagerInfo -AzureADUser $azureADUser -OutputPath $outputPath
        
        if ($manager) {
            Write-Host "  Manager found: $($manager.DisplayName)" -ForegroundColor Cyan
        }
        
        # Lokalen AD-Benutzer aktualisieren
        $updateResult = Set-LocalADUserAttributes -AzureADUser $azureADUser -AzureADManager $manager -Country $country
        if ($updateResult) {
            $successCount++
        } else {
            $failureCount++
        }
    } else {
        $failureCount++
    }
}

# Zusammenfassung anzeigen
Write-Host "`nSynchronization complete." -ForegroundColor Green
Write-Host "Successfully processed: $successCount users" -ForegroundColor Green
if ($failureCount -gt 0) {
    Write-Host "Failed to process: $failureCount users" -ForegroundColor Red
}

# Entra Delta Sync starten
$runSync = Read-Host "Would you like to start an Entra Delta Sync? (Y/N)"
if ($runSync -eq "Y" -or $runSync -eq "y") {
    # Azure AD Connect Delta Sync starten
    try {
        Import-Module ADSync
        Start-ADSyncSyncCycle -PolicyType Delta
        Write-Host "Delta synchronization has been started." -ForegroundColor Green
    } catch {
        Write-Error "Error starting delta synchronization: $_"
    }
}

#EndRegion Main Script
