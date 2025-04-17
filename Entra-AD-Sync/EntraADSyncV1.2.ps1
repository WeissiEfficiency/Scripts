<#
.SYNOPSIS
    Synchronizes attributes from Azure AD (Entra ID) users to local AD users based on UserPrincipalName.

.DESCRIPTION
    This script synchronizes selected attributes from Azure AD (Entra ID) users to on-premises Active Directory users.
    It reads a CSV file containing Azure AD UserPrincipalNames and Country information, 
    retrieves the user information from Azure AD using the AzureAD module, finds the corresponding local AD user,
    and updates specified attributes in the local AD.
    It also handles manager relationships based on UPN and converts country names to ISO codes.
    The script calculates the ImmutableID based on the local AD ObjectGUID but DOES NOT set it in Azure AD.

.PARAMETER CsvPath
    Path to the CSV file containing user information. 
    The CSV must include at least 'UserPrincipalName' and 'Country' columns.

.EXAMPLE
    .\Sync-AzureADToLocalAD.ps1 -CsvPath "C:\Temp\users.csv"

.NOTES
    Prerequisites:
    - AzureAD PowerShell module must be installed (`Install-Module AzureAD`). 
      Note: Microsoft Graph PowerShell SDK (`Microsoft.Graph`) is the modern replacement.
    - RSAT AD PowerShell tools must be installed (`Install-WindowsFeature RSAT-AD-PowerShell`).
    - Must be connected to Azure AD using `Connect-AzureAD` before running this script.
    - Must have appropriate permissions in both Azure AD and local AD.
    - Review attribute mappings (especially country codes and target attributes like 'c' vs 'co') and adjust if needed.
    - The script calculates the ImmutableID but does NOT set it in Azure AD. Setting ImmutableID manually is risky and typically handled by Azure AD Connect.

.AUTHOR
    Created on: April 17, 2024 (Updated April 17, 2025)
#>

#Requires -Modules ActiveDirectory, AzureAD

param(
    [Parameter(Mandatory=$true)]
    [string]$CsvPath
)

#Region Functions

# Funktion zum Abrufen der AzureAD-Benutzerinformationen (using AzureAD module)
function Get-EntraUserInfo {
    param(
        [Parameter(Mandatory=$true)]
        [string]$UserPrincipalName
    )
    try {
        # Using the older AzureAD module here. Consider migrating to Microsoft.Graph (`Get-MgUser`).
        $azureADUser = Get-AzureADUser -Filter "userPrincipalName eq '$UserPrincipalName'" -ErrorAction Stop
        if ($azureADUser) {
            Write-Verbose "Successfully retrieved Azure AD user: $UserPrincipalName"
            return $azureADUser
        } else {
            # Get-AzureADUser throws an error if not found when -ErrorAction Stop is used, so this might not be reached.
            # If -ErrorAction SilentlyContinue was used, this would be relevant.
            Write-Warning "User $UserPrincipalName not found in Azure AD."
            return $null
        }
    } catch {
        Write-Warning "Error retrieving Azure AD user $UserPrincipalName: $($_.Exception.Message)"
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
    $exportedManager = $null
    try {
        # Sicherstellen, dass der Ausgabepfad existiert
        if (-not (Test-Path -Path $OutputPath)) {
            New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
        }

        # Dateinamen sicher gestalten (ersetzt ungültige Zeichen)
        $safeUpn = $AzureADUser.UserPrincipalName -replace '[\\/:*?"<>|]', '_'
        
        # Exportieren der Benutzerinformationen
        $xmlUserPath = Join-Path -Path $OutputPath -ChildPath "$($safeUpn).user.xml"
        $AzureADUser | Export-Clixml -Path $xmlUserPath
        Write-Verbose "Exported user details for $($AzureADUser.UserPrincipalName) to $xmlUserPath"

        # Abrufen und Exportieren der Manager-Informationen
        try {
            # Using the older AzureAD module here. Consider migrating to Microsoft.Graph (`Get-MgUserManager`).
            $manager = Get-AzureADUserManager -ObjectId $AzureADUser.ObjectId -ErrorAction SilentlyContinue # Use SilentlyContinue to handle "not found" gracefully
            if ($manager) {
                $xmlManagerPath = Join-Path -Path $OutputPath -ChildPath "$($safeUpn).manager.xml"
                $manager | Export-Clixml -Path $xmlManagerPath
                Write-Verbose "Exported manager details for $($AzureADUser.UserPrincipalName) to $xmlManagerPath"
                $exportedManager = $manager
            } else {
                 Write-Host "  No manager found in Azure AD for user $($AzureADUser.UserPrincipalName)" -ForegroundColor Gray
            }
        } catch {
            # Catches errors other than "manager not found"
            Write-Warning "  Error retrieving Azure AD Manager for $($AzureADUser.UserPrincipalName): $($_.Exception.Message)"
        }       
        
        return $exportedManager # Return the manager object if found, otherwise null
    } catch {
        Write-Warning "Error during XML export process for $($AzureADUser.UserPrincipalName): $($_.Exception.Message)"
        return $exportedManager # Return manager if already retrieved, otherwise null
    }
}

# Funktion zum Aktualisieren des lokalen AD-Benutzers mit AzureAD-Attributen
function Set-LocalADUserAttributes {
    param(
        [Parameter(Mandatory=$true)]
        $AzureADUser,
        [Parameter(Mandatory=$false)]
        $AzureADManager, # Manager object from Azure AD
        [Parameter(Mandatory=$true)]
        [string]$CsvCountry # Country name from CSV
    )
    try {
        # Lokalen AD-Benutzer anhand des UPN finden
        $localADUser = Get-ADUser -Filter "UserPrincipalName -eq '$($AzureADUser.UserPrincipalName)'" -Properties * -ErrorAction Stop # Stop if local user not found

        if ($localADUser) {
            # Hashtable für zu aktualisierende Attribute erstellen
            # Using -Splatter prevents errors if an attribute is $null
            $UpdateAttributes = @{} 
            
            # Attribute aus AzureAD übernehmen, wenn sie vorhanden sind und sich vom lokalen Wert unterscheiden
            if ($AzureADUser.TelephoneNumber -and $AzureADUser.TelephoneNumber -ne $localADUser.OfficePhone) { $UpdateAttributes.Add("OfficePhone", $AzureADUser.TelephoneNumber) }
            if ($AzureADUser.StreetAddress -and $AzureADUser.StreetAddress -ne $localADUser.StreetAddress) { $UpdateAttributes.Add("StreetAddress", $AzureADUser.StreetAddress) }
            if ($AzureADUser.PostalCode -and $AzureADUser.PostalCode -ne $localADUser.PostalCode) { $UpdateAttributes.Add("PostalCode", $AzureADUser.PostalCode) }
            if ($AzureADUser.JobTitle -and $AzureADUser.JobTitle -ne $localADUser.Title) { $UpdateAttributes.Add("Title", $AzureADUser.JobTitle) }
            if ($AzureADUser.Department -and $AzureADUser.Department -ne $localADUser.Department) { $UpdateAttributes.Add("Department", $AzureADUser.Department) }
            if ($AzureADUser.City -and $AzureADUser.City -ne $localADUser.City) { $UpdateAttributes.Add("l", $AzureADUser.City) } # AD attribute for City is 'l' (lowercase L)
            if ($AzureADUser.CompanyName -and $AzureADUser.CompanyName -ne $localADUser.Company) { $UpdateAttributes.Add("Company", $AzureADUser.CompanyName) }

            # Ländercode konvertieren (ISO 3166-1 alpha-2) und zum Hashtable hinzufügen, wenn er sich unterscheidet
            # Zielattribut ist hier 'c' (ISO Code). Wenn 'co' (Name) gewünscht ist, dies anpassen.
            $countryCode = switch ($CsvCountry) {
                "Netherlands"    { "NL" }
                "United Kingdom" { "GB" }
                "Germany"        { "DE" } # Beispiel hinzugefügt
                # Fügen Sie hier weitere Länder hinzu
                default          { Write-Warning "  Country '$CsvCountry' not mapped to ISO code, skipping country update."; $null } 
            }
            
            if ($countryCode -and $countryCode -ne $localADUser.c) {
                 $UpdateAttributes.Add("c", $countryCode)
            } elseif (!$countryCode -and $localADUser.c) {
                # Optional: Clear local country if no mapping found? Or leave as is?
                # $UpdateAttributes.Add("c", $null) 
            }

            # Manager setzen, falls ein Manager in Azure AD gefunden wurde
            $managerDN = $null # Distinguished Name of the manager in local AD
            if ($AzureADManager) {
                try {
                    $managerLocalAD = Get-ADUser -Filter "UserPrincipalName -eq '$($AzureADManager.UserPrincipalName)'" -ErrorAction Stop
                    if ($managerLocalAD) {
                         $managerDN = $managerLocalAD.DistinguishedName
                    } else {
                        Write-Warning "  Manager with UPN $($AzureADManager.UserPrincipalName) found in Azure AD but NOT found in local AD. Cannot set manager."
                    }
                } catch {
                     Write-Warning "  Error finding local AD manager with UPN $($AzureADManager.UserPrincipalName): $($_.Exception.Message)"
                }
            } else {
                # If no manager in Azure AD, maybe clear the local manager? Depends on requirements.
                # $managerDN = $null # Ensure it's null if no Azure manager
            }

            # Update manager if the new DN is different from the current one
            if (($managerDN -ne $localADUser.Manager) -or ($managerDN -eq $null -and $localADUser.Manager -ne $null)) {
                 Write-Host "  Updating manager..." -ForegroundColor DarkCyan
                 try {
                     Set-ADUser -Identity $localADUser.SamAccountName -Manager $managerDN -ErrorAction Stop
                     Write-Host "    Set Manager to: $(if($managerDN){$managerDN}else{'$null (Cleared)'})" -ForegroundColor Cyan
                 } catch {
                      Write-Warning "    Failed to set manager for $($localADUser.SamAccountName): $($_.Exception.Message)"
                 }
            }

            # Andere Attribute nur aktualisieren, wenn Änderungen vorliegen
            if ($UpdateAttributes.Count -gt 0) {
                Write-Host "  Applying attribute updates..." -ForegroundColor DarkCyan
                try {
                     Set-ADUser -Identity $localADUser.SamAccountName -Replace $UpdateAttributes -ErrorAction Stop # Use -Replace for safety
                     Write-Host "    Updated attributes: $($UpdateAttributes.Keys -join ', ')" -ForegroundColor Cyan
                     # Optional: Output old/new values for logging
                     # $UpdateAttributes.GetEnumerator() | ForEach-Object { Write-Host "      $($_.Key): '$($localADUser.($_.Key))' -> '$($_.Value)'" }                    
                } catch {
                     Write-Warning "    Failed to update attributes for $($localADUser.SamAccountName): $($_.Exception.Message)"
                     return $false # Indicate failure on attribute update
                }
            } else {
                Write-Host "  No attribute changes detected (excluding manager)." -ForegroundColor Gray
            }

            # ImmutableID Berechnen (basierend auf ObjectGUID des LOKALEN AD-Benutzers)
            # Dieser Wert wird normalerweise von AAD Connect verwendet, um den lokalen Benutzer mit dem Azure AD Benutzer zu verknüpfen.
            # Das Skript SETZT diesen Wert NICHT in Azure AD.
            $immutableID = [System.Convert]::ToBase64String($localADUser.ObjectGUID.ToByteArray())
            Write-Host "  Local AD user '$($localADUser.SamAccountName)' processed. Calculated ImmutableID: $immutableID" -ForegroundColor Green
            Write-Host "  NOTE: ImmutableID is calculated but NOT set in Azure AD by this script." -ForegroundColor Yellow

            # === WARNUNG: ImmutableID in Azure AD setzen ===
            # Das Setzen der ImmutableID sollte nur in spezifischen Szenarien erfolgen (z.B. Konvertierung Cloud->Synced, Reparatur).
            # Führen Sie dies NICHT bei regulär synchronisierten Benutzern aus, da dies die Synchronisierung stören kann!
            # Beispiel (AUSKOMMENTIERT):
            # try {
            #    Write-Host "  Attempting to set ImmutableID in Azure AD (USE WITH EXTREME CAUTION)..."
            #    Set-AzureADUser -ObjectId $AzureADUser.ObjectId -ImmutableId $immutableID
            #    Write-Host "    Successfully set ImmutableID in Azure AD for $($AzureADUser.UserPrincipalName)." -ForegroundColor Green
            # } catch {
            #    Write-Warning "    FAILED to set ImmutableID in Azure AD for $($AzureADUser.UserPrincipalName): $($_.Exception.Message)"
            # }
            # === ENDE WARNUNG ===

            return $true # Indicate success
        } else {
            # Sollte wegen -ErrorAction Stop in Get-ADUser nicht erreicht werden, aber sicherheitshalber drinlassen.
            Write-Warning "Local AD user with UPN $($AzureADUser.UserPrincipalName) not found."
            return $false
        }
    } catch {
        # Catches errors from Get-ADUser or other unexpected issues within the function
         Write-Error "Error processing local AD user for UPN '$($AzureADUser.UserPrincipalName)': $($_.Exception.Message)"
        return $false
    }
}

#EndRegion Functions

#Region Main Script

# Überprüfen der AzureAD-Verbindung
try {
    $azureADConnection = Get-AzureADCurrentSessionInfo -ErrorAction Stop
    Write-Host "Successfully connected to Azure AD tenant: $($azureADConnection.TenantId) ($($azureADConnection.TenantDomain))" -ForegroundColor Green
} catch {
    Write-Error "Not connected to Azure AD. Please run Connect-AzureAD first."
    exit 1 # Use non-zero exit code for errors
}

# Überprüfen, ob die CSV-Datei existiert
if (-not (Test-Path -Path $CsvPath -PathType Leaf)) { # Leaf ensures it's a file
    Write-Error "The specified CSV file was not found or is not a file: $CsvPath"
    exit 1
}

# CSV-Datei einlesen
try {
    # Specify Delimiter if it's not a comma, e.g. -Delimiter ';'
    $users = Import-Csv -Path $CsvPath -ErrorAction Stop 
    if ($null -eq $users -or $users.Count -eq 0) {
         Write-Error "CSV file is empty or could not be read properly: $CsvPath"
         exit 1
    }
     # Check for required columns
    if (-not $users[0].PSObject.Properties.Name -contains 'UserPrincipalName' -or -not $users[0].PSObject.Properties.Name -contains 'Country') {
         Write-Error "CSV file must contain 'UserPrincipalName' and 'Country' columns."
         exit 1
    }
    $userCount = $users.Count
    Write-Host "Found $userCount users in CSV file '$CsvPath'." -ForegroundColor Cyan
} catch {
    Write-Error "Error reading CSV file '$CsvPath': $($_.Exception.Message)"
    exit 1
}

# Ausgabepfad für XML-Dateien (Unterordner 'Export' im Ordner der CSV)
$csvFolder = Split-Path -Path $CsvPath -Parent
$outputPath = Join-Path -Path $csvFolder -ChildPath "Export_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
Write-Host "XML export files will be saved to: $outputPath" -ForegroundColor Cyan

# Statistik-Variablen
$successCount = 0
$failureCount = 0
$processedCount = 0

# Verarbeitung der Benutzer
Write-Host "`nStarting synchronization process..." -ForegroundColor Yellow
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

foreach ($userRecord in $users) {
    $processedCount++
    $upn = $userRecord.UserPrincipalName.Trim() # Trim whitespace
    $country = $userRecord.Country.Trim()
    
    Write-Host "------------------------------------------------------"
    Write-Host "Processing user $processedCount/$userCount : $upn" -ForegroundColor Yellow

    if (-not $upn -or -not $country) {
        Write-Warning "  Skipping record due to missing UserPrincipalName or Country."
        $failureCount++
        continue # Skip to the next user
    }

    # 1. AzureAD-Benutzer abrufen
    $azureADUser = Get-EntraUserInfo -UserPrincipalName $upn
    
    if ($azureADUser) {
        # 2. Optional: XML-Dateien exportieren und Manager-Informationen abrufen
        $azureADManager = Export-UserManagerInfo -AzureADUser $azureADUser -OutputPath $outputPath
        
        if ($azureADManager) {
            Write-Host "  Azure AD Manager found: $($azureADManager.DisplayName) ($($azureADManager.UserPrincipalName))" -ForegroundColor Cyan
        }

        # 3. Lokalen AD-Benutzer aktualisieren
        $updateResult = Set-LocalADUserAttributes -AzureADUser $azureADUser -AzureADManager $azureADManager -CsvCountry $country
        
        if ($updateResult) {
            $successCount++
        } else {
            $failureCount++
            Write-Warning "  Failed to fully process local AD user for $upn." # More specific error in function
        }
    } else {
        # Azure AD user not found
        $failureCount++
        # No further processing possible for this user
    }
}

$stopwatch.Stop()
Write-Host "------------------------------------------------------"
Write-Host "`nSynchronization complete." -ForegroundColor Green
Write-Host "Duration: $($stopwatch.Elapsed.ToString('g'))"
Write-Host "Successfully processed: $successCount users" -ForegroundColor Green
if ($failureCount -gt 0) {
    Write-Host "Failed/Skipped records: $failureCount users" -ForegroundColor Red
    Write-Host "Check warnings and errors above for details." -ForegroundColor Red
}

# Optional: Azure AD Connect Delta Sync starten
$runSync = Read-Host "`nWould you like to start an Azure AD Connect Delta Sync cycle? (Y/N)"
if ($runSync -eq "Y" -or $runSync -eq "y") {
    Write-Host "Attempting to start Azure AD Connect Delta Sync..."
    try {
        # Ensure the ADSync module is available on the machine running the script (usually the AAD Connect server)
        Import-Module ADSync -ErrorAction Stop 
        Start-ADSyncSyncCycle -PolicyType Delta
        Write-Host "Azure AD Connect Delta synchronization cycle has been requested." -ForegroundColor Green
    } catch {
        Write-Error "Error starting delta synchronization: $($_.Exception.Message)"
        Write-Error "Ensure the ADSync module is installed and you are running this on the Azure AD Connect server or have appropriate permissions."
    }
}

#EndRegion Main Script
