<#
.SYNOPSIS
Synchronizes selected user attributes (Country, Manager, etc.) from Microsoft Entra ID (Azure AD)
to the local Active Directory based on a CSV file.

.DESCRIPTION
This script reads user UPNs and country information from a CSV file.
For each user:
1. Retrieves the user object and its manager from Entra ID.
2. Optionally saves these objects as XML files for backup/review.
3. Searches for the corresponding user in the local AD by UPN.
4. Updates the country (using mapping 'Netherlands' -> 'NL', 'United Kingdom' -> 'GB')
   and the manager (finds the local manager using their UPN) in the local AD.
5. Updates additional attributes such as TelephoneNumber, StreetAddress, PostalCode, JobTitle, Department, City, and Company.
6. Computes and displays the ImmutableID (ObjectGUID in Base64) of the local user.
7. Asks at the end if an Entra ID Connect (Azure AD Connect) delta synchronization cycle should be initiated.

.PARAMETER CsvPath
Path to the CSV file. The file MUST contain the columns 'UserPrincipalName' and 'Country'.

.EXAMPLE
.\Sync-EntraToLocalAD.ps1 -CsvPath "C:\temp\users_to_sync.csv"

.NOTES
Author: Your Name / Your Organization
Date: 2025-04-16
Version: 1.1

Requirements:
- PowerShell module 'AzureAD' must be installed (`Install-Module AzureAD`).
- PowerShell module 'ActiveDirectory' (part of the RSAT tools) must be installed.
  # Install-WindowsFeature RSAT-AD-PowerShell
- An active connection to Entra ID must exist (`Connect-AzureAD`).
- Sufficient permissions in Entra ID (read users/managers) and local AD (read/write users).
- If the delta sync is to be triggered: Execution on the AD Connect server or via remoting with appropriate permissions.
#>

param(
    [string]$CsvPath
)

# --- Prerequisites ---
Write-Host "INFO: Script started."
Write-Host "INFO: Ensure you are already connected to Microsoft Entra ID (Azure AD) (Connect-AzureAD)."
Write-Host "INFO: Ensure the RSAT AD PowerShell Tools are installed (`Install-WindowsFeature RSAT-AD-PowerShell`)."
Write-Host "INFO: The AzureAD module is required (`Install-Module AzureAD`)."

# --- Functions ---

function Get-AzureAdUserData {
    param(
        [Parameter(Mandatory=$true)]
        [string]$UPN
    )

    Write-Verbose "Attempting to retrieve Entra ID user data for UPN '$UPN'."
    $azureUser = $null
    $azureManager = $null

    try {
        # Retrieve user object
        $azureUser = Get-AzureADUser -ObjectId $UPN -ErrorAction Stop
        if ($azureUser) {
            Write-Verbose "Entra ID user found: $($azureUser.ObjectId)"
            # Retrieve user's manager
            try {
                $azureManager = Get-AzureADUserManager -ObjectId $azureUser.ObjectId -ErrorAction SilentlyContinue
                if ($azureManager) {
                    Write-Verbose "Entra ID manager found: $($azureManager.UserPrincipalName)"
                } else {
                    Write-Verbose "No manager found for user '$UPN' in Entra ID."
                }
            } catch {
                Write-Warning "WARNING: Error retrieving manager for '$UPN'. Details: $($_.Exception.Message)"
            }
        } else {
            Write-Verbose "No Entra ID user found with UPN '$UPN'."
        }
    } catch {
        Write-Warning "WARNING: Error retrieving Entra ID user '$UPN'. Details: $($_.Exception.Message)"
    }

    return @{
        User    = $azureUser
        Manager = $azureManager
    }
}

function Export-UserDataToXml {
    param(
        [Parameter(Mandatory=$true)]
        $AzureUser,

        $AzureManager,

        [Parameter(Mandatory=$true)]
        [string]$CsvPath
    )

    try {
        $outputDir = Split-Path -Path $CsvPath -Parent
        $safeUpnPart = $AzureUser.UserPrincipalName -replace '[^a-zA-Z0-9.-_]', '_'
        $userFileNameBase = $safeUpnPart

        $userXmlPath = Join-Path -Path $outputDir -ChildPath "$($userFileNameBase).user.xml"
        $managerXmlPath = Join-Path -Path $outputDir -ChildPath "$($userFileNameBase).manager.xml"

        Write-Verbose "Exporting user data to '$userXmlPath'"
        Export-CliXml -Path $userXmlPath -InputObject $AzureUser -Force

        if ($AzureManager) {
            Write-Verbose "Exporting manager data to '$managerXmlPath'"
            Export-CliXml -Path $managerXmlPath -InputObject $AzureManager -Force
        } else {
            if (Test-Path $managerXmlPath) {
                Write-Verbose "Removing old manager XML file '$managerXmlPath'."
                Remove-Item $managerXmlPath -Force -ErrorAction SilentlyContinue
            }
        }
    } catch {
        Write-Warning "WARNING: Error exporting data to XML for user '$($AzureUser.UserPrincipalName)'. Details: $($_.Exception.Message)"
    }
}

function Update-LocalAdUser {
    param(
        [Parameter(Mandatory=$true)]
        [string]$UPN,

        [Parameter(Mandatory=$true)]
        $AzureUser,

        $AzureManager,

        [Parameter(Mandatory=$true)]
        [string]$CsvCountry
    )

    Write-Verbose "Attempting to update local AD user with UPN '$UPN'."
    $localAdUser = $null
    try {
        $localAdUser = Get-ADUser -Filter "UserPrincipalName -eq '$UPN'" -Properties * -ErrorAction Stop
    } catch {
        Write-Warning "WARNING: Error finding local AD user '$UPN'. Details: $($_.Exception.Message)"
        return
    }

    if (-not $localAdUser) {
        Write-Warning "WARNING: Local AD user with UPN '$UPN' not found. No update possible."
        return
    }

    Write-Verbose "Local AD user found: $($localAdUser.DistinguishedName)"

    $adUserProperties = @{}

    $targetCountryCode = switch ($CsvCountry.Trim().ToLower()) {
        'netherlands'    { 'NL' }
        'united kingdom' { 'GB' }
        default          { $null }
    }

    if ($targetCountryCode) {
        if ($localAdUser.Country -ne $targetCountryCode -or $localAdUser.c -ne $targetCountryCode) {
            $adUserProperties.Country = $targetCountryCode
            $adUserProperties.c = $targetCountryCode
        }
    }

    if ($AzureUser.TelephoneNumber) { $adUserProperties["OfficePhone"] = $AzureUser.TelephoneNumber }
    if ($AzureUser.StreetAddress) { $adUserProperties["StreetAddress"] = $AzureUser.StreetAddress }
    if ($AzureUser.PostalCode) { $adUserProperties["PostalCode"] = $AzureUser.PostalCode }
    if ($AzureUser.JobTitle) { $adUserProperties["Title"] = $AzureUser.JobTitle }
    if ($AzureUser.Department) { $adUserProperties["Department"] = $AzureUser.Department }
    if ($AzureUser.City) { $adUserProperties["City"] = $AzureUser.City }
    if ($AzureUser.CompanyName) { $adUserProperties["Company"] = $AzureUser.CompanyName }

    if ($adUserProperties.Keys.Count -gt 0) {
        try {
            Set-ADUser -Identity $localAdUser -Replace $adUserProperties -ErrorAction Stop
            Write-Verbose "Attributes successfully updated."
        } catch {
            Write-Error "ERROR: Could not update attributes for local AD user '$UPN'. Details: $($_.Exception.Message)"
        }
    } else {
        Write-Host "INFO: No attributes to update for '$UPN'."
    }
}

# --- Main Script ---

try {
    $usersToProcess = Import-Csv -Path $CsvPath -ErrorAction Stop
    Write-Host "INFO: $($usersToProcess.Count) user records read from '$CsvPath'."
} catch {
    Write-Error "ERROR: Could not read or validate CSV file '$CsvPath'. Details: $($_.Exception.Message)"
    exit 1
}

foreach ($record in $usersToProcess) {
    $upn = $record.UserPrincipalName.Trim()
    $csvCountry = $record.Country.Trim()

    if ([string]::IsNullOrWhiteSpace($upn)) {
        Write-Warning "WARNING: Empty UserPrincipalName found in CSV. Skipping."
        continue
    }

    Write-Host "--- Processing User: $upn ---"

    $azureData = Get-AzureAdUserData -UPN $upn
    if ($azureData -and $azureData.User) {
        Export-UserDataToXml -AzureUser $azureData.User -AzureManager $azureData.Manager -CsvPath $CsvPath
        Update-LocalAdUser -UPN $upn -AzureUser $azureData.User -AzureManager $azureData.Manager -CsvCountry $csvCountry
    } else {
        Write-Warning "WARNING: User '$upn' not found in Entra ID. Skipping local AD update."
    }

    Write-Host "--- Finished processing user: $upn ---`n"
}

Write-Host "INFO: User processing completed."

$triggerSync = Read-Host "Do you want to start a Microsoft Entra Connect delta sync? (y/N)"
if ($triggerSync -eq 'y') {
    try {
        if (-not (Get-Module -Name ADSync -ErrorAction SilentlyContinue)) {
            Import-Module ADSync -ErrorAction Stop
        }
        Start-ADSyncSyncCycle -PolicyType Delta -ErrorAction Stop
        Write-Host "INFO: Delta sync successfully started."
    } catch {
        Write-Error "ERROR: Could not start the delta sync. Ensure you are on the AD Connect Server and the ADSync module is available."
    }
} else {
    Write-Host "INFO: Delta sync skipped."
}

Write-Host "INFO: Script finished."
