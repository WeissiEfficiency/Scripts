<#
.SYNOPSIS
    Synchronizes attributes from Azure AD (Entra ID) users to local AD users and sets the ImmutableId in Azure AD.

.DESCRIPTION
    This script synchronizes selected attributes from Azure AD (Entra ID) users to on-premises Active Directory users.
    It reads a CSV file containing Azure AD UserPrincipalNames and Country information.
    For each user, it retrieves the Azure AD user and manager information using the AzureAD module.
    It finds the corresponding local AD user by UPN.
    It **FORCE OVERWRITES** specified local AD attributes with values from Azure AD (null values from Azure AD will clear local attributes).
    It finds the local AD object for the manager (based on the Azure AD manager's UPN) and sets the manager link in the local AD.
    ***CRITICAL STEP***: It calculates the ImmutableID from the local AD user's ObjectGUID and attempts to SET this ImmutableId on the corresponding AZURE AD USER OBJECT.

.PARAMETER CsvPath
    Path to the CSV file containing user information. 
    The CSV must include at least 'UserPrincipalName' and 'Country' columns.

.EXAMPLE
    .\Sync-AzureADToLocalAD_SetImmutableID.ps1 -CsvPath "C:\Temp\users.csv"

.NOTES
    Prerequisites:
    - AzureAD PowerShell module must be installed (`Install-Module AzureAD`). 
      Note: Microsoft Graph PowerShell SDK (`Microsoft.Graph`) is the modern replacement.
    - **RSAT AD PowerShell tools must be installed** (`Install-WindowsFeature RSAT-AD-PowerShell` or via Settings > Optional features).
    - Must be connected to Azure AD using `Connect-AzureAD` before running this script.
    - Must have appropriate permissions in both Azure AD and local AD.
    - Review attribute mappings (especially country codes and target attributes like 'c' vs 'co') and adjust if needed.
    - ***EXTREME CAUTION*** when running this script due to the setting of ImmutableId in Azure AD. Understand the implications for Azure AD Connect synchronization. Use only if you need to prepare users for sync matching or fix specific sync issues. DO NOT run on already correctly synchronized users.

.AUTHOR
    Created on: April 17, 2024 (Updated April 17, 2025 for ImmutableID set) 
#>

#Requires -Modules ActiveDirectory, AzureAD, RSAT.ActiveDirectory

param(
    [Parameter(Mandatory=$true)]
    [string]$CsvPath
)

#Region Functions

# Function to retrieve Azure AD user information (using AzureAD module)
function Get-AzureAdUserInfoForSync {
    param(
        [Parameter(Mandatory=$true)]
        [string]$UserPrincipalName
    )
    try {
        # Using the older AzureAD module here. Consider migrating to Microsoft.Graph (`Get-MgUser`).
        $azureUser = Get-AzureADUser -Filter "userPrincipalName eq '$UserPrincipalName'" -ErrorAction Stop
        Write-Verbose "Successfully retrieved Azure AD user: $UserPrincipalName"
        return $azureUser
    } catch {
        Write-Warning "Error retrieving Azure AD user $UserPrincipalName: $($_.Exception.Message)"
        return $null
    }
}

# Function to export user and manager information to XML files (optional logging/backup)
function Export-UserAndManagerInfoToXml {
    param(
        [Parameter(Mandatory=$true)]
        $AzureUser, # Azure AD User object passed in
        [Parameter(Mandatory=$true)]
        [string]$OutputPath # Folder path for XML files
    )
    $azureManager = $null # Initialize manager variable
    try {
        # Ensure the output path exists
        if (-not (Test-Path -Path $OutputPath)) {
            New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
        }

        # Create a safe filename from the UPN (replace invalid characters)
        $safeUpn = $AzureUser.UserPrincipalName -replace '[\\/:*?"<>|]', '_'
        
        # Export user information
        $xmlUserPath = Join-Path -Path $OutputPath -ChildPath "$($safeUpn).user.xml"
        $AzureUser | Export-Clixml -Path $xmlUserPath
        Write-Verbose "Exported user details for $($AzureUser.UserPrincipalName) to $xmlUserPath"

        # Retrieve and export manager information
        try {
            # Using the older AzureAD module here. Consider migrating to Microsoft.Graph (`Get-MgUserManager`).
            $manager = Get-AzureADUserManager -ObjectId $AzureUser.ObjectId -ErrorAction SilentlyContinue # Use SilentlyContinue to handle "not found" gracefully
            if ($manager) {
                $xmlManagerPath = Join-Path -Path $OutputPath -ChildPath "$($safeUpn).manager.xml"
                $manager | Export-Clixml -Path $xmlManagerPath
                Write-Verbose "Exported manager details for $($AzureUser.UserPrincipalName) to $xmlManagerPath"
                $azureManager = $manager # Assign found manager to return variable
            } else {
                 Write-Host "  No manager found in Azure AD for user $($AzureUser.UserPrincipalName)" -ForegroundColor Gray
            }
        } catch {
            # Catches errors other than "manager not found" during manager retrieval/export
            Write-Warning "  Error retrieving/exporting Azure AD Manager for $($AzureUser.UserPrincipalName): $($_.Exception.Message)"
        }       
        
        return $azureManager # Return the Azure AD manager object if found, otherwise null
    } catch {
        # Catches errors in the outer block (e.g., path issues, user export)
        Write-Warning "Error during XML export process for $($AzureUser.UserPrincipalName): $($_.Exception.Message)"
        return $azureManager # Return manager if already retrieved, otherwise null
    }
}

# Function to update the local AD user and set the ImmutableId in Azure AD
function Set-LocalADUserAndAzureImmutableId {
    param(
        [Parameter(Mandatory=$true)]
        $AzureUser, # Azure AD User Object from Get-AzureAdUserInfoForSync
        [Parameter(Mandatory=$false)]
        $AzureManager, # Azure AD Manager Object from Export-UserAndManagerInfoToXml
        [Parameter(Mandatory=$true)]
        [string]$CsvCountry # Country name from the CSV file
    )
    $localUser = $null # Initialize local user variable
    try {
        # 1. Find the local AD user based on the UPN
        $localUser = Get-ADUser -Filter "UserPrincipalName -eq '$($AzureUser.UserPrincipalName)'" -Properties * -ErrorAction Stop # Stop if local user not found

        Write-Host "  Found local AD user: $($localUser.SamAccountName)" -ForegroundColor Cyan

        # 2. Create hashtable for attributes to be OVERWRITTEN in local AD
        # Values are taken directly from Azure AD. No comparison with local values.
        $attributesToSet = @{
            OfficePhone   = $AzureUser.TelephoneNumber # Null from Azure clears local
            StreetAddress = $AzureUser.StreetAddress
            PostalCode    = $AzureUser.PostalCode
            Title         = $AzureUser.JobTitle
            Department    = $AzureUser.Department
            l             = $AzureUser.City # AD attribute for City is 'l' (lowercase L)
            Company       = $AzureUser.CompanyName
        }

        # Replace the full country name with the abbreviation
        $countryCode = switch ($CsvCountry) {
            "Netherlands"    { "NL" }
            "United Kingdom" { "GB" }
            "Germany"        { "DE" } 
            # Add more country mappings here as needed
            default          { Write-Warning "  Country '$CsvCountry' not mapped to ISO code."; $null } 
        }
        # Add country code to attributes to be set, even if null (to potentially clear it)
        $attributesToSet.Add("c", $countryCode)

        # 3. Prepare the local AD manager (Lookup is REQUIRED for Set-ADUser -Manager)
        $managerDistinguishedName = $null # DN of the manager in local AD
        if ($AzureManager) {
            Write-Host "  Azure AD Manager is $($AzureManager.DisplayName). Attempting to find corresponding local AD object..." -ForegroundColor DarkCyan
            try {
                # Find the local AD manager object using the UPN from the Azure AD manager object
                $localManager = Get-ADUser -Filter "UserPrincipalName -eq '$($AzureManager.UserPrincipalName)'" -ErrorAction Stop
                if ($localManager) {
                     $managerDistinguishedName = $localManager.DistinguishedName
                     Write-Host "    Found local manager: $($localManager.SamAccountName)" -ForegroundColor Cyan
                } 
                # ErrorAction Stop handles the "not found" case by throwing an error
            } catch {
                 Write-Warning "    Local AD Manager with UPN '$($AzureManager.UserPrincipalName)' NOT FOUND. Cannot set manager link. Error: $($_.Exception.Message)"
                 # Manager remains $null
            }
        } else {
             Write-Host "  No manager specified in Azure AD. Clearing local manager (if set)." -ForegroundColor DarkCyan
             $managerDistinguishedName = $null # Ensure local manager is cleared if no Azure manager
        }
        
        # 4. Update the local AD user (Attributes and Manager)
        try {
             Write-Host "  Updating local AD user '$($localUser.SamAccountName)'..." -ForegroundColor DarkCyan
             # Set attributes using the prepared hashtable
             Set-ADUser -Identity $localUser.SamAccountName -Replace $attributesToSet # -Replace overwrites attributes and clears them if value is $null

             # Set manager using the found DN (or $null to clear)
             Set-ADUser -Identity $localUser.SamAccountName -Manager $managerDistinguishedName 
             
             Write-Host "    Local AD attributes updated/overwritten." -ForegroundColor Green
             Write-Host "    Local AD Manager set to: $(if($managerDistinguishedName){$managerDistinguishedName}else{'$null (Cleared)'})" -ForegroundColor Green

        } catch {
            Write-Warning "  FAILED to update local AD user $($localUser.SamAccountName): $($_.Exception.Message)"
            return $false # Indicate failure for this user processing step
        }

        # 5. *** CRITICAL STEP: Set ImmutableId in Azure AD ***
        #    This step hard-links the local AD user to the Azure AD user.
        #    ONLY EXECUTE IF YOU FULLY UNDERSTAND THE IMPLICATIONS!
        try {
            # Calculate the ImmutableId from the ObjectGUID of the LOCAL AD user
            $immutableIdValue = [System.Convert]::ToBase64String($localUser.ObjectGUID.ToByteArray())
            Write-Host "  Calculated ImmutableID from local ObjectGUID: $immutableIdValue" -ForegroundColor Cyan
            
        #    Write-Host "  Attempting to set ImmutableID on AZURE AD User '$($AzureUser.UserPrincipalName)'..." -ForegroundColor Yellow
        #   Write-Host "  ---> WARNING <--- This permanently links the Azure AD user to the local AD user using this value." -ForegroundColor Yellow
        #   Write-Host "  ---> WARNING <--- Do NOT do this for users already correctly synchronized by Azure AD Connect!" -ForegroundColor Yellow
            
            # Using the older AzureAD module command
            Set-AzureADUser -ObjectId $AzureUser.ObjectId -ImmutableId $immutableIdValue -ErrorAction Stop
            # For Microsoft.Graph module, the equivalent would be:
            # Update-MgUser -UserId $AzureUser.Id -OnPremisesImmutableId $immutableIdValue
            
            Write-Host "    Successfully SET ImmutableID in Azure AD for '$($AzureUser.UserPrincipalName)'." -ForegroundColor Green

        } catch {
            Write-Error "    *** FAILED TO SET IMMUTABLEID IN AZURE AD for '$($AzureUser.UserPrincipalName)' *** Error: $($_.Exception.Message)"
            Write-Error "    The Azure AD user might already have an ImmutableID or other issues occurred."
            # Optional: Decide if failure to set ImmutableId constitutes overall failure for the user processing
            # return $false 
        }

        return $true # Indicate overall success (or partial success if ImmutableID failed but local AD updated)

    } catch {
        # Catches errors from the initial Get-ADUser or other unexpected issues within the function
        if ($localUser -eq $null) {
             # This condition might be redundant if Get-ADUser uses -ErrorAction Stop, but keeps the logic clear
             Write-Error "Error: Local AD user with UPN '$($AzureUser.UserPrincipalName)' not found."
        } else {
             # Catch other potential errors during processing
             Write-Error "Error processing user '$($AzureUser.UserPrincipalName)' / '$($localUser.SamAccountName)': $($_.Exception.Message)"
        }
        return $false # Indicate failure
    }
}

#EndRegion Functions

#Region Main Script Execution

<# Check Azure AD connection
try {
    $azureADConnection = Get-AzureADCurrentSessionInfo -ErrorAction Stop
    Write-Host "Successfully connected to Azure AD tenant: $($azureADConnection.TenantId) ($($azureADConnection.TenantDomain))" -ForegroundColor Green
} catch {
    Write-Error "Not connected to Azure AD. Please run Connect-AzureAD first."
    exit 1 # Use non-zero exit code for errors
}
#>

<# Check if the CSV file exists
if (-not (Test-Path -Path $CsvPath -PathType Leaf)) { # Leaf ensures it's a file, not a directory
    Write-Error "The specified CSV file was not found or is not a file: $CsvPath"
    exit 1
}
#>

# Import the CSV file
try {
    # Specify Delimiter if it's not a comma, e.g. -Delimiter ';'
    $csvUsers = Import-Csv -Path $CsvPath -ErrorAction Stop 
    if ($null -eq $csvUsers -or $csvUsers.Count -eq 0) {
         Write-Error "CSV file is empty or could not be read properly: $CsvPath"
         exit 1
    }
     # Check for required columns ('UserPrincipalName' and 'Country')
    $csvColumns = $csvUsers[0].PSObject.Properties.Name
    if (-not $csvColumns -contains 'UserPrincipalName' -or -not $csvColumns -contains 'Country') {
         Write-Error "CSV file must contain 'UserPrincipalName' and 'Country' columns."
         exit 1
    }
    $userCount = $csvUsers.Count
    Write-Host "Found $userCount users in CSV file '$CsvPath'." -ForegroundColor Cyan
} catch {
    Write-Error "Error reading CSV file '$CsvPath': $($_.Exception.Message)"
    exit 1
}

# Set up output path for optional XML export files (subfolder named 'Export_YYYYMMDD_HHmmss' in the CSV's directory)
$csvFolder = Split-Path -Path $CsvPath -Parent
$xmlOutputPath = Join-Path -Path $csvFolder -ChildPath "Export_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
Write-Host "Optional XML export files will be saved to: $xmlOutputPath" -ForegroundColor Cyan

# Initialize statistics counters
$successCount = 0
$failureCount = 0
$processedCount = 0

# Start processing users
<#
Write-Host "`nStarting synchronization and ImmutableID set process..." -ForegroundColor Yellow
Write-Host "*** WARNING: This script will attempt to SET the ImmutableID in Azure AD for matched users. Review code and understand implications! ***" -ForegroundColor Red
Read-Host -Prompt "Press Enter to continue, or CTRL+C to abort"
#>

$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

foreach ($csvUserRecord in $csvUsers) {
    $processedCount++
    $upn = $csvUserRecord.UserPrincipalName#.Trim() # Trim whitespace from UPN
    $countryName = $csvUserRecord.Country#.Trim()       # Trim whitespace from Country
    
#    Write-Host "------------------------------------------------------"
#    Write-Host "Processing user $processedCount/$userCount : $upn" -ForegroundColor Yellow

    # Basic validation for the current record
    if (-not $upn -or -not $countryName) {
        Write-Warning "  Skipping record $processedCount due to missing UserPrincipalName or Country in CSV."
        $failureCount++
        continue # Skip to the next user in the loop
    }

    # 1. Retrieve Azure AD user information
    $azureADUser = Get-AzureAdUserInfoForSync -UserPrincipalName $upn
    
    if ($azureADUser) {
        #Retrieve Azure AD Manager info & Export User/Manager to XML
        $azureADManager = Export-UserAndManagerInfoToXml -AzureUser $azureADUser -OutputPath $xmlOutputPath
        
        #Update Local AD User Attributes AND Set ImmutableId in Azure AD
        $updateResult = Set-LocalADUserAndAzureImmutableId -AzureUser $azureADUser -AzureManager $azureADManager -CsvCountry $countryName
        
        if ($updateResult) {
            $successCount++
        } else {
            $failureCount++
            Write-Warning "  Failed to fully process user $upn. Check errors logged above." 
        }
    } else {
        # Azure AD user not found via Get-AzureAdUserInfoForSync
        $failureCount++
        Write-Warning "  Azure AD User $upn not found. Skipping this record."
        # No further processing possible for this user record
    }
} # End foreach user loop

<#
$stopwatch.Stop()
Write-Host "------------------------------------------------------"
Write-Host "`nProcess complete." -ForegroundColor Green
Write-Host "Duration: $($stopwatch.Elapsed.ToString('g'))"
Write-Host "Successfully processed records (incl. partial successes): $successCount" -ForegroundColor Green
if ($failureCount -gt 0) {
    Write-Host "Failed/Skipped records: $failureCount" -ForegroundColor Red
    Write-Host "Check warnings and errors logged above for details." -ForegroundColor Red
}
#>

# Optional: Start Azure AD Connect Delta Sync cycle
# Note: After manually setting ImmutableIDs, a Full Sync (Initial) might be more appropriate 
# to ensure everything aligns correctly, but Delta is offered here.
$runSync = Read-Host "`nWould you like to start an Azure AD Connect Delta Sync cycle? (Y/N)"
if ($runSync -eq "Y" -or $runSync -eq "y") {
    Write-Host "Attempting to start Azure AD Connect Delta Sync..."
    try {
        # Ensure the ADSync module is available on the machine running the script 
        # (typically the AAD Connect server itself or a machine with the tools installed)
        Import-Module ADSync -ErrorAction Stop 
        Start-ADSyncSyncCycle -PolicyType Delta
        Write-Host "Azure AD Connect Delta synchronization cycle has been requested." -ForegroundColor Green
        Write-Host "Consider running a Full Sync (Start-ADSyncSyncCycle -PolicyType Initial) if significant changes were made, especially involving ImmutableIDs." -ForegroundColor Yellow
    } catch {
        Write-Error "Error starting delta synchronization: $($_.Exception.Message)"
        Write-Error "Ensure the ADSync module is installed and you are running this script on the Azure AD Connect server or have appropriate permissions and remote execution configured."
    }
}

#EndRegion Main Script Execution