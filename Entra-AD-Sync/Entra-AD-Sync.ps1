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

#Requires -Modules ActiveDirectory, AzureAD

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
                $xmlManagerPath = Join-Path -Path <span class="math-inline">OutputPath \-ChildPath "</span>($safeUpn).manager.xml"
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
        <span class="math-inline">localUser \= Get\-ADUser \-Filter "UserPrincipalName \-eq '</span>($AzureUser.UserPrincipalName)'" -Properties * -ErrorAction Stop # Stop if local user not found

        Write-Host "  Found local AD user: $($localUser.SamAccountName)" -ForegroundColor Cyan

        # 2. Create hashtable for attributes to be OVERWRITTEN in local AD
        # Values are taken directly from Azure AD. No
