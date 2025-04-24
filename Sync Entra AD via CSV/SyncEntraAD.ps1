<#
.SYNOPSIS
    Synchronizes Azure AD user data with local AD users and exports the information as XML.
.DESCRIPTION
    This script reads a CSV file containing Azure AD user information, retrieves the corresponding local AD user,
    updates local AD attributes based on Azure AD data, and exports both Azure AD and local AD user information to XML files.
.EXAMPLE
    .\EntraADSync.ps1 -CsvPath "C:\path\to\your\file.csv"
.PARAMETER CsvPath
    The path to the CSV file containing Azure AD user information. The CSV should have a header with "UserPrincipalName".
.NOTES
    Requirements:
      - Modules: AzureAD and ActiveDirectory must be installed and connected.
      - The script must be run as administrator.
.AUTHOR
    Stefan Weiß, April 17, 2025
#>

#region Variables
# Progress tracking variables
$script:totalUsers = 0
$script:currentUser = 0
$script:progressActivity = "Azure AD to Local AD Synchronization"
$script:progressStatus = "Initializing synchronization..."
#endregion Variables

<#region Check Azure AD Connection
try {
    $null = Get-AzureADTenantDetail -ErrorAction Stop
    Write-Verbose "Successfully verified Azure AD connection"
} catch {
    Write-Error "Not connected to Azure AD. Please run Connect-AzureAD first."
    Return
}
#endregion Check Azure AD Connection#>

param(
    [Parameter(Mandatory = $true)]
    [string]$CsvPath
)

# Check if the CSV file exists
if (-not (Test-Path -Path $CsvPath)) {
    Write-Error "CSV file not found: $CsvPath"
    Return
}

# Determine the folder path of the CSV file
$csvFolder = Split-Path -Path $CsvPath -Parent

# Read the CSV file
try {
    $csvUsers = Import-Csv -Path $CsvPath -ErrorAction Stop
} catch {
    $msg = "Failed to read CSV file: {0}" -f $_.Exception.Message
    Write-Error $msg
    Return
}

# Initialize progress tracking
$script:totalUsers = $csvUsers.Count
$script:currentUser = 0

#region Process CSV Records
foreach ($csvUser in $csvUsers) {
    $script:currentUser++
    $percentComplete = ($script:currentUser / $script:totalUsers) * 100

    Write-Progress -Activity $script:progressActivity |
        -Status "Processing $script:currentUser of $script:totalUsers users" |
        -PercentComplete $percentComplete |
        -CurrentOperation "Current user: $($csvUser.UserPrincipalName)"

    $upn = $csvUser.UserPrincipalName
    if (-not $upn) {
        Write-Verbose "Skipping record with missing UserPrincipalName."
        continue
    }
    Write-Progress "Processing user: $upn" -ForegroundColor Cyan

    # Retrieve the Azure AD user
    $azureUser = Get-AzureADUser -Filter "UserPrincipalName eq '$upn'" -ErrorAction SilentlyContinue
    if ($null -eq $azureUser) {
        Write-Verbose "Azure AD user not found for $upn"
        continue
    }

    # Generate a safe filename (replace invalid characters)
    $safeUpn = $azureUser.UserPrincipalName -replace '[\\/:*?"<>|]', '_'

    # Export Azure AD user information
    $userXmlFileName = "{0}.user.xml" -f $safeUpn
    $userXmlFile = Join-Path -Path $csvFolder -ChildPath $userXmlFileName
    $azureUser | Export-Clixml -Path $userXmlFile
    Write-Verbose "Exported user data to: $userXmlFile" -ForegroundColor Green

    # Retrieve manager information
    $azureManager = Get-AzureADUserManager -ObjectId $azureUser.ObjectId -ErrorAction SilentlyContinue
    if ($null -ne $azureManager) {
        # Check if manager file already exists
        $managerXmlFileName = "{0}.manager.xml" -f $safeUpn
        $managerXmlFile = Join-Path -Path $csvFolder -ChildPath $managerXmlFileName

        if (Test-Path -Path $managerXmlFile) {
            Write-Verbose "Manager data file already exists: $managerXmlFile" -ForegroundColor Yellow
            continue
        }

        # Export manager information
        $azureManager | Export-Clixml -Path $managerXmlFile
        Write-Verbose "Exported manager data to: $managerXmlFile" -ForegroundColor Green
    }

    # Query the local AD user using the UPN
    $localAdUser = Get-ADUser -Filter "UserPrincipalName -eq '$($azureUser.UserPrincipalName)'" -ErrorAction SilentlyContinue
    if ($null -eq $localAdUser) {
        Write-Verbose "Local AD user not found for $($azureUser.UserPrincipalName)"
        continue
    }
    try {
        Set-ADUser -Identity $localAdUser.DistinguishedName -OfficePhone $azureUser.TelephoneNumber
        Set-ADUser -Identity $localAdUser.DistinguishedName -StreetAddress $azureUser.StreetAddress
        Set-ADUser -Identity $localAdUser.DistinguishedName -PostalCode $azureUser.PostalCode
        Set-ADUser -Identity $localAdUser.DistinguishedName -Title $azureUser.JobTitle
        Set-ADUser -Identity $localAdUser.DistinguishedName -Department $azureUser.Department
        Set-ADUser -Identity $localAdUser.DistinguishedName -City $azureUser.City
        Set-ADUser -Identity $localAdUser.DistinguishedName -Company $azureUser.CompanyName
        Set-ADUser -Identity $localAdUser.DistinguishedName -Country $azureUser.Country

        Write-Output "Updated local AD attributes for $($azureUser.UserPrincipalName)" -ForegroundColor Green
    } catch {
        $msgADUser = "Failed to update local AD attributes for {0}: {1}" -f $azureUser.UserPrincipalName, $_.Exception.Message
        Write-Error $msgADUser
    }

    # Compute and display the ImmutableID
    try {
        $refreshedUser = Get-ADUser -Identity $localAdUser.ObjectGUID -Properties ObjectGUID
        $immutableId = [System.Convert]::ToBase64String($refreshedUser.ObjectGUID.ToByteArray())
        Write-Progress "INFO: Computed ImmutableID (based on ObjectGUID) for '$upn': $immutableId"
        Set-AzureADUser -ObjectId $azureUser.ObjectId -ImmutableId $immutableId
        Write-Verbose "Computed ImmutableID for $($azureUser.UserPrincipalName): $immutableId" -ForegroundColor Green
    } catch {
        $msgImmutableId = "Failed to compute ImmutableID for {0}: {1}" -f $azureUser.UserPrincipalName, $_.Exception.Message
        Write-Error $msgImmutableId
    }

    # Set ExtensionAttribute10 to "365"
    try {
        Set-ADUser -Identity $localAdUser.DistinguishedName -Replace @{extensionAttribute10 = "365"}
        Write-Verbose "Set extensionAttribute10 to 365 for $($azureUser.UserPrincipalName)" -ForegroundColor Green
    } catch {
        $msg = "Failed to set extensionAttribute10 for {0}: {1}" -f $azureUser.UserPrincipalName, $_.Exception.Message
        Write-Error $msg
    }

    # Set the local AD manager, if available
    if ($null -ne $azureManager) {
        try {
            $localManager = Get-ADUser -Filter "UserPrincipalName -eq '$($azureManager.UserPrincipalName)'" -ErrorAction Stop
            Set-ADUser -Identity $localAdUser.DistinguishedName -Manager $localManager.DistinguishedName
            Write-Verbose "Set local AD manager for $($azureUser.UserPrincipalName)" -ForegroundColor Green
        } catch {
            $msgManager = "Failed to set local AD manager for {0}: {1}" -f $azureUser.UserPrincipalName, $_.Exception.Message
            Write-Error $msgManager
        }
    }
}

Write-Progress -Activity $script:progressActivity -Completed
#endregion Process CSV Records

#region Final Information
Write-Output "Execute an Entra Sync Delta now by running the appropriate command." -ForegroundColor Cyan
Write-Output "Note: After the Sync Delta, remember to activate the remote mailboxes." -ForegroundColor Cyan
#endregion Final Information
