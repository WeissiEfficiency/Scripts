<#
.SYNOPSIS
    Synchronizes Azure AD user data with local AD users and exports the information as XML.
.DESCRIPTION
    This script reads a CSV file path which provides each record's UserPrincipalName (UPN).
    For each UPN, the corresponding Azure AD user is retrieved via Get-AzureADUser. The user information
    is exported to an XML file (named as [UserPrincipalName].user.xml) in the same folder as the CSV file,
    and if available, the manager information is exported to a separate XML file ([UserPrincipalName].manager.xml).
    Next, the local AD user is queried via the UPN, and the following attributes are updated from the Azure AD user:
      - OfficePhone, StreetAddress, PostalCode, JobTitle, Department, City, Company, Country.
    If a manager is found, the corresponding local AD account for the manager is set.
    In addition, the ImmutableID is computed based on the ObjectGUID and the ExtensionAttribute10 is set to "365".
    Finally, a message is displayed indicating that an Entra Sync Delta can be executed and that the remote mailboxes
    should be activated afterwards.
.NOTES
    Requirements:
      - Modules: AzureAD and ActiveDirectory must be installed and connected.
      - The script must be run as administrator.
.AUTHOR
    Your Name, April 17, 2025
#>

#region Parameters & CSV Import
param(
    [Parameter(Mandatory = $true)]
    [string]$CsvPath
)

# Check if the CSV file exists
if (-not (Test-Path -Path $CsvPath)) {
    Write-Error "CSV file not found: $CsvPath"
    exit 1
}

# Determine the folder path of the CSV file
$csvFolder = Split-Path -Path $CsvPath -Parent

# Read the CSV file
try {
    $csvUsers = Import-Csv -Path $CsvPath -ErrorAction Stop
} catch {
    Write-Error "Failed to read CSV file: $($_.Exception.Message)"
    exit 1
}
#endregion Parameters & CSV Import

#region Process CSV Records
foreach ($csvUser in $csvUsers) {
    $upn = $csvUser.UserPrincipalName
    if (-not $upn) {
        Write-Verbose "Skipping record with missing UserPrincipalName."
        continue
    }
    Write-Host "Processing user: $upn" -ForegroundColor Cyan

    # Retrieve the Azure AD user
    try {
        $azureUser = Get-AzureADUser -Filter "UserPrincipalName eq '$upn'" -ErrorAction Stop
    } catch {
        Write-Verbose "Could not retrieve Azure AD user for $upn"
        continue
    }

    if ($azureUser) {
        # Generate a safe filename (replace invalid characters)
        $safeUpn = $azureUser.UserPrincipalName -replace '[\\/:*?"<>|]', '_'
        
        # Export Azure AD user information
        $userXmlFile = Join-Path -Path $csvFolder -ChildPath "$safeUpn.user.xml"
        $azureUser | Export-Clixml -Path $userXmlFile
        Write-Host "Exported user data to: $userXmlFile" -ForegroundColor Green

        # Retrieve manager information
        try {
            $azureManager = Get-AzureADUserManager -ObjectId $azureUser.ObjectId -ErrorAction Stop
        } catch {
            Write-Warning "No manager found for $upn"
            $azureManager = $null
        }

        if ($azureManager) {
            # Export manager information
            $managerXmlFile = Join-Path -Path $csvFolder -ChildPath "$safeUpn.manager.xml"
            $azureManager | Export-Clixml -Path $managerXmlFile
            Write-Host "Exported manager data to: $managerXmlFile" -ForegroundColor Green
        }

        # Query the local AD user using the UPN
        try {
            $localAdUser = Get-ADUser -Filter "UserPrincipalName -eq '$($azureUser.UserPrincipalName)'" -ErrorAction Stop
        } catch {
            Write-Warning "Local AD user not found for $($azureUser.UserPrincipalName)"
            continue
        }
        
        # Update local AD attributes – individual Set-ADUser calls
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

        # Set the local AD manager, if available
        if ($azureManager) {
            try {
                $localManager = Get-ADUser -Filter "UserPrincipalName -eq '$($azureManager.UserPrincipalName)'" -ErrorAction Stop
                Set-ADUser -Identity $localAdUser.DistinguishedName -Manager $localManager.DistinguishedName
                Write-Host "Set local AD manager for $($azureUser.UserPrincipalName)" -ForegroundColor Green
            } catch {
                Write-Warning "Failed to set local AD manager for $($azureUser.UserPrincipalName): $($_.Exception.Message)"
            }
        }

        # Compute and display the ImmutableID
        try {
            $refreshedUser = Get-ADUser -Identity $localAdUser.ObjectGUID -Properties ObjectGUID
            $immutableId = [System.Convert]::ToBase64String($refreshedUser.ObjectGUID.ToByteArray())
            Write-Host "INFO: Computed ImmutableID (based on ObjectGUID) for '$upn': $immutableId"
        } catch {
            Write-Verbose "WARNING: Could not compute ImmutableID for '$upn'. Details: $($_.Exception.Message)"
        }

        # Set ExtensionAttribute10 to "365"
        try {
            Set-ADUser -Identity $localAdUser.DistinguishedName -Replace @{extensionAttribute10 = "365"}
            Write-Host "Set extensionAttribute10 to 365 for $($azureUser.UserPrincipalName)" -ForegroundColor Green
        } catch {
            Write-Warning "Failed to set extensionAttribute10 for $($azureUser.UserPrincipalName): $($_.Exception.Message)"
        }
    }
}
#endregion Process CSV Records

#region Final Information
Write-Host "Execute an Entra Sync Delta now by running the appropriate command." -ForegroundColor Cyan
Write-Host "Note: After the Sync Delta, remember to activate the remote mailboxes." -ForegroundColor Cyan
#endregion Final Information
