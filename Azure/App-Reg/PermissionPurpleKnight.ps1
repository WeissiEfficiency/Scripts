<#
.SYNOPSIS
    Creates an Azure AD application registration for Purple Knight with necessary permissions.
.DESCRIPTION
    This script registers an application in Azure Active Directory, assigns it necessary permissions for Purple Knight,
    and provides the application ID, tenant ID, and secret for use in Purple Knight.
.PARAMETER TenantId
    The Azure AD tenant ID where the application will be registered.
.PARAMETER AppName
    The name of the application to be registered.
.PARAMETER RequiredPermissions
    An array of permissions to be assigned to the application. Defaults to a set of permissions required by Purple Knight.
.EXAMPLE
    .\AppRegPurpleKnightv2.ps1 -TenantId "your-tenant-id" -AppName "PurpleKnightApp"
    This command registers a new application named "PurpleKnightApp" in the specified tenant with the default permissions.
.VERSION
    1.0
.AUTHOR
    Stefan Weiß
.LINK
https://www.semperis.com/purple-knight/
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$TenantId,

    [Parameter(Mandatory = $true)]
    [string]$AppName,

    [Parameter()]
    [switch]$PauseForCopy,

    [Parameter(Mandatory = $false)]
    [string[]]$RequiredPermissions = @()
)

#region functions

#Permissions required by Purple Knight v5.0
# If not specified, the script will use the default set of permissions.
 $RequiredPermissions = @("AdministrativeUnit.Read.All",
        "Application.Read.All",
        "AuditLog.Read.All",
        "Device.Read.All",
        "Directory.Read.All",
        "GroupMember.Read.All",
        "IdentityRiskyUser.Read.All",
        "MailboxSettings.Read",
        "OnPremDirectorySynchronization.Read.All",
        "Policy.Read.All",
        "PrivilegedAccess.Read.AzureAD",
        "Reports.Read.All",
        "RoleEligibilitySchedule.Read.Directory",
        "RoleManagement.Read.All",
        "RoleManagement.Read.Directory",
        "User.Read.All",
        "UserAuthenticationMethod.Read.All",
        "Organization.Read.All",
        "PrivilegedEligibilitySchedule.Read.AzureADGroup"
    )

# Check if the required modules are installed, if not, install them
#Requires -Modules  Microsoft.Graph.Applications, Microsoft.Graph.Authentication, Microsoft.Graph.Users

<#.SYNOPSIS
    Registers a new application in Azure AD and creates a service principal.
.DESCRIPTION
    This function registers a new application in Azure Active Directory, creates a service principal for it,
    and returns the application and service principal objects.
.PARAMETER TenantId
    The Azure AD tenant ID where the application will be registered.
.PARAMETER AppName
    The name of the application to be registered. If an application with this name already exists, the script will exit.
#>
function New-EntraAppRegistration {
    try {
        $ErrorActionPreference = 'Stop'

        Connect-MgGraph -TenantId $TenantId -Scopes "Application.ReadWrite.All", "AppRoleAssignment.ReadWrite.All"

        Write-Information "Checking if application '$AppName' already exists..."
        $existingApp = Get-MgApplication -Filter "displayName eq '$AppName'"
        if ($existingApp) {
            Write-Error "An application with the name '$AppName' already exists. Exiting."
            throw "Application already exists."
        }

        Write-Information "Creating app registration..."
        $app = New-MgApplication -DisplayName $AppName -IsFallbackPublicClient:$false

        Write-Information "Creating client secret..."
        $secret = Add-MgApplicationPassword -ApplicationId $app.Id -PasswordCredential @{
            displayName = "$AppName-Secret"
        }

        Write-Information "Creating service principal..."
        $sp = New-MgServicePrincipal -AppId $app.AppId

        return @{
            Application = $app
            SecretValue = $secret.SecretText
            ServicePrincipal = $sp
        }

    } catch {
        Write-Error "Error during app registration: $_"
        throw
    }
}

<#.SYNOPSIS
    Grants specified permissions to the service principal for Microsoft Graph.
.DESCRIPTION
    This function assigns the specified permissions to the service principal of the application.
    It retrieves the Microsoft Graph service principal and assigns the specified app roles.
.PARAMETER ServicePrincipal
    The service principal object of the application to which permissions will be assigned.
.PARAMETER Permissions
    An array of permission values to be assigned to the service principal.
.Link
    https://learn.microsoft.com/en-us/powershell/entra-powershell/how-to-grant-revoke-api-permissions?view=entra-powershell&pivots=grant-delegated-permissions
#>
function Set-AppPermission {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        $ServicePrincipal,

        [Parameter(Mandatory = $true)]
        [string[]]$Permissions
    )

    try {
        $ErrorActionPreference = 'Stop'

        Write-Information "Locating Microsoft Graph service principal..."
        $graphSp = Get-MgServicePrincipal -Filter "displayName eq 'Microsoft Graph'"

        foreach ($perm in $Permissions) {
            Write-Information "Assigning permission: $perm"

            $appRole = $graphSp.AppRoles | Where-Object {
                $_.Value -eq $perm -and $_.AllowedMemberTypes -contains "Application"
            }

            if ($appRole) {
                $params = @{
                    ServicePrincipalId = $ServicePrincipal.Id
                    PrincipalId        = $ServicePrincipal.Id
                    ResourceId         = $graphSp.Id
                    AppRoleId          = $appRole.Id
                }
                New-MgServicePrincipalAppRoleAssignment @params
            }
        }

        Write-Warning "Permission assignment complete. Admin consent may still be required."
    } catch {
        Write-Error "Failed to assign permissions: $_"
    }
}

<#.SYNOPSIS
    Displays and copies a value to the clipboard.
.DESCRIPTION
    This function writes a label and value to the console, copies the value to the clipboard,
    and optionally pauses for user input.
.PARAMETER Label
    The label to display before the value.
.PARAMETER Value
    The value to display and copy to the clipboard.
.PARAMETER Pause
    If specified, the script will pause and wait for user input after displaying the value.
#>
function Write-Message {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Label,

        [Parameter(Mandatory = $true)]
        [string]$Value,

        [Parameter()]
        [switch]$Pause
    )

    Write-Information "${Label}:" -InformationAction Continue
    Write-Output $Value
    $Value | Set-Clipboard
    Write-Information "(Copied to clipboard!)" -InformationAction Continue
    if ($Pause) {
        Read-Host "Press Enter after you have pasted the $Label into Purple Knight"
    }
}


<#
.SYNOPSIS
    Sets the current user as an owner of the application.
.DESCRIPTION
    This function retrieves the current signed-in user and adds them as an owner of the specified application.
.PARAMETER Application
The application object for which the current user will be set as an owner.
#>
function Set-AppOwnerCurrentUser {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [object]$Application
    )
    try {
        # Get the current signed-in user
        $currentUser = Get-MgUser -UserId (Get-MgContext).Account

        # Add the current user as an owner of the application
        New-MgApplicationOwnerByRef -ApplicationId $Application.Id -BodyParameter @{
            "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($currentUser.Id)"
        }

        Write-Information "Set current user as application owner." -InformationAction Continue
    } catch {
        Write-Warning "Failed to set current user as application owner: $_"
    }
}

#endregion functions

#region main

# Validate Creation of the app registration
$appResult = New-EntraAppRegistration
if ($null -eq $appResult) {
    Write-Error "App registration failed. Exiting script."
    exit 1
}
#Write-Information "App registration successful. Application ID: $($appResult.Application.AppId)" -InformationAction Continue
Set-AppPermission -ServicePrincipal $appResult.ServicePrincipal -Permissions $RequiredPermissions -Verbose
Write-Information "App registration and permission assignment completed successfully." -InformationAction Continue

# Output the Tenant ID, Application ID, and Secret
Write-Message -Label "Tenant ID" -Value $TenantId -Pause:$PauseForCopy -Verbose
Write-Message -Label "Application ID" -Value $appResult.Application.AppId -Pause:$PauseForCopy -Verbose
Write-Message -Label "Application Secret" -Value $appResult.SecretValue -Pause:$PauseForCopy -Verbose

Set-AppOwnerCurrentUser -Application $appResult.Application

Write-Information "Done!" -InformationAction Continue


#endregion main
