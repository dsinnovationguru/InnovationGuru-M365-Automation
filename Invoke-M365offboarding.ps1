<#
.SYNOPSIS
    Automated M365 Employee Offboarding Script
.DESCRIPTION
    Blocks sign-in, revokes tokens, converts mailbox to shared, delegates manager access,
    configures OOO reply, hides from GAL, and reclaims licenses.
.AUTHOR
    Innovation Guru (https://dsinnovationguru.com)
#>

param (
    [Parameter(Mandatory = $true)]
    [string]$UserPrincipalName,

    [Parameter(Mandatory = $true)]
    [string]$ManagerUPN,

    [Parameter(Mandatory = $false)]
    [string]$AutoReplyMessage = "This employee has left the organization. Please contact their manager for assistance."
)

$ErrorActionPreference = "Stop"
$LogPath = "C:\Logs\M365_Offboarding.log"

function Write-Log ($Message) {
    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -Path $LogPath -Value "[$TimeStamp] $Message"
    Write-Host "[$TimeStamp] $Message" -ForegroundColor Green
}

try {
    Write-Log "Starting offboarding workflow for user: $UserPrincipalName"

    # Step 1: Block Sign-In
    Update-MgUser -UserId $UserPrincipalName -AccountEnabled:$false
    Write-Log "Account sign-in disabled."

    # Step 2: Revoke Active Sessions & Refresh Tokens
    Revoke-MgUserSignInSession -UserId $UserPrincipalName | Out-Null
    Write-Log "Active user refresh tokens revoked."

    # Step 3: Convert Mailbox to Shared & Hide from GAL
    Set-Mailbox -Identity $UserPrincipalName -Type Shared -HiddenFromAddressListsEnabled $true
    Write-Log "Mailbox converted to Shared Mailbox and hidden from GAL."

    # Step 4: Delegate Full Access and Send As Permissions to Manager
    Add-MailboxPermission -Identity $UserPrincipalName -User $ManagerUPN -AccessRights FullAccess -AutoMapping $true
    Add-RecipientPermission -Identity $UserPrincipalName -Trustee $ManagerUPN -AccessRights SendAs -Confirm:$false
    Write-Log "FullAccess and SendAs permissions assigned to manager: $ManagerUPN"

    # Step 5: Set Out of Office (OOF) Auto-Reply
    Set-MailboxAutoReplyConfiguration -Identity $UserPrincipalName -AutoReplyState Enabled `
        -InternalMessage $AutoReplyMessage -ExternalMessage $AutoReplyMessage
    Write-Log "Out of Office auto-reply enabled."

    # Step 6: Reclaim Licenses (Preserving Data)
    $UserLicenses = (Get-MgUser -UserId $UserPrincipalName -Property AssignedLicenses).AssignedLicenses
    if ($UserLicenses) {
        foreach ($License in $UserLicenses) {
            Set-MgUserLicense -UserId $UserPrincipalName -RemoveLicenses @($License.SkuId) -AddLicenses @{}
        }
        Write-Log "All assigned M365 licenses revoked."
    } else {
        Write-Log "No active licenses found on account."
    }

    Write-Log "Offboarding successfully completed for $UserPrincipalName"
}
catch {
    Write-Log "ERROR: $_"
}
