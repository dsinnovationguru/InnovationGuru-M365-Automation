<#
.SYNOPSIS
    Microsoft 365 Inactive Account & License Reclamation Audit Script
.DESCRIPTION
    Scans Entra ID (Azure AD) for user accounts inactive for X days, checks assigned licenses,
    and outputs a CSV audit report to assist in software asset management and cost optimization.
.AUTHOR
    Innovation Guru (https://dsinnovationguru.com)
#>

param (
    [Parameter(Mandatory = $false)]
    [int]$DaysInactive = 90,

    [Parameter(Mandatory = $false)]
    [string]$ExportPath = "C:\Logs\M365_Inactive_Users_Report.csv"
)

$ErrorActionPreference = "Stop"

# Calculate threshold date
$CutoffDate = (Get-Date).AddDays(-$DaysInactive)

Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
Connect-MgGraph -Scopes "AuditLog.Read.All", "User.Read.All" | Out-Null

Write-Host "Fetching users and sign-in activity (Threshold: $DaysInactive days)..." -ForegroundColor Green

# Query Entra ID users with sign-in activity and license properties
$Users = Get-MgUser -All -Property Id, DisplayName, UserPrincipalName, AccountEnabled, SignInActivity, AssignedLicenses

$Report = [System.Collections.Generic.List[PSObject]]::new()

foreach ($User in $Users) {
    $LastSignIn = $User.SignInActivity.LastSignInDateTime

    # Determine status if user has never signed in or logged in before the cutoff date
    if ($null -ne $LastSignIn -and $LastSignIn -lt $CutoffDate) {
        $HasLicense = ($User.AssignedLicenses.Count -gt 0)
        
        $Report.Add([PSCustomObject]@{
            DisplayName       = $User.DisplayName
            UserPrincipalName = $User.UserPrincipalName
            AccountEnabled    = $User.AccountEnabled
            LastSignInDate    = $LastSignIn
            HasActiveLicense  = $HasLicense
            LicenseCount      = $User.AssignedLicenses.Count
            AuditStatus       = "INACTIVE - Action Required"
        })
    }
    elseif ($null -eq $LastSignIn -and $User.AccountEnabled -eq $true) {
        $Report.Add([PSCustomObject]@{
            DisplayName       = $User.DisplayName
            UserPrincipalName = $User.UserPrincipalName
            AccountEnabled    = $User.AccountEnabled
            LastSignInDate    = "Never Logged In"
            HasActiveLicense  = ($User.AssignedLicenses.Count -gt 0)
            LicenseCount      = $User.AssignedLicenses.Count
            AuditStatus       = "NEVER LOGGED IN - Review Account"
        })
    }
}

# Export Results
if ($Report.Count -gt 0) {
    $Report | Export-Csv -Path $ExportPath -NoTypeInformation
    Write-Host "Audit completed! Identified $($Report.Count) inactive accounts. Report saved to: $ExportPath" -ForegroundColor Yellow
} else {
    Write-Host "No inactive accounts found matching the criteria." -ForegroundColor Green
}
