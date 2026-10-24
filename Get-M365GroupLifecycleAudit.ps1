<#
.SYNOPSIS
    Microsoft 365 Groups & Teams Lifecycle Governance Audit Script
.DESCRIPTION
    Audits Microsoft 365 Groups and Teams for orphan status (no active owners),
    external guest count, visibility settings, and compliance health.
.AUTHOR
    Innovation Guru (https://dsinnovationguru.com)
#>

param (
    [Parameter(Mandatory = $false)]
    [string]$ExportPath = "C:\Logs\M365_Group_Governance_Report.csv"
)

$ErrorActionPreference = "Stop"

function Write-Log ($Message, $Color = "Green") {
    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host "[$TimeStamp] $Message" -ForegroundColor $Color
}

try {
    Write-Log "Connecting to Microsoft Graph Groups & Users..." "Cyan"
    Connect-MgGraph -Scopes "Group.Read.All", "Directory.Read.All", "GroupMember.Read.All" | Out-Null

    Write-Log "Retrieving all unified Microsoft 365 Groups & Teams..." "Green"
    $Groups = Get-MgGroup -All -Filter "groupTypes/any(c:c eq 'Unified')" -Property Id, DisplayName, Mail, Visibility, CreatedDateTime

    $Report = [System.Collections.Generic.List[PSObject]]::new()

    foreach ($Group in $Groups) {
        # Retrieve Group Owners
        $Owners = Get-MgGroupOwner -GroupId $Group.Id -All
        $ActiveOwners = 0

        if ($Owners) {
            foreach ($Owner in $Owners) {
                # Verify if owner account is enabled
                $UserObj = Get-MgUser -UserId $Owner.Id -Property AccountEnabled -ErrorAction SilentlyContinue
                if ($UserObj -and $UserObj.AccountEnabled -eq $true) {
                    $ActiveOwners++
                }
            }
        }

        # Retrieve Group Members & Guest Count
        $Members = Get-MgGroupMember -GroupId $Group.Id -All
        $GuestCount = 0
        
        foreach ($Member in $Members) {
            $MemberUser = Get-MgUser -UserId $Member.Id -Property UserType -ErrorAction SilentlyContinue
            if ($MemberUser -and $MemberUser.UserType -eq "Guest") {
                $GuestCount++
            }
        }

        # Determine Governance Status
        $GovernanceStatus = if ($ActiveOwners -eq 0) {
            "ORPHAN GROUP - No Active Owners"
        } elseif ($ActiveOwners -eq 1) {
            "WARNING - Single Owner Risk"
        } else {
            "HEALTHY"
        }

        $Report.Add([PSCustomObject]@{
            GroupName        = $Group.DisplayName
            GroupEmail       = $Group.Mail
            Visibility       = $Group.Visibility
            CreatedDate      = $Group.CreatedDateTime
            TotalOwners      = $Owners.Count
            ActiveOwners     = $ActiveOwners
            TotalMembers     = $Members.Count
            GuestMemberCount = $GuestCount
            GovernanceStatus = $GovernanceStatus
        })
    }

    # Export CSV Report
    if ($Report.Count -gt 0) {
        $Report | Export-Csv -Path $ExportPath -NoTypeInformation
        Write-Log "Group Governance Audit complete! Processed $($Report.Count) groups. Report saved to: $ExportPath" "Yellow"
        
        # Display Summary Table
        $Report | Format-Table DisplayName, Visibility, ActiveOwners, GuestMemberCount, GovernanceStatus -AutoSize
    } else {
        Write-Log "No M365 Groups found in tenant." "Yellow"
    }
}
catch {
    Write-Log "ERROR: $_" "Red"
}
