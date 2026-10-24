<#
.SYNOPSIS
    M365 24-Hour License Assignment & Tenant Capacity Executive Report
.DESCRIPTION
    Audits license assignments made within the last 24 hours from Entra ID audit logs,
    summarizes tenant-wide SKU capacity and health status, and generates a formatted 
    HTML report for stakeholders alongside a CSV output.
.AUTHOR
    Innovation Guru (https://dsinnovationguru.com)
#>

param (
    [Parameter(Mandatory = $false)]
    [string]$ReportFolderPath = "C:\Reports",

    [Parameter(Mandatory = $false)]
    [int]$HoursLookback = 24
)

$ErrorActionPreference = "Stop"

# Create output folder if it doesn't exist
if (-not (Test-Path $ReportFolderPath)) {
    New-Item -ItemType Directory -Path $ReportFolderPath -Force | Out-Null
}

$DateStamp = Get-Date -Format "yyyy-MM-dd"
$HtmlReportPath = Join-Path $ReportFolderPath "M365_License_Executive_Report_$DateStamp.html"
$CsvReportPath  = Join-Path $ReportFolderPath "M365_Daily_Assignments_$DateStamp.csv"

function Write-Log ($Message, $Color = "Green") {
    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host "[$TimeStamp] $Message" -ForegroundColor $Color
}

try {
    Write-Log "Connecting to Microsoft Graph..." "Cyan"
    Connect-MgGraph -Scopes "AuditLog.Read.All", "Organization.Read.All", "Directory.Read.All" | Out-Null

    # 1. Fetch Tenant SKU Summary & Capacity
    Write-Log "Retrieving tenant SKU capacity..."
    $SubscribedSkus = Get-MgSubscribedSku -All
    $SkuInventory = [System.Collections.Generic.List[PSObject]]::new()

    foreach ($Sku in $SubscribedSkus) {
        $TotalPurchased = $Sku.PrepaidUnits.Enabled
        $Consumed       = $Sku.ConsumedUnits
        $Available      = $TotalPurchased - $Consumed

        $UtilizationPct = if ($TotalPurchased -gt 0) {
            [math]::Round(($Consumed / $TotalPurchased) * 100, 1)
        } else { 0 }

        $Status = if ($Available -eq 0) { "EXHAUSTED" } 
                  elseif ($UtilizationPct -ge 90) { "CRITICAL (<10% Left)" } 
                  else { "HEALTHY" }

        $SkuInventory.Add([PSCustomObject]@{
            SkuName        = $Sku.SkuPartNumber
            TotalPurchased = $TotalPurchased
            AssignedSeats  = $Consumed
            AvailableSeats = $Available
            UtilizationPct = "$UtilizationPct%"
            Status         = $Status
        })
    }

    # 2. Fetch Assignments in the Last 24 Hours from Entra ID Audit Logs
    $StartTime = (Get-Date).AddHours(-$HoursLookback).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
    Write-Log "Querying Entra ID Audit Logs for license assignments since: $StartTime..." "Yellow"

    $AuditLogs = Get-MgAuditLogDirectoryAudit -Filter "activityDateTime ge $StartTime and category eq 'UserManagement'" -All

    $RecentAssignments = [System.Collections.Generic.List[PSObject]]::new()

    foreach ($Log in $AuditLogs) {
        if ($Log.ActivityDisplayName -in @("Change user license", "Add user")) {
            foreach ($Target in $Log.TargetResources) {
                # Look for license modification details in modified properties
                $LicenseMod = $Target.ModifiedProperties | Where-Object { $_.DisplayName -eq "IncludedUpdatedProperties.GroupType" -or $_.DisplayName -eq "AssignedLicenses" }
                
                if ($null -ne $LicenseMod -or $Log.ActivityDisplayName -eq "Change user license") {
                    $RecentAssignments.Add([PSCustomObject]@{
                        Timestamp       = $Log.ActivityDateTime
                        UserPrincipal   = $Target.UserPrincipalName
                        InitiatedBy     = if ($Log.InitiatedBy.User.UserPrincipalName) { $Log.InitiatedBy.User.UserPrincipalName } else { "System / Automation" }
                        Activity        = $Log.ActivityDisplayName
                        Result          = $Log.Result
                    })
                }
            }
        }
    }

    # Export Raw CSV of Recent Assignments
    $RecentAssignments | Export-Csv -Path $CsvReportPath -NoTypeInformation
    Write-Log "Assignments CSV exported to: $CsvReportPath" "Yellow"

    # 3. Generate Executive HTML Report
    Write-Log "Building Stakeholder HTML Report..." "Cyan"

    $SkuRowsHtml = ""
    foreach ($Item in $SkuInventory) {
        $StatusClass = if ($Item.Status -eq "HEALTHY") { "badge-green" } else { "badge-red" }
        $SkuRowsHtml += "<tr>
            <td><strong>$($Item.SkuName)</strong></td>
            <td>$($Item.TotalPurchased)</td>
            <td>$($Item.AssignedSeats)</td>
            <td>$($Item.AvailableSeats)</td>
            <td>$($Item.UtilizationPct)</td>
            <td><span class='$StatusClass'>$($Item.Status)</span></td>
        </tr>"
    }

    $AssignmentRowsHtml = ""
    if ($RecentAssignments.Count -gt 0) {
        foreach ($Assign in $RecentAssignments) {
            $AssignmentRowsHtml += "<tr>
                <td>$($Assign.Timestamp)</td>
                <td>$($Assign.UserPrincipal)</td>
                <td>$($Assign.Activity)</td>
                <td>$($Assign.InitiatedBy)</td>
                <td><span class='badge-green'>$($Assign.Result)</span></td>
            </tr>"
        }
    } else {
        $AssignmentRowsHtml = "<tr><td colspan='5' style='text-align:center;'>No license assignment changes detected in the last $HoursLookback hours.</td></tr>"
    }

    $TotalAssigned24h = $RecentAssignments.Count

    $HtmlContent = @"
<!DOCTYPE html>
<html>
<head>
    <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #f4f6f9; color: #333; margin: 0; padding: 20px; }
        .container { max-width: 1000px; margin: auto; background: #fff; padding: 30px; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
        h1 { color: #0078d4; font-size: 24px; margin-bottom: 5px; }
        .subtitle { color: #666; font-size: 14px; margin-bottom: 25px; }
        .metrics-grid { display: flex; gap: 15px; margin-bottom: 30px; }
        .card { flex: 1; background: #f8f9fa; border-left: 4px solid #0078d4; padding: 15px; border-radius: 4px; }
        .card-title { font-size: 12px; color: #666; text-transform: uppercase; font-weight: bold; }
        .card-value { font-size: 28px; font-weight: bold; color: #111; margin-top: 5px; }
        table { width: 100%; border-collapse: collapse; margin-bottom: 30px; }
        th { background-color: #0078d4; color: white; text-align: left; padding: 10px; font-size: 13px; }
        td { padding: 10px; border-bottom: 1px solid #ddd; font-size: 13px; }
        tr:hover { background-color: #f1f1f1; }
        .badge-green { background: #d4edda; color: #155724; padding: 4px 8px; border-radius: 4px; font-weight: bold; font-size: 11px; }
        .badge-red { background: #f8d7da; color: #721c24; padding: 4px 8px; border-radius: 4px; font-weight: bold; font-size: 11px; }
        .footer { text-align: center; font-size: 12px; color: #888; margin-top: 20px; }
    </style>
</head>
<body>
    <div class="container">
        <h1>M365 License Governance &amp; Capacity Report</h1>
        <div class="subtitle">Generated on $(Get-Date -Format "MMMM dd, yyyy HH:mm tt") | Audit Window: Last $HoursLookback Hours</div>

        <div class="metrics-grid">
            <div class="card">
                <div class="card-title">Total SKUs Tracked</div>
                <div class="card-value">$($SkuInventory.Count)</div>
            </div>
            <div class="card">
                <div class="card-title">24h License Activity</div>
                <div class="card-value">$TotalAssigned24h</div>
            </div>
            <div class="card">
                <div class="card-title">Report Scope</div>
                <div class="card-value">Tenant-Wide</div>
            </div>
        </div>

        <h2>1. Current License SKU Inventory &amp; Status</h2>
        <table>
            <thead>
                <tr>
                    <th>SKU Name</th>
                    <th>Total Purchased</th>
                    <th>Assigned</th>
                    <th>Available</th>
                    <th>Utilization</th>
                    <th>Status</th>
                </tr>
            </thead>
            <tbody>
                $SkuRowsHtml
            </tbody>
        </table>

        <h2>2. License Assignment Log (Last $HoursLookback Hours)</h2>
        <table>
            <thead>
                <tr>
                    <th>Timestamp (UTC)</th>
                    <th>Target User</th>
                    <th>Activity</th>
                    <th>Initiated By</th>
                    <th>Result</th>
                </tr>
            </thead>
            <tbody>
                $AssignmentRowsHtml
            </tbody>
        </table>

        <div class="footer">Report generated automatically by Innovation Guru M365 Automation Framework</div>
    </div>
</body>
</html>
"@

    $HtmlContent | Out-File -FilePath $HtmlReportPath -Encoding utf8
    Write-Log "Executive HTML Report successfully generated at: $HtmlReportPath" "Green"
}
catch {
    Write-Log "ERROR: $_" "Red"
}
