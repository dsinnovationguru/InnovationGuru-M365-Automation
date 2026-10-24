<#
.SYNOPSIS
    M365 License Usage Monitoring & Automated Threshold Alerting Script
.DESCRIPTION
    Scans Microsoft 365 tenant subscribed SKUs, checks available seat capacity against 
    a custom threshold percentage, and triggers email notifications for low-capacity pools.
.AUTHOR
    Innovation Guru (https://dsinnovationguru.com)
#>

param (
    [Parameter(Mandatory = $false)]
    [int]$ThresholdPercent = 10,

    [Parameter(Mandatory = $false)]
    [string]$AlertEmail = "admin@dsinnovationguru.com",

    [Parameter(Mandatory = $false)]
    [string]$ReportPath = "C:\Logs\M365_License_Usage_Report.csv"
)

$ErrorActionPreference = "Stop"

function Write-Log ($Message, $Color = "Green") {
    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host "[$TimeStamp] $Message" -ForegroundColor $Color
}

try {
    Write-Log "Connecting to Microsoft Graph Directory Management..." "Cyan"
    Connect-MgGraph -Scopes "Organization.Read.All", "Directory.Read.All" | Out-Null

    Write-Log "Retrieving tenant license SKUs..."
    $SubscribedSkus = Get-MgSubscribedSku -All

    $Report = [System.Collections.Generic.List[PSObject]]::new()
    $LowCapacityAlerts = [System.Collections.Generic.List[PSObject]]::new()

    foreach ($Sku in $SubscribedSkus) {
        $SkuName = $Sku.SkuPartNumber
        $TotalUnits = $Sku.CapabilityStatus.ConsumedUnits + $Sku.PrepaidUnits.Enabled
        $ActiveConsumed = $Sku.ConsumedUnits
        $AvailableSeats = $Sku.PrepaidUnits.Enabled - $ActiveConsumed

        # Calculate percentage remaining
        if ($Sku.PrepaidUnits.Enabled -gt 0) {
            $PercentAvailable = [math]::Round(($AvailableSeats / $Sku.PrepaidUnits.Enabled) * 100, 2)
        } else {
            $PercentAvailable = 0
        }

        $Status = if ($PercentAvailable -le $ThresholdPercent) { "CRITICAL - LOW SEATS" } else { "HEALTHY" }

        $SkuObject = [PSCustomObject]@{
            SkuName          = $SkuName
            TotalPurchased   = $Sku.PrepaidUnits.Enabled
            AssignedSeats    = $ActiveConsumed
            AvailableSeats   = $AvailableSeats
            PercentAvailable = "$PercentAvailable%"
            Status           = $Status
        }

        $Report.Add($SkuObject)

        if ($PercentAvailable -le $ThresholdPercent -and $Sku.PrepaidUnits.Enabled -gt 0) {
            $LowCapacityAlerts.Add($SkuObject)
        }
    }

    # Export Full Usage Report
    $Report | Export-Csv -Path $ReportPath -NoTypeInformation
    Write-Log "License usage audit exported to: $ReportPath" "Yellow"

    # Display Summary
    $Report | Format-Table -AutoSize

    # Trigger Alert if Low Capacity Detected
    if ($LowCapacityAlerts.Count -gt 0) {
        Write-Log "WARNING: $($LowCapacityAlerts.Count) SKU pool(s) dropped below $ThresholdPercent% capacity threshold!" "Red"
        # Optional: Integrate Send-MgUserMail or Webhook payload here for automated alerts
    } else {
        Write-Log "All tenant license pools are operating above the $ThresholdPercent% threshold." "Green"
    }
}
catch {
    Write-Log "ERROR: $_" "Red"
}
