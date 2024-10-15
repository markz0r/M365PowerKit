<#
.SYNOPSIS
This module contains functions for:
 - Exporting reports from MS Defender.
    
.DESCRIPTION
Exporting a list of software installed on Microsoft Defender for Endpoint managed devices. Uses interactive authentication to connect to the Microsoft Graph API.

.PARAMETER UPN 
The user principal name (UPN) of the user running the script.

.LINK
GitHub: https://github.com/markz0r/M365PowerKit
#>

$ErrorActionPreference = 'Stop'; $DebugPreference = 'Continue'
# Start transcript logging
#$TranscriptPath = "$PSScriptRoot\Trans\$(Get-Date -Format 'yyyyMMdd_hhmmss')-Transcript.log"

function Export-SoftwareInventory {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$UPN
    )

    # Import the module
    Import-Module -Name Microsoft.Graph.DeviceManagement

    # Connect to the Microsoft Graph Device Management API
    Connect-MgDeviceManagement -UPN $UPN

    # Get the software inventory
    $SoftwareInventory = Get-MgDeviceManagementSoftwareInventory

    # Export the software inventory
    $SoftwareInventory | Export-Csv -Path "$PSScriptRoot\Reports\SoftwareInventory.csv" -NoTypeInformation
}

function Export-InfoSecOpsWeeklyReport {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$UPN
    )

    # Import the module
    Import-Module -Name Microsoft.Graph.Security

    # Connect to the Microsoft Graph Security API
    Connect-MgGraphSecurity -UPN $UPN

    # Get the weekly report
    $Report = Get-MgSecurityInfoSecOpsWeeklyReport

    # Export the report
    $Report | Export-Csv -Path "$PSScriptRoot\Reports\InfoSecOpsWeeklyReport.csv" -NoTypeInformation
}

# MS Defender has Incident and Alert, the difference is that Incident is a collection of Alerts.
# This function will export all Incidents and Alerts to a CSV file:
function Export-IncidentsAndAlerts {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$UPN
    )

    # Import the module
    Import-Module -Name Microsoft.Graph.Security

    # Connect to the Microsoft Graph Security API
    Connect-MgGraphSecurity -UPN $UPN

    # Get all incidents
    $Incidents = Get-MgSecurityIncident

    # Get all alerts
    $Alerts = Get-MgSecurityAlert

    # Export the incidents
    $Incidents | Export-Csv -Path "$PSScriptRoot\Reports\Incidents.csv" -NoTypeInformation

    # Export the alerts
    $Alerts | Export-Csv -Path "$PSScriptRoot\Reports\Alerts.csv" -NoTypeInformation
}
