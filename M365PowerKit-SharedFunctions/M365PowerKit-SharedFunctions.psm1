# Set exit on error
$ErrorActionPreference = 'Stop'; $DebugPreference = 'Continue'


# Function: New-IPPSSession
# Description: This function creates a new Exchange Online PowerShell session.
function New-IPPSSession {
    param (
        [Parameter(Mandatory = $true)]
        [string]$UPN
    )
    try {
        Write-Debug 'Starting New-IPPSSession...'
        Connect-IPPSSession -UserPrincipalName $UPN
        Write-Debug 'IPS session created successfully'
    }
    catch {
        Write-Debug 'Failed to create Exchange Online PowerShell session, see:'
        Write-Debug '   - https://learn.microsoft.com/en-us/powershell/exchange/connect-to-scc-powershell?view=exchange-ps'
        Write-Error 'Failed establish IPS session'
    }
}