# Set exit on error
$ErrorActionPreference = 'Stop'; $DebugPreference = 'Continue'
# Function: Install-SharedDependencies
# Description: This function installs the shared dependencies for the M365PowerKit module.
function Install-SharedDependencies {
    function Get-PSModules {
        $REQUIRED_MODULES = @('ExchangeOnlineManagement')
        $REQUIRED_MODULES | ForEach-Object {
            if (-not (Get-Module -ListAvailable -Name $_)) {
                try {
                    Install-Module -Name $_
                    Write-Debug "$_ module installed successfully"
                }
                catch {
                    Write-Error "Failed to install $_ module"
                }
            }
            else {
                Write-Debug "$_ module already installed"
            }
            try {
                Import-Module -Name $_
                Write-Debug "Loading the $_ module..."
                Write-Debug "$_ module loaded successfully"
            }
            catch {
                Write-Error "Failed to import $_ module"
            }
        }
        Write-Debug ' All required modules imported successfully'
    }
    function Test-PowerShellVersion {
        $MIN_PS_VERSION = (7, 3)
        if ($PSVersionTable.PSVersion.Major -lt $MIN_PS_VERSION[0] -or ($PSVersionTable.PSVersion.Major -eq $MIN_PS_VERSION[0] -and $PSVersionTable.PSVersion.Minor -lt $MIN_PS_VERSION[1])) { Write-Host "Please install PowerShell $($MIN_PS_VERSION[0]).$($MIN_PS_VERSION[1]) or later, see: https://learn.microsoft.com/en-us/powershell/scripting/install/installing-powershell-on-windows" -ForegroundColor Red; exit }
    }
    Write-Debug 'Installing required PS modules...'
    Get-PSModules
    Write-Debug 'Required modules installed successfully...'
}


# Function: New-IPPSSession
# Description: This function creates a new Exchange Online PowerShell session.
function New-IPPSSession {
    # Check if there is an existing session
    if (!$env:M365PowerKitUPN) {
        Write-Error 'No UPN found in the environment variable M365PowerKitUPN'
    }
    else {
        try {
            Write-Debug 'Starting New-IPPSSession...'
            Connect-IPPSSession -UserPrincipalName $env:M365PowerKitUPN
        }
        catch {
            Write-Debug 'Failed to create Exchange Online PowerShell session, see:'
            Write-Debug '   - https://learn.microsoft.com/en-us/powershell/exchange/connect-to-scc-powershell?view=exchange-ps'
            Write-Error 'Failed establish IPS session'
        }
        Write-Debug 'IPS session created successfully'
    }
}

# New EXO Session
function New-EXOSession {
    # Check if there is an existing session
    if (!$env:M365PowerKitUPN) {
        Write-Error 'No UPN found in the environment variable M365PowerKitUPN'
    }
    else {
        try {
            Write-Debug 'Starting New-EXOSession...'
            Connect-ExchangeOnline -UserPrincipalName $env:M365PowerKitUPN
        }
        catch {
            Write-Debug 'Failed to create Exchange Online PowerShell session, see:'
            Write-Debug '   - https://learn.microsoft.com/en-us/powershell/exchange/connect-to-scc-powershell?view=exchange-ps'
            Write-Error 'Failed establish EXO session'
        }
        Write-Debug 'EXO session created successfully'
    }
}


# Function to authenticate using APP ID and Secret
function New-OAUTH2Session {
    param (
        [Parameter(Mandatory = $false)]
        [string]$AppID,
        [Parameter(Mandatory = $false)]
        [string]$TenantID,
        [Parameter(Mandatory = $false)]
        [securestring]$AppSecret
    )
    if (-not $AppID -or -not $TenantID -or -not $AppSecret) {
        Write-Output 'Required parameters: -AppID, -TenantID, and -AppSecret'
        Write-Error 'Parameters missing'
    }
    $OAUTH2Session = @{
        AppID        = $AppID
        TenantID     = $TenantID
        ClientSecret = $AppSecret
    }


    $body = @{
        Grant_Type    = 'client_credentials'
        Scope         = 'https://graph.microsoft.com/.default'
        Client_Id     = $appid
        Client_Secret = $secret
    }
 
    $connection = Invoke-RestMethod `
        -Uri https://login.microsoftonline.com/$tenantid/oauth2/v2.0/token `
        -Method POST `
        -Body $body
 
    $token = $connection.access_token

    $secureToken = ConvertTo-SecureString $token -AsPlainText -Force
 
    Connect-MgGraph -AccessToken $secureToken -NoWelcome
    return $OAUTH2Session
}