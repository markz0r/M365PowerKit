<#
.SYNOPSIS
This module contains functions for:
 - creating M365 Exchange Security & Compliance
 - exporting compliance search results to a PST file
    - downloading the PST file
    - extracting attachments from the PST file
    
.DESCRIPTION
The module contains the following functions:
- Get-M365ExchangeAttachments: This is the main function that calls the other functions. It takes a user principal name (UPN), mailbox name, start date, and subject as parameters.
- Get-M365ExchangeAttachmentsFromSearch: This is an alternative way to run the script when a search is already created and completed. It only needs the search name and UPN.

.PARAMETER UPN 
The user principal name (UPN) of the user running the script.

.PARAMETER MailboxName
The mailbox name of the user whose email attachments you want to retrieve.

.PARAMETER StartDate
The start date for the search. The script will only include emails received after this date format: YYYY-MM-DD.

.PARAMETER Subject
The subject of the email attachments you want to retrieve.

.PARAMETER Sender
The sender of the email attachments you want to retrieve.

.PARAMETER AttachmentExtension
The extension of the email attachments you want to retrieve.

.PARAMETER BASE_DIR
The base directory where the PST files will be saved.

.PARAMETER DisableDebug
Disables debug output.

.PARAMETER InstallDepsOnly
Installs the required dependencies only.

.PARAMETER SkipModules
Skips importing the required modules.

.PARAMETER SkipConnIPS
Skips connecting to Exchange Online PowerShell.

.PARAMETER SkipDownload
Skips downloading the PST files.

.EXAMPLE
To install the module, run the following commands:
git clone https://github.com/markz0r/Get-M365ExchangeAttachmentsBySearch.git
cd Get-M365ExchangeAttachmentsBySearch
Import-Module .\Get-M365ExchangeAttachmentsBySearch.psd1 -force
Get-M365ExchangeAttachments -InstallDepsOnly
Get-M365ExchangeAttachments -UPN "user@example.com" -MailboxName "mailbox@example.com" -StartDate "2024-01-01" -Subject "Important Documents" -Sender "test.example" -AttachmentExtension "pdf" 

.EXAMPLE
PS> Get-M365ExchangeAttachments -UPN "user@example.com" -MailboxName "mailbox@example.com" -StartDate "2024-01-01" -Subject "Important Documents" -Sender "test.example" -AttachmentExtension "pdf" 
This example retrieves email attachments from the specified mailbox that have the subject "Important Documents*" and were received after 2024-01-01. The script will only include emails from the sender address that contains "test.example" and attachments with the extension ".pdf".

.EXAMPLE
PS> Get-M365ExchangeAttachmentsFromSearch -UPN "user@example.com" -SearchName "Search1"
This example retrieves email attachments from a previously created and completed compliance search with the name "Search1".


.NOTES
This module requires the ExchangeOnlineManagement module and the Microsoft Office Unified Export Tool to be installed on the machine running the script. 
The script will attempt to automatically download dependencies if they are not already installed.

.LINK
GitHub: https://github.com/markz0r/Get-M365ExchangeAttachmentsBySearch
#>


# On any error, stop the script
# Function: Get-M365ExchangeAttachments
# Description: This is the main function that calls the other functions. It takes a user principal name (UPN), mailbox name, start date, and subject as parameters.
function Get-M365ExchangeAttachments {
    param (
        [Parameter(Mandatory = $false)]
        [string]$UPN,
        [Parameter(Mandatory = $false)]
        [string]$MailboxName,
        [Parameter(Mandatory = $false)]
        [string]$Subject,
        [Parameter(Mandatory = $false)]
        [string]$StartDate,
        [Parameter(Mandatory = $false)]
        [string]$Sender,
        [Parameter(Mandatory = $false)]
        [string]$AttachmentExtension,
        [Parameter(Mandatory = $false)]
        [string]$BASE_DIR = 'PSTExports',
        [Parameter(Mandatory = $false)]
        [switch]$SkipModules = $false,
        [Parameter(Mandatory = $false)]
        [switch]$SkipConnIPS = $false,
        [Parameter(Mandatory = $false)]
        [switch]$SkipDownload = $false,
        [Parameter(Mandatory = $false)]
        [switch]$DisableDebug = $false,
        [Parameter(Mandatory = $false)]
        [switch]$InstallDepsOnly = $false
    )
    # Catch no upn provided
    if (!$InstallDepsOnly -and (!$UPN -or !$MailboxName -or !$Subject)) {
        Write-Error 'Insufficient args provided, unless installing dependencies only, you must provide: "-UPN", "-MailboxName", "-StartDate", "-Subject" and "-Sender"parameters.'
    } 
    $ErrorActionPreference = 'Stop'; $DebugPreference = 'Continue'
    if ($DisableDebug) {
        $DebugPreference = 'SilentlyContinue'
    }
    if ($InstallDepsOnly) {
        Install-Dependencies
        return
    }
    # If either or both of the StartDate and Subject parameters are not provided, advise user and ask if they want to continue anyway
    if (-not $StartDate -or -not $Sender) {
        # Provide red warning message to the user the parameters were not provide and this query may take a long time and create a large amount of data, make the text red
        Write-Debug 'Warning: You have not provided the StartDate and/or Sender parameters. This query may take a long time and create a large amount of data.' -ForegroundColor Red
        $Continue = Read-Host 'Do you want to continue anyway? (Y/N)'
        if ($Continue -ne 'Y') {
            Write-Debug "You can provide these parameters as follows: -StartDate 'yyyy-MM-dd' -Sender 'sumologic.com'"
            Write-Debug "A full example would be: Get-M365ExchangeAttachments -UPN test@test.com -MailboxName billy@test.com -StartDate '2022-01-01' -Subject 'Important Documents' -Sender 'sumologic.com' -AttachmentExtension '.pdf'"
            Write-Debug 'Exiting script...'
            exit
        }
        else {
            Write-Debug 'Continuing without StartDate and/or Sender parameters...'
        }
    }
    else {
        if ($StartDate -notmatch '^\d{4}-\d{2}-\d{2}$') {
            Write-Error "StartDate: $StartDate - does not match the format yyyy-MM-dd."
        }
    }
    if ($AttachmentExtension -and $AttachmentExtension -notmatch '^\..*') {
        Write-Debug 'AttachmentExtension does not start with a dot, adding a dot to the start'
        $AttachmentExtension = ".$AttachmentExtension"
    }
    # Import the required modules
    Write-Debug 'Checking PSModules ...'
    Get-PSModules
    if (-not $SkipConnIPS) {
        New-IPPSSession -UPN $UPN
    }
    $SearchName = "$(Get-Date -Format 'yyyyMMdd_hhmmss')-Export-Job"
    Write-Debug "Creating a new compliance search for mailbox: $MailboxName, start date: $StartDate, subject: $Subject, sender: $Sender..."
    Write-Debug '--------------------------------------------------'
    Write-Debug "Search Name: $SearchName"
    Write-Debug '--------------------------------------------------'
    New-CustomComplianceSearch -SearchName $SearchName -MailboxName $MailboxName -fDate $StartDate -subject $Subject -Sender $Sender
    Wait-CustomComplianceSearch -SearchName $SearchName
    Get-M365ExchangeAttachmentsFromSearch -UPN $UPN -SearchName $SearchName -SkipModules -SkipConnIPS -AttachmentExtension $AttachmentExtension -BASE_DIR "$BASE_DIR"
}

# Function: Get-M365ExchangeAttachmentsFromSearch
# Description: This is an alternative way to run the script when a search is already created and completed. It only needs the search name and UPN.
function Get-M365ExchangeAttachmentsFromSearch {
    param (
        [Parameter(Mandatory = $false)]
        [string]$UPN,
        [Parameter(Mandatory = $true)]
        [string]$SearchName,
        [Parameter(Mandatory = $false)]
        [string]$AttachmentExtension,
        [Parameter(Mandatory = $false)]
        [switch]$SkipModules = $false,
        [Parameter(Mandatory = $false)]
        [switch]$SkipConnIPS = $false,
        [Parameter(Mandatory = $false)]
        [switch]$SkipDownload = $false,
        [Parameter(Mandatory = $false)]
        [string]$BASE_DIR = 'PSTExports',
        [Parameter(Mandatory = $false)]
        [switch]$DisableDebug = $false
    )
    $ErrorActionPreference = 'Stop'; $DebugPreference = 'Continue'
    if ($DisableDebug) {
        $DebugPreference = 'SilentlyContinue'
    }
    $SEARCH_DIR = "$BASE_DIR\$SearchName"
    if (-not (Test-Path -Path $BASE_DIR -PathType Container)) {
        New-Item -Path $BASE_DIR -ItemType Directory -Force
    }
    if (-not (Test-Path -Path $SEARCH_DIR -PathType Container)) {
        New-Item -Path $SEARCH_DIR -ItemType Directory -Force
    }
    if (!$UPN -and !$SkipDownload) {
        Write-Error 'UPN is required to run the script, unless -SkipDownload is used...'
    } 
    $ScriptParams = @{
        Export_RunTime      = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        SearchName          = $SearchName
        SkipModules         = $SkipModules
        SkipConnIPS         = $SkipConnIPS
        AttachmentExtension = $AttachmentExtension
        SkipDownload        = $SkipDownload
        BASE_DIR            = $BASE_DIR
    }
    $ScriptParams | Out-File -FilePath "$SEARCH_DIR\$SearchName-ScriptParams.txt"
    if ($AttachmentExtension -and $AttachmentExtension -notmatch '^\..*') {
        Write-Debug 'AttachmentExtension does not start with a dot, adding a dot to the start'
        $AttachmentExtension = ".$AttachmentExtension"
    }
    if (-not $SkipModules) {
        # Import the required modules
        Install-Dependencies
    }
    if (!$SkipConnIPS -and !$SkipDownload) {
        New-IPPSSession -UPN $UPN
    }
    # Export the compliance search results to a PST file
    # Get the Outlook COM object
    if (-not $SkipDownload) {
        try {
            $ClickOnceApplication_Exe = Get-ClickOnceApplication
            if (-not $ClickOnceApplication_Exe) {
                Write-Error 'Failed to get MS Unified Export Tool Application - get from: https://complianceclientsdf.blob.core.windows.net/v16/Microsoft.Office.Client.Discovery.UnifiedExportTool.application'
            }
            Export-CustomComplianceSearchResults -SearchName $SearchName -ClickOnceApplicationExecutable $ClickOnceApplication_Exe -SEARCH_DIR "$SEARCH_DIR"
        }
        catch {
            Write-Debug ($_ | Format-List * | Out-String)
            Write-Error "Export-CustomComplianceSearchResults -SearchName $SearchName -ClickOnceApplicationExecutable $ClickOnceApplication_Exe -SEARCH_DIR $SEARCH_DIR failed..."
        }
        Write-Debug 'Exported compliance search results to a PST file using Export-CustomComplianceSearchResults function successfully...'
    }
    Write-Debug 'Getting Outlook COM object...'
    $outlook = Get-OutlookObject
    Write-Debug 'Outlook COM object obtained successfully...'
    Write-Debug 'Saving attachments...'
    Get-ChildItem -Path "$SEARCH_DIR" -Filter '*.pst' -Recurse -ErrorAction Ignore | ForEach-Object {
        Write-Debug "Processing PST file: $($_.Name)"
        Get-AttachmentsFromPST -PSTFile $_.Name -outlook $outlook -SearchName $SearchName -AttachmentExtension $AttachmentExtension -SEARCH_DIR "$SEARCH_DIR"
    }
    Write-Debug 'Attachments saved successfully...'
    Write-Debug 'Closing the Outlook COM object...'
    $outlook.Quit()
    Start-Sleep -Seconds 5
    # Check if the Outlook process is running
    $outlookProcess = Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue
    if ($outlookProcess) {
        Write-Debug 'Outlook process is running, closing the existing instance...'
        Stop-Process -Name OUTLOOK -Force
    }
    Write-Debug 'Get-M365ExchangeAttachmentsFromSearch: Script completed successfully'
}


# Function: Install-Dependencies
# Description: This function installs the required modules and dependencies for the script to run.
function Install-Dependencies {
    Write-Debug 'Installing required PS modules...'
    Get-PSModules
    Write-Debug 'Required modules installed successfully...'
    Write-Debug 'Verifying/Installing Unified Export Tool...'
    Get-ClickOnceApplication
    Write-Debug 'ClickOnceApplication present...'
}

function Get-ClickOnceApplication {
    $ClickOnceApp = (Get-ChildItem -Path "$($env:LOCALAPPDATA)\Apps\2.0\" -Filter microsoft.office.client.discovery.unifiedexporttool.exe -Recurse).FullName | Where-Object { $_ -notmatch '_none_' } | Select-Object -First 1
    $ClickOnceAppUri = 'https://complianceclientsdf.blob.core.windows.net/v16/Microsoft.Office.Client.Discovery.UnifiedExportTool.application'
    if ($ClickOnceApp) {
        Write-Debug "ClickOnce Application Installed - Path: $ClickOnceApp"
    }
    else {
        try {
            Write-Debug 'Starting installation of UnifiedExportTool Application via browser'
            Start-Process -FilePath $ClickOnceAppUri
            # Ask user to hit enter once the application is installed or 'c' to cancel and install manually
            $Input = Read-Host 'Press Enter once the application is installed or "c" to cancel and install manually'
            if ($Input -eq 'c') {
                # fail, invoking catch block
                throw 'User cancelled installation of UnifiedExportTool Application'
            }
            else {
                Write-Debug 'Continuing with the script...'
            }
        }
        catch {
            Write-Debug "Failed to install UnifiedExportTool Application, suggest installing manually via: $ClickOnceAppUri"
            Write-Error 'Failed to install UnifiedExportTool Application'
        }
    }
    $ClickOnceApp = (Get-ChildItem -Path "$($env:LOCALAPPDATA)\Apps\2.0\" -Filter microsoft.office.client.discovery.unifiedexporttool.exe -Recurse).FullName | Where-Object { $_ -notmatch '_none_' } | Select-Object -First 1
    if (-not $ClickOnceApp) {
        Write-Debug 'Failed to get ClickOnceApplication, try manual install see:'
        Write-Debug '   - https://learn.microsoft.com/en-us/purview/ediscovery-configure-edge-to-export-search-results#install-and-run-the-ediscovery-export-tool'
        Write-Debug '   - https://complianceclientsdf.blob.core.windows.net/v16/Microsoft.Office.Client.Discovery.UnifiedExportTool.application'
        Write-Error 'Failed to get ClickOnceApplication'
    }
    $ClickOnceApp
}

function Get-PSModules {
    $REQUIRED_MODULES = @('ExchangeOnlineManagement')
    $REQUIRED_MODULES | ForEach-Object {
        if (-not (Get-InstalledModule -Name $_)) {
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

function New-KQLQuery {
    param (
        [Parameter(Mandatory = $true)]
        [string]$FDate,
        [Parameter(Mandatory = $true)]
        [string]$Subject,
        [Parameter(Mandatory = $false)]
        [string]$Sender
    )

    $KQL_QUERY_STRING = '((received>={0}) AND (subject:"{1}")' -f $FDate, $Subject
    if ($Sender) {
        $KQL_QUERY_STRING += ' AND (participants:{0})' -f $Sender
    }
    $KQL_QUERY_STRING += ')'
    return $KQL_QUERY_STRING
}

# Function: New-CustomComplianceSearch
# Description: This function creates a new compliance search in Exchange Online for a specific mailbox, date, and subject.
function New-CustomComplianceSearch {
    param (
        [Parameter(Mandatory = $true)]
        [string]$MailboxName,
        [Parameter(Mandatory = $true)]
        [string]$SearchName,
        [Parameter(Mandatory = $true)]
        [string]$Subject,
        [Parameter(Mandatory = $false)]
        [string]$FDate,
        [Parameter(Mandatory = $false)]
        [string]$Sender
    )
    # New-KQLQuery -StartDate "2024-04-10" -Subject "SSG-OpsWeekly" -Sender "sumologic.com"
    # Create a new compliance search
    $KQL_QUERY_STRING = New-KQLQuery -FDate $FDate -Subject $Subject -Sender $Sender
    New-ComplianceSearch -Name "$SearchName" -ExchangeLocation $MailboxName -ContentMatchQuery "$KQL_QUERY_STRING" -AllowNotFoundExchangeLocationsEnabled $true -Confirm:$false
    Write-Debug "Search created successfully - Search Name: $SearchName"
    Start-Sleep -Seconds 5
    Start-ComplianceSearch -Identity $SearchName
    Write-Debug "Search started successfully - Search Name: $SearchName"
}

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

# Function: Wait-CustomComplianceSearch
# Description: This function waits for the compliance search to complete.
function Wait-CustomComplianceSearch {
    param (
        [Parameter(Mandatory = $true)]
        [string]$SearchName
    )
    # Wait for the the search to be complete - DO NOT USE the LongRunningFunction function here
    $SLEEP_TIME = 5
    Write-Debug "Starting Wait-CustomComplianceSearch - SearchName = $SearchName"
    Write-Debug "Starting with a sleep time of $SLEEP_TIME seconds..."
    Start-Sleep -Seconds $SLEEP_TIME
    while ((Get-ComplianceSearch -Identity $SearchName -ErrorAction SilentlyContinue).Status -ne 'Completed' ) {
        Write-Debug "Cheking if  $SearchName search is complete..."
        Write-Debug "Current status: $((Get-ComplianceSearch -Identity $SearchName -ErrorAction SilentlyContinue).Status)"
        Write-Debug '--------------------------------------------------'
        if ((Get-ComplianceSearchAction -Identity "$SearchName" -ErrorAction SilentlyContinue).Status -ne 'Completed') {
            Write-Debug "Sleeping for $SLEEP_TIME seconds..."
            Start-Sleep -Seconds $SLEEP_TIME
        }
    }
    Write-Debug "Search completed successfully - Search Name: $SearchName"
}

# Function: Wait-CustomComplianceSearchExport
# Description: This function waits for the compliance search export to complete and returns the download URL.
function Wait-CustomComplianceSearchExport {
    param (
        [Parameter(Mandatory = $true)]
        [string]$SearchName
    )
    $EXPORT_ACTION_NAME = "${SearchName}_Export"
    # Wait for the the search action to be complete - DO NOT USE the LongRunningFunction function here
    $SLEEP_TIME = 5
    Write-Debug "Starting with a sleep time of $SLEEP_TIME seconds..."
    Start-Sleep -Seconds $SLEEP_TIME
    while ((Get-ComplianceSearchAction -Identity "$EXPORT_ACTION_NAME" -ErrorAction SilentlyContinue).Status -ne 'Completed') {
        Write-Debug "Cheking if $EXPORT_ACTION_NAME action is complete..."
        Write-Debug "Current status: $((Get-ComplianceSearchAction -Identity $EXPORT_ACTION_NAME -ErrorAction SilentlyContinue).Status)"
        # Display all properties of the search action export to the console
        if ((Get-ComplianceSearchAction -Identity $EXPORT_ACTION_NAME -ErrorAction SilentlyContinue).Status -ne 'Completed') {
            Write-Debug "$EXPORT_ACTION_NAME not finished, Sleeping for $SLEEP_TIME seconds..."
            Start-Sleep -Seconds $SLEEP_TIME
        }
    }
    $COMPLETED_JOB_RESULTS = (Get-ComplianceSearchAction -Identity "$EXPORT_ACTION_NAME" -IncludeCredential -Details).Results
    # Wait for $COMPLETED_JOB_RESULTS to contain something very similar to: "SAS token: ?sv=2018-03-28&sr=c&si=eDiscoveryBlobPolicy9%7C0&sig=7J50Xdz%2BmKIg0b6SM8iGLm3gwUpw0KKHxJZwxpGcfas%3D; "
    while ($COMPLETED_JOB_RESULTS -notmatch 'SAS token: \?.* ') {
        Write-Debug 'Waiting for the SAS token to be generated...'
        Write-Debug "Current results: $COMPLETED_JOB_RESULTS"
        Write-Debug '--------------------------------------------------'
        Write-Debug "Sleeping for $SLEEP_TIME seconds..."
        Start-Sleep -Seconds $SLEEP_TIME
    }
    $COMPLETED_JOB_RESULTS = (Get-ComplianceSearchAction -Identity "$EXPORT_ACTION_NAME" -IncludeCredential -Details).Results
    $CONTAINER_URL = $COMPLETED_JOB_RESULTS -replace '.*Container url: (.*?);.*', '$1'
    $SAS_TOKEN = $COMPLETED_JOB_RESULTS -replace '.*SAS token: (.*?);.*', '$1'
    Write-Debug "Search action ${SearchName}_Export completed successfully:"
    $DL_DETAILS = @{
        CONTAINER_URL = $CONTAINER_URL
        SAS_TOKEN     = $SAS_TOKEN
    }
    return $DL_DETAILS
}

# Function: Invoke-ComplianceSearchExportDownload
# Description: Using the provided download URL, SAS token, $ClickOnceApplicationExecutable and output directory, download the export.
# Invoke-ComplianceSearchExportDownload -SearchName $SearchName -BASE_DIR $BASE_DIR -DOWNLOAD_URL $($DOWNLOAD_DETAILS.CONTAINER_URL) -EXPORT_SAS_TOKEN $($DOWNLOAD_DETAILS.SAS_TOKEN) -ClickOnceApplicationExecutable $ClickOnceApplicationExecutable
function Invoke-ComplianceSearchExportDownload {
    param (
        [Parameter(Mandatory = $true)]
        [string]$SearchName,
        [Parameter(Mandatory = $true)]
        [string]$SEARCH_DIR,
        [Parameter(Mandatory = $true)]
        [string]$DOWNLOAD_URL,
        [Parameter(Mandatory = $true)]
        [string]$EXPORT_SAS_TOKEN,
        [Parameter(Mandatory = $true)]
        [string]$ClickOnceApplicationExecutable

    )
    if (-not (Test-Path -Type Container "$SEARCH_DIR")) {
        Write-Debug "Creating directory: $SEARCH_DIR"
        New-Item -ItemType Directory -Path "$SEARCH_DIR" -Force | Out-Null
        # Allow read and write access to the directory
        Write-Debug "Setting ACL for directory: $SEARCH_DIR"
        $acl = Get-Acl $SEARCH_DIR; $acl.SetAccessRuleProtection($false, $false)
        $rule = New-Object System.Security.AccessControl.FileSystemAccessRule('Everyone', 'FullControl', 'ContainerInherit,ObjectInherit', 'None', 'Allow')
        $acl.AddAccessRule($rule); Set-Acl $SEARCH_DIR $acl
        Write-Debug "ACL set for directory: $SEARCH_DIR - Success!"
    }
    Get-ChildItem -Path $SEARCH_DIR -Filter '*.pst' -Recurse -ErrorAction Ignore | ForEach-Object {
        Throw "PST files already exists - Path: $($_.FullName) - Size: $($_.Length / 1MB) MB, CLEAN UP FIRST!!!"
        # Forcefully remove the existing PST file without confirmation
    }
    $Arguments = "-name ""$SearchName""", "-source ""$DOWNLOAD_URL""", "-key ""$EXPORT_SAS_TOKEN""", "-dest ""$SEARCH_DIR""", '-trace true'
    Write-Debug 'Starting the export download using: '
    Write-Debug "ClickOnceApplicationExecutable: $ClickOnceApplicationExecutable arguments: $Arguments outputdir: $SEARCH_DIR ----"
    # Run the export download using the ClickOnceApplicationExecutable ensuring all output is displayed
    Start-Process -FilePath "$ClickOnceApplicationExecutable" -ArgumentList $Arguments
    # Show output of the export download
    while (Get-Process microsoft.office.client.discovery.unifiedexporttool -ErrorAction SilentlyContinue) {
        Write-Debug '--------------------------------------------------'
        # Get-Process -Name microsoft.office.client.discovery.unifiedexporttool -ErrorAction SilentlyContinue | Format-List *
        Write-Debug 'Invoke-ComplianceSearchExportDownload: unifiedexporttool still running... files being downloaded and their sizes:'
        Get-ChildItem -Path $SEARCH_DIR -Filter '*.pst' -Recurse -ErrorAction Ignore | ForEach-Object {
            $FileSize = [math]::Round(($_.Length / 1MB), 2)
            Write-Debug "    - $($_.FullName) - Size: $FileSize MB"
        }   
        Write-Debug '--------------------------------------------------'
        Start-Sleep -Seconds 5
    }
    Write-Debug 'Invoke-ComplianceSearchExportDownload: unifiedexporttool finished running...'
    Write-Debug "Renaming PST files to $SearchName-*.pst"
    Get-ChildItem -Path $SEARCH_DIR -Filter '*.pst' -Recurse -ErrorAction Ignore | ForEach-Object {
        $NewName = "$SearchName-$($_.Name)"
        Write-Debug "Renaming $($_.FullName) to and Moving to  $SEARCH_DIR\$NewName"
        Move-Item -Path $_.FullName -Destination "$SEARCH_DIR\$NewName" -Force
    }
    # Set permissions on $SEARCH_DIR and all subdirectories and files
    Write-Debug "Setting ACL for directory: $SEARCH_DIR"
    $acl = Get-Acl $SEARCH_DIR; $acl.SetAccessRuleProtection($false, $false)
    $rule = New-Object System.Security.AccessControl.FileSystemAccessRule('Everyone', 'FullControl', 'ContainerInherit,ObjectInherit', 'None', 'Allow')
    $acl.AddAccessRule($rule); Set-Acl $SEARCH_DIR $acl
    Write-Debug "ACL set for directory: $SEARCH_DIR - Success!"
    Write-Debug 'Download process completed successfully - PST file path(s): '
    Write-Debug '**********************************************************'
    (Get-ChildItem -Path $SEARCH_DIR -Filter '*.pst' -Recurse -ErrorAction Ignore) | ForEach-Object {
        # File size in MB
        $FileSize = [math]::Round(($_.Length / 1MB), 2)
        Write-Debug "    - $($_.FullName) - Size: $FileSize MB"
    }
    Write-Debug '**********************************************************'
    Get-ChildItem -Path $SEARCH_DIR -Filter '*.pst' -Recurse -ErrorAction Ignore
}

# Function: Export-CustomComplianceSearchResults
# Description: This function exports the compliance search results to a PST file and downloads it.
function Export-CustomComplianceSearchResults {
    param (
        [Parameter(Mandatory = $true)]
        [string]$SearchName,
        [Parameter(Mandatory = $true)]
        [string]$ClickOnceApplicationExecutable,
        [Parameter(Mandatory = $true)]
        [string]$SEARCH_DIR

    )
    # Export the compliance search results to a PST file - DO NOT USE the LongRunningFunction function here 
    # New-ComplianceSearchAction -SearchName $SearchName -Export -ExchangeArchiveFormat SinglePst -Format FxStream -Scope BothIndexedAndUnindexedItems -Confirm:$false | Out-Null
    New-ComplianceSearchAction -SearchName $SearchName -Export -ExchangeArchiveFormat SinglePst -Format FxStream -Scope IndexedItemsOnly -Confirm:$false | Out-Null
    $DOWNLOAD_DETAILS = Wait-CustomComplianceSearchExport -SearchName $SearchName
    Write-Debug 'Export-CustomComplianceSearchResults Running: '
    Write-Debug "Invoke-ComplianceSearchExportDownload -SearchName $SearchName -SEARCH_DIR $SEARCH_DIR -DOWNLOAD_URL $($DOWNLOAD_DETAILS.CONTAINER_URL) -EXPORT_SAS_TOKEN $($DOWNLOAD_DETAILS.SAS_TOKEN) -ClickOnceApplicationExecutable $ClickOnceApplicationExecutable)"
    try {
        Invoke-ComplianceSearchExportDownload -SearchName $SearchName -SEARCH_DIR $SEARCH_DIR -DOWNLOAD_URL $($DOWNLOAD_DETAILS.CONTAINER_URL) -EXPORT_SAS_TOKEN $($DOWNLOAD_DETAILS.SAS_TOKEN) -ClickOnceApplicationExecutable $ClickOnceApplicationExecutable
    } 
    catch {
        throw "Invoke-ComplianceSearchExportDownload -SearchName $SearchName -SEARCH_DIR $SEARCH_DIR -DOWNLOAD_URL $($DOWNLOAD_DETAILS.CONTAINER_URL) -EXPORT_SAS_TOKEN $($DOWNLOAD_DETAILS.SAS_TOKEN) -ClickOnceApplicationExecutable $ClickOnceApplicationExecutable)"
    }
    # Write to the console that the export was successful and print the full (absolute) path to the PST file
    # Get the absolute path of $SearchName
    Write-Debug 'Export completed successfully - PST file path(s): '
    Write-Debug '**********************************************************'
    (Get-ChildItem -Path "$SEARCH_DIR" -Filter '*.pst' -Recurse -ErrorAction Ignore) | ForEach-Object {
        # File size in MB
        $FileSize = [math]::Round(($_.Length / 1MB), 2)
        Write-Debug "    - $($_.FullName) - Size: $FileSize MB"
    }
    Write-Debug '**********************************************************'
}

function Save-Attachments {
    param (
        [Parameter(Mandatory = $true)]
        $folder,
        [Parameter(Mandatory = $true)]
        [string]$SEARCH_DIR,
        [Parameter(Mandatory = $false)]
        [string]$AttachmentExtension
    )
    Write-Debug "Checking folder: $($folder.Name)"
    $SEARCH_DIR_OBJECT = Get-Item -Path $SEARCH_DIR
    $SAVE_PATH = $($SEARCH_DIR_OBJECT.FullName)
    #Write-Debug '--------------------------------------------------'
    #Write-Debug ($folder | Format-List * | Out-String)
    #Write-Debug '--------------------------------------------------'
    # Write all properties of the folder to the console, ensuring it is written as Debug output
    foreach ($item in $folder.Items) {
        # Get the email subject
        Write-Debug '--------------------------------------------------'
        Write-Debug "Checking item: $($item.Subject)"
        Write-Debug "Item properties: $($item | Format-List * | Out-String)"
        Write-Debug '--------------------------------------------------'
        # If the item has attachments, save them
        if ($item.Attachments.Count -gt 0) {
            $ReceivedDate = $item.ReceivedTime
            $ReceivedDateString = $ReceivedDate.ToString('yyyy-MM-dd_HHmm')
            $Subject = $item.Subject
            Write-Debug "Found email with attachments: $Subject - Received datetime: $ReceivedDateString"
            foreach ($attachment in $item.Attachments) {
                # If $AttachmentExtension is provided, only save attachments with that extension
                if ($AttachmentExtension -and ($attachment.FileName -notlike "*$AttachmentExtension")) {
                    Write-Debug "Skipping attachment with extension: $($attachment.FileName), it does not match the provided extension filter: $AttachmentExtension"
                }
                else {
                    if ($AttachmentExtension) {
                        $FILENAMEEXT = $AttachmentExtension
                    }
                    else {
                        $FILENAMEEXT = $attachment.FileName.Split('.')[-1]
                        if ($FILENAMEEXT -notmatch '^[a-zA-Z0-9]{1,4}$') {
                            Write-Warning "Unknown file extension: $FILENAMEEXT - using anyway"
                        }
                    }
                    $FILENAME = "$($Subject)-$($ReceivedDateString)$($FILENAMEEXT)"
                    $i = 1
                    while (Test-Path "$SAVE_PATH\$FILENAME") {
                        $FILENAME = "$Subject-$ReceivedDateString-$i$FILENAMEEXT"
                        $i++
                    }
                    Write-Debug '**********************************************************'
                    Write-Debug "Saving attachment: $($attachment.FileName) as $SAVE_PATH\$FILENAME"
                    try {
                        $attachment.SaveAsFile("$SAVE_PATH\$FILENAME")
                    } 
                    catch {
                        Write-Error "Failed to save attachment: $($attachment.FileName) to $SAVE_PATH\$FILENAME"
                        Write-Error ($_ | Format-List * | Out-String)
                        throw "Save-Attachments: Failed to save attachment: $($attachment.FileName) to $SAVE_PATH\$FILENAME"
                    }
                    Write-Debug "Attachment saved as: $SAVE_PATH\$FILENAME"
                    Write-Debug '**********************************************************'
                }
            }
        }
        else {
            Write-Debug "Skipping email without attachments: $folder - $Subject"
        }
    }
    # Recursively call the function for subfolders
    Write-Debug 'Checking subfolders...'
    foreach ($subfolder in $($folder.Folders)) { 
        Write-Debug "Checking subfolder: $($subfolder.Name)"
        try {
            Save-Attachments -folder $subfolder -SEARCH_DIR $SEARCH_DIR -AttachmentExtension $AttachmentExtension
        }
        catch {
            Write-Debug "Failed to check subfolder: $($subfolder.Name)"
            Write-Debug ($_ | Format-List * | Out-String)
            throw "Save-Attachments: Failed to check subfolder: $($subfolder.Name)"
        }
        Write-Debug "Subfolder checked: $($subfolder.Name)"
    }
    Write-Debug "Save-Attachments: Completed for $($folder.Name)"
}

# Function: Get Outlook COM object
# Description: This function gets/starts the Outlook COM object and returns the Outlook application object
function Get-OutlookObject {
    # Check if the Outlook process is running
    $outlookProcess = Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue
    while ($outlookProcess) {
        Write-Debug 'Failed to get the existing instance, killing the existing process and starting a new instance...'
        Stop-Process -Name OUTLOOK -Force | Out-Null
        Start-Sleep -Seconds 5 | Out-Null
        $outlookProcess = Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue
    }
    while (-not $outlookProcess) {
        Write-Debug 'Confirmed  not running Outlook creating a new COM object...'
        $outlook = New-Object -ComObject Outlook.Application
        Start-Sleep -Seconds 5 | Out-Null
        $outlookProcess = Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue
    }
    if ($outlookProcess -and $outlook) {
        Write-Debug 'Outlook process is running and the COM object was created successfully...'
    }
    return $outlook
}

# Extract attachements from the PST file
# Get-AttachmentsFromPST -PSTFile $_.FullName -outlook $outlook -SearchName $SearchName -AttachmentExtension $AttachmentExtension -BASE_DIR $BASE_DIR
function Get-AttachmentsFromPST {
    param (
        [Parameter(Mandatory = $true)]
        [string]$PSTFile,
        [Parameter(Mandatory = $true)]
        [System.__ComObject]$outlook,
        [Parameter(Mandatory = $true)]
        [string]$SearchName,
        [Parameter(Mandatory = $false)]
        [string]$AttachmentExtension,
        [Parameter(Mandatory = $true)]
        [string]$SEARCH_DIR
    )
    $PST_FILE_OBJECT = Get-Item -Path "$SEARCH_DIR\$PSTFile"
    if (-not $PST_FILE_OBJECT) {
        Write-Error "PST file: $PSTFile does not exist in the output directory: $SEARCH_DIR"
        throw "Get-AttachmentsFromPST: PST file: $PSTFile does not exist in the output directory: $SEARCH_DIR"
    }
    Write-Debug "Extracting attachments from PST file: $($PST_FILE_OBJECT.FullName) to output directory: $SEARCH_DIR..."
    try {
        $NameSpace = $outlook.GetNamespace('MAPI')
    } 
    catch {
        Write-Debug '$namespace = $outlook.GetNamespace(MAPI) failed, trying to get the namespace from the Outlook application...'
        Write-Debug ($_ | Format-List * | Out-String)
        throw 'Get-AttachmentsFromPST: Failed to get the namespace from the Outlook application'
    }
    Write-Debug '$namespace = $outlook.GetNamespace(MAPI) succeeded...'
    Write-Debug 'Adding PST file: $PSTFile to the Outlook NameSpace...'
    try {
        #$storeCount = 0
        Write-Debug 'Checking existing stores'
        $NameSpace.Stores | ForEach-Object {
            #$storeCount++
            #Write-Debug '---------------------------------'
            #Write-Debug "Store($storeCount): $($_.Name)"
            #Write-Debug ($_ | Format-List * | Out-String)
            Write-Debug "Checking Store: $($_.DisplayName) [id: $($_.StoreID)]"
            if (($_.isDataFileStore) -and ($_.FilePath -eq $($PST_FILE_OBJECT.FullName))) {
                Write-Debug "Store: $($_.FilePath) matches the PST file: $($PST_FILE_OBJECT.FullName) - REMOVING..."
                $rootFolder = $_.GetRootFolder()
                $NameSpace.RemoveStore($rootFolder)
                Write-Debug "Store: $($_.DisplayName) [id: $($_.StoreID)] removed successfully..."
            }
        }
        Write-Debug "Sleeping for 10 seconds before adding the PST file: $($PST_FILE_OBJECT.FullName)"
        Start-Sleep -Seconds 10
        $NameSpace.AddStore($($PST_FILE_OBJECT.FullName))
        Write-Debug "Sleeping for 10 seconds after adding the PST file: $($PST_FILE_OBJECT.FullName)'
        Start-Sleep -Seconds 10
        Write-Debug 'PST Store: $($PST_FILE_OBJECT.FullName) added successfully..."
    }
    catch {
        # Write the error to the console but do not stop the script
        Write-Debug ($_ | Format-List * | Out-String)
        Write-Debug "PST Store: $($PST_FILE_OBJECT.FullName) failed to add to the Outlook NameSpace..."
        Write-Error '$NameSpace.AddStore($PSTFile) failed'
    }
    Write-Debug '$NameSpace.AddStore($PSTFile) succeeded...'
    $pstStore = $NameSpace.Stores | Where-Object { (($_.isDataFileStore) -and ($_.FilePath -eq $($PST_FILE_OBJECT.FullName))) }
    try {
        Write-Debug "Getting the root folder from the PST store: $($pstStore.Name) in $($pstStore.FilePath) ..."
        $rootFolder = $pstStore.GetRootFolder()
    }
    catch {
        Write-Debug '$rootFolder = $pstStore.GetRootFolder() failed'
        Write-Debug ($_ | Format-List * | Out-String)
        throw 'Get-AttachmentsFromPST: Failed to get the root folder from the PST store'
    }
    Write-Debug "rootFolder = pstStore.GetRootFolder() succeeded... for $($pstStore.Name) in $PSTFile"
    try {
        Write-Debug "START --- SAVING ATTACHMENTS... for $($pstStore.Name) in $PSTFile"
        Save-Attachments -folder $rootFolder -SEARCH_DIR $SEARCH_DIR -AttachmentExtension $AttachmentExtension
        Write-Debug "END ----SAVING ATTACHMENTS... for $($pstStore.Name) in $PSTFile"
    }
    catch {
        Write-Debug ($_ | Format-List * | Out-String)
        Write-Debug "ERROR Saving attachements from $($pstStore.Name) in $PSTFile, to output directory: $SEARCH_DIR..."
        Write-Error "Save-Attachments -folder $rootFolder -SEARCH_DIR $SEARCH_DIR -AttachmentExtension $AttachmentExtension failed..."
        throw 'Get-AttachmentsFromPST: Failed to save attachments from the root folder'
    }
    Write-Debug "Save-Attachments -SEARCH_DIR $SEARCH_DIR -AttachmentExtension $AttachmentExtension was successful for: $($pstStore.Exchange)"
    # Remove the PST file from the Outlook NameSpace
    try {
        $NameSpace.Stores | Where-Object { (($_.isDataFileStore) -and ($_.FilePath -eq $($PST_FILE_OBJECT.FullName))) } | ForEach-Object {
            if ($_.FilePath -eq $($PST_FILE_OBJECT.FullName)) {
                Write-Debug "Store: $($_.FilePath) matches the PST file: $($PST_FILE_OBJECT.FullName) - REMOVING..."
                $NameSpace.RemoveStore($rootFolder)
                Write-Debug "Store: $($_.FilePath) removed successfully..."
            }
            else {
                Write-Debug "Store: $($_.FilePath) does not match the PST file: $PSTFile - skipping..."
            }
        }
    }
    catch {
        Write-Debug '$Failed to remove the PST from NameSpace - continuing with the next PST file...'
    }
}