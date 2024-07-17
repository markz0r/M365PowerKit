<#
.SYNOPSIS
    Atlassian Cloud PowerKit module for interacting with Atlassian Cloud REST API.
.DESCRIPTION
    Atlassian Cloud PowerKit module for interacting with Atlassian Cloud REST API.
    - Dependencies: M365PowerKit-Shared
    - Functions:
      - Use-M365PowerKit: Interactive function to run any function in the module.
    - Debug output is enabled by default. To disable, set $DisableDebug = $true before running functions.
.EXAMPLE
    Use-M365PowerKit
    This example lists all functions in the M365PowerKit module.
.EXAMPLE
    Use-M365PowerKit
    Simply run the function to see a list of all functions in the module and nested modules.
.EXAMPLE
    Get-DefinedPowerKitVariables
    This example lists all variables defined in the M365PowerKit module.
.LINK
    GitHub:

#>
$ErrorActionPreference = 'Stop'; $DebugPreference = 'Continue'

function Get-ClickOnceApplication {
    $Default_Path = "$($env:LOCALAPPDATA)\Apps\2.0\"
    $Default_Filename = 'microsoft.office.client.discovery.unifiedexporttool.exe'
    $Default_URL = 'https://complianceclientsdf.blob.core.windows.net/v16/Microsoft.Office.Client.Discovery.UnifiedExportTool.application'
    function Write-ClickOnceInstructuions {
        Write-Host 'To install the Unified Export Tool manually, follow these steps:'
        Write-Host "   - Open a browser and navigate to: $Default_URL"
        Write-Host '   - Click on the "Install" button to download and install the application'
        Write-Host '   - Once the installation is complete, hit "C" to continue or any other key to exit'
        $Key = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown').VirtualKeyCode
        if ($Key -ne 67) {
            Write-Error 'Failed to get ClickOnceApplication'
            throw 'Failed to get ClickOnceApplication'
        }
    }
    while ((-not (Test-Path -Path $Default_Path -PathType Container)) -or (-not(Get-ChildItem -Path $Default_Path -Filter $Default_Filename -Recurse))) {
        Write-Debug 'Failed to get ClickOnceApplication, looking in '
        Write-ClickOnceInstructuions
    }
    $ClickOnceApp = (Get-ChildItem -Path $Default_Path -Filter $Default_Filename -Recurse).FullName | Where-Object { $_ -notmatch '_none_' } | Select-Object -First 1
    while (!$ClickOnceApp) {
        Write-Debug 'Failed to get ClickOnceApplication, try manual install see:'
        Write-ClickOnceInstructuions
    }
    Write-Debug "ClickOnce Application Installed - Path: $ClickOnceApp"
    $ClickOnceApp
}


# Function: Install-Dependencies
# Description: This function installs the required modules and dependencies for the script to run.
function Install-Dependencies {
    Write-Host 'Not implemented yet'
}

function Import-NestedModules {
    # Location of the M365PowerKit.psm1 file
    # Find *.psd1 files in $PSScriptRoot subdirectories and import them
    $NESTED_MODULE_ARRAY = Get-ChildItem -Path $PSScriptRoot -Filter '*.psd1' -Recurse -Exclude 'M365PowerKit.psd1', 'M365PowerKit-SharedFunctions.psd1' | ForEach-Object {
        # If the module is not already imported, import it
        if (-not (Get-Module -Name $_.BaseName)) {
            Write-Debug "Importing module: $($_.FullName)"
            Import-Module $_.FullName -Force
        }
        else {
            Write-Debug "Module already imported: $($_.FullName)"
        }
        # Validate that the module was imported
        if (-not (Get-Module -Name $_.BaseName)) {
            Write-Error "Failed to import module: $($_.FullName)"
        }
        return $_.BaseName
    }
    $NESTED_MODULE_ARRAY
}


# function to run provided functions with provided parameters (as hash table)
function Invoke-M365PowerKitFunction {
    param (
        [Parameter(Mandatory = $true)]
        [string]$FunctionName,
        [Parameter(Mandatory = $true)]
        [hashtable]$Parameters,
        [Parameter(Mandatory = $false)]
        [switch]$SkipNestedModuleImport

    )
    if (-not $SkipNestedModuleImport) { Import-NestedModules }
        
    try {
        $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        # Invoke expression to run the function, splatting the parameters
        $stopwatch.Start()
        Write-Debug "Running function: $FunctionName"
        Write-Debug "Parameters: $($Parameters | Out-String)"
        & $FunctionName @Parameters
        $stopwatch.Stop()
        Write-Debug "Function: $FunctionName completed in $($stopwatch.Elapsed.TotalSeconds) seconds"
    }
    catch {
        Write-Error "Failed to run function: $FunctionName"
    }
}

function Show-M365PowerKitFunctions {
    # List nested modules and their exported functions to the console in a readable format, grouped by module
    $colors = @('Green', 'Cyan', 'Red', 'Magenta', 'Yellow')
    $nestedModules = Import-NestedModules

    $colorIndex = 0
    $functionReferences = @{}
    $nestedModules | ForEach-Object {
        $MODULE = Get-Module -Name $_
        # Select a color from the list
        $color = $colors[$colorIndex % $colors.Count]
        $spaces = ' ' * (52 - $MODULE.Name.Length)
        Write-Host '' -BackgroundColor Black
        Write-Host "Module: $($MODULE.Name)" -BackgroundColor $color -ForegroundColor White -NoNewline
        Write-Host $spaces  -BackgroundColor $color -NoNewline
        Write-Host ' ' -BackgroundColor Black
        $spaces = ' ' * 41
        Write-Host " Exported Commands:$spaces" -BackgroundColor "Dark$color" -ForegroundColor White -NoNewline
        Write-Host ' ' -BackgroundColor Black
        $MODULE.ExportedCommands.Keys | ForEach-Object {
            # Assign a letter reference to the function
            $functRefNum = $colorIndex
            $functionReferences[$functRefNum] = $_

            Write-Host ' ' -NoNewline -BackgroundColor "Dark$color"
            Write-Host '   ' -NoNewline -BackgroundColor Black
            Write-Host "$functRefNum -> " -NoNewline -BackgroundColor Black
            Write-Host "$_" -NoNewline -BackgroundColor Black -ForegroundColor $color
            # Calculate the number of spaces needed to fill the rest of the line
            $spaces = ' ' * (50 - $_.Length)
            Write-Host $spaces -NoNewline -BackgroundColor Black
            Write-Host ' ' -NoNewline -BackgroundColor "Dark$color"
            Write-Host ' ' -BackgroundColor Black
            # Increment the color index for the next function
            $colorIndex++
        }
        $spaces = ' ' * 60
        Write-Host $spaces -BackgroundColor "Dark$color" -NoNewline
        Write-Host ' ' -BackgroundColor Black
    }
    Write-Host 'Note: You can run functoins without this interface by calling them directly.' 
    Write-Host "Example: Invoke-M365PowerKitFunction -FunctionName 'FunctionName' -Parameters @{ 'ParameterName' = 'ParameterValue' }" 
    # Write separator for readability
    Write-Host "`n" -BackgroundColor Black
    Write-Host '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++' -BackgroundColor Black -ForegroundColor DarkGray
    # Ask the user which function they want to run
    $selectedFunction = Read-Host -Prompt "`nSelect a function to run by ID, or FunctionName [parameters] (or hit enter to exit):"
    # if the user enters a function name, run it with the provided parameters as a hash table
    if ($selectedFunction -match '(\w+)\s*\[(.*)\]') {
        $functionName = $matches[1]
        $parameters = $matches[2] -split '\s*,\s*' | ForEach-Object {
            $key, $value = $_ -split '\s*=\s*'
            @{ $key = $value }
        }
        Invoke-M365PowerKitFunction -FunctionName $functionName -Parameters $parameters -SkipNestedModuleImport
    }
    elseif ($selectedFunction -match '(\d+)') {
        $selectedFunction = [int]$selectedFunction
        Invoke-M365PowerKitFunction -FunctionName $functionReferences[$selectedFunction] -SkipNestedModuleImport
    }
    elseif ($selectedFunction -eq '') {
        return $null
    }
    else {
        Write-Host 'Invalid selection. Please try again.' -ForegroundColor Red
        Show-M365PowerKitFunctions
    }
    # Ask the user if they want to run another function
    $runAnother = Read-Host -Prompt 'Run another function? (Y / any key to exit)'
    if ($runAnother -eq 'Y') {
        Show-M365PowerKitFunctions
    }
    else {
        Write-Host 'Have a great day!'
        return $null
    }
}

function Use-M365PowerKit {
    param (
        [Parameter(Mandatory = $false)]
        [switch]$InstallDependencies = $false
    )
    Import-NestedModules
    if ($InstallDependencies) {
        Install-Dependencies
    }
    Show-M365PowerKitFunctions
}