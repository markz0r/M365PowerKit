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
# Function display console interface to run any function in the module
function Show-M365PowerKitFunctions {
    # List nested modules and their exported functions to the console in a readable format, grouped by module
    $colors = @('Green', 'Cyan', 'Red', 'Magenta', 'Yellow')
    $nestedModules = Get-Module -Name M365PowerKit | Select-Object -ExpandProperty NestedModules | Where-Object Name -Match 'M365PowerKit-.*'

    $colorIndex = 0
    $functionReferences = @{}
    $nestedModules | ForEach-Object {
        # Select a color from the list
        $color = $colors[$colorIndex % $colors.Count]
        $spaces = ' ' * (52 - $_.Name.Length)
        Write-Host '' -BackgroundColor Black
        Write-Host "Module: $($_.Name)" -BackgroundColor $color -ForegroundColor White -NoNewline
        Write-Host $spaces  -BackgroundColor $color -NoNewline
        Write-Host ' ' -BackgroundColor Black
        $spaces = ' ' * 41
        Write-Host " Exported Commands:$spaces" -BackgroundColor "Dark$color" -ForegroundColor White -NoNewline
        Write-Host ' ' -BackgroundColor Black
        $_.ExportedCommands.Keys | ForEach-Object {
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

    # Write separator for readability
    Write-Host "`n" -BackgroundColor Black
    Write-Host '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++' -BackgroundColor Black -ForegroundColor DarkGray
    # Ask the user which function they want to run
    $selectedFunction = Read-Host -Prompt "`nSelect a function to run (or hit enter to exit):"
    # Attempt to convert the input string to a char
    try {
        $selectedFunction = [int]$selectedFunction
    }
    catch {
        if (!$selectedFunction) {
            return $true
        }
        Write-Host 'Invalid selection. Please try again.'
        Show-M365PowerKitFunctions
    }
    # Run the selected function timing the execution
    Write-Host "`n"
    Write-Host "You selected:  $($functionReferences.$selectedFunction)" -ForegroundColor Green
    try {
        $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        Invoke-Expression ($functionReferences.$selectedFunction)
        $stopwatch.Stop()
    }
    catch {
        # Write all output including errors to the console from the selected function
        Write-Host $_.Exception.Message -ForegroundColor Red
        throw "Error running function: $functionReferences[$selectedFunction] failed. Exiting."
        # Exit with an error code
        exit 1
    }
    finally {
        # Ask the user if they want to run another function
    }   if ($runAnother -eq 'Y') {
        Get-PowerKitFunctions
    }
    else {
        Write-Host 'Have a great day!'
        return $true
    }
}
function Use-M365PowerKit {
    Show-M365PowerKitFunctions
}