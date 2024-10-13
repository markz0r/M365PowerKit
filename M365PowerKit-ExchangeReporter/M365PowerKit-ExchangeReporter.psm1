<#
.SYNOPSIS
    Report on various Exchange Online Data
.LINK
GitHub: https://github.com/markz0r/M365PowerKit
#>

$ErrorActionPreference = 'Stop'; $DebugPreference = 'Continue'
# Function to Get all SMTP addresses configured for a tenant in Exchange Online
function Get-AllSMTPAddresses {
    Write-Debug 'Running Get-AllSMTPAddresses...'

    # Check $env:
    $RUN_TIMESTAMP = Get-Date -Format 'yyyyMMdd_HHmm'
    $CLEAN_UPN = $($env:M365PowerKitUPN -replace '@', '_').ToLower().Trim()
    $OUTPUT_DIR = $(Get-Location).Path + "\$CLEAN_UPN"
    If (-not (Test-Path -Path $OUTPUT_DIR)) {
        New-Item -Path $OUTPUT_DIR -ItemType Directory -Force
    }
    $OUTPUT_FILE = "$($CLEAN_UPN)_$RUN_TIMESTAMP.json"
    Write-Debug 'Getting all primary SMTP addresses...'
    Write-Output '############# PRIMARY SMTP ADDRESSES #############'
    Get-EXOMailbox | ConvertTo-Json -Depth 100 | Tee-Object -FilePath "$OUTPUT_DIR\Get-EXOMailbox_$OUTPUT_FILE"
    # Write PRIMARY_SMTP_ADDRESSES to console as a formatted table
    #Write-Output $PRIMARY_SMTP_ADDRESSES | Format-Table
    Write-Output '##################################################'
    #Write-Debug 'Getting all SMTP addresses...'
    Write-Output '############# Get-EXORecipient SMTP ADDRESSES #############'
    Get-EXORecipient -ResultSize Unlimited | ConvertTo-Json -Depth 100 | Tee-Object -FilePath "$OUTPUT_DIR\Get-EXORecipient_$OUTPUT_FILE"
    # Write SMTP_ADDRESSES to console as a formatted table
    #Write-Output $SMTP_ADDRESSES | Format-Table 
    Write-Output '##############################################'
    Write-Output '############# Get-Recipient SMTP ADDRESSES #############'
    Get-Recipient -ResultSize Unlimited | ConvertTo-Json -Depth 100 | Tee-Object -FilePath "$OUTPUT_DIR\Get-Recipient_$OUTPUT_FILE"
    Write-Output '##############################################'
    #Get-ChildItem -Filter "*$RUN_TIMESTAMP.json" | ForEach-Object { Write-Output "$($_.Name) -> $($($_ | Get-Content | Select-String -Pattern '[smtp|SMTP]:').Count)" }
    # Create a XLSX file with all SMTP addresses and save it to the current directory, there should be 3 columns one for all each JSON file
    $XLSX_FILE = "$OUTPUT_DIR\$($MyInvocation.MyCommand.Name)-$CLEAN_UPN-$RUN_TIMESTAMP.xlsx"
    Write-Debug "Creating XLSX file: $XLSX_FILE"
    # Declare raw data a PSObject with keys that match the JSON file base names, the values area an array of SMTP addresses
    $COMBINED_DATA = @()
    try {
        Push-Location $OUTPUT_DIR
        Write-Debug "Working in $OUTPUT_DIR"

        Get-ChildItem -Filter "*$OUTPUT_FILE" | ForEach-Object {
            # Create a new object with the KEY and the SMTP_ADDRESS_ARRAY
            Write-Debug "Processing $($_.Name)"
            $FILE = $_
            $KEY = ($FILE.BaseName -split '_')[0]
            # Get the content of the file
            $FILE_DATA = Get-Content $FILE | Select-String -Pattern '(smtp|SMTP)\:' | ForEach-Object {
                [PSCustomObject]@{
                    'Source'       = $KEY
                    'SMTP_Address' = $($_ -replace '.*:(.*)".*$', '$1' )
                }

            } 
            $FILE_DATA | Export-Excel -Path $XLSX_FILE -WorksheetName $KEY -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle Dark8 -ClearSheet
            $COMBINED_DATA += $FILE_DATA
        }
        $COMBINED_DATA | Export-Excel -Path $XLSX_FILE -WorksheetName 'Combined' -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle Dark8 -ClearSheet -MoveToStart
        # For each unique SMTP address, create a [PSCustomObject]@{ 'SMTP_Address' = $SMTP_ADDRESS, 'Sources' = KEYS }
        $UNIQUE_SMTP_ADDRESSES = $COMBINED_DATA | Select-Object -Property 'SMTP_Address' -Unique | ForEach-Object {
            $SMTP_ADDRESS = $_.SMTP_Address
            $SOURCES = $COMBINED_DATA | Where-Object { $_.SMTP_Address -eq $SMTP_ADDRESS } | Select-Object -ExpandProperty 'Source'
            [PSCustomObject]@{
                'SMTP_Address' = $SMTP_ADDRESS
                'Sources'      = $SOURCES -join ', '
            }
        }
        $UNIQUE_SMTP_ADDRESSES | Export-Excel -Path $XLSX_FILE -WorksheetName 'Unique_Addresses' -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle Dark8 -ClearSheet -MoveToStart
    }
    finally {
        # Remove the JSON files
        Get-ChildItem -Filter "*$OUTPUT_FILE" | Remove-Item -Force
        Pop-Location
    }
}