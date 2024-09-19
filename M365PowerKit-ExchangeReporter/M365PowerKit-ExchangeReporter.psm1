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
    $OUTPUT_FILE = "$($MyInvocation.MyCommand.Name)_$UPN_$RUN_TIMESTAMP.json"
    Write-Debug 'Getting all primary SMTP addresses...'
    Write-Output '############# PRIMARY SMTP ADDRESSES #############'
    Get-EXOMailbox -ResultSize Unlimited | ConvertTo-Json -Depth 100 | Tee-Object -FilePath "Get-EXOMailbox_$OUTPUT_FILE"
    # Write PRIMARY_SMTP_ADDRESSES to console as a formatted table
    #Write-Output $PRIMARY_SMTP_ADDRESSES | Format-Table
    Write-Output '##################################################'
    #Write-Debug 'Getting all SMTP addresses...'
    Write-Output '############# Get-EXORecipient SMTP ADDRESSES #############'
    Get-EXORecipient -ResultSize Unlimited | ConvertTo-Json -Depth 100 | Tee-Object -FilePath "Get-EXOReceipient_$OUTPUT_FILE"
    # Write SMTP_ADDRESSES to console as a formatted table
    #Write-Output $SMTP_ADDRESSES | Format-Table 
    Write-Output '##############################################'
    Write-Output '############# Get-Recipient SMTP ADDRESSES #############'
    Get-Recipient -ResultSize Unlimited | ConvertTo-Json -Depth 100 | Tee-Object -FilePath "Get-Receipient_$OUTPUT_FILE"
    Write-Output '##############################################'
    Get-ChildItem -Filter "*$RUN_TIMESTAMP.json" | ForEach-Object { Write-Output "$($_.Name) -> $($($_ | Get-Content | Select-String -Pattern '[smtp|SMTP]:').Count)" }
    # Create a XLSX file with all SMTP addresses and save it to the current directory, there should be 3 columns one for all each JSON file
    $XLSX_FILE = "$($MyInvocation.MyCommand.Name)_$UPN_$RUN_TIMESTAMP.xlsx"
    # Declare raw data a PSObject with keys that match the JSON file base names, the values area an array of SMTP addresses

    # Loop through each JSON file

    $COMBINED_DATA = $(Get-ChildItem -Filter "*$RUN_TIMESTAMP.json" | ForEach-Object {
            # Create a new object with the KEY and the SMTP_ADDRESS_ARRAY
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
            $COMBINED_OBJECT = [PSCustomObject]@{
                $KEY = $FILE_DATA.SMTP_Address
            }
            $COMBINED_OBJECT
        })
    Write-Debug $COMBINED_DATA.GetType()
    Write-Debug $COMBINED_DATA.Count
    
    # BUILD a json object like:
    #  '{
    #     "Get-EXOMailbox": [
    #     "info@zoak.com.au",
    #     "mark.culhane7531@zoaksolutions.onmicrosoft.com"
    #     ],
    #     "Get-EXOReceipient": ["info@zoak.com.au", "info@zoak.solutions"],
    #     "Get-Receipient": [
    #     "DiscoverySearchMailbox{D919BA05-46A6-415f-80AD-7E09334BB852}@zoaksolutions.onmicrosoft.com",
    #     "mark.culhane.ssg@zoak.solutions",
    #     "mark.culhane.rm@zoak.solutions"
    #     ]
    # }'

    $COMBINED_JSON = '{' 
    $COMBINED_DATA | ForEach-Object {
        Write-Debug "PSObject: $($_) - Properties: $($_.PSObject.Properties)"
        $_.PSObject.Properties | ForEach-Object {
            $COMBINED_JSON = $COMBINED_JSON + '"{0}": {1},' -f $_.Name, ($_.Value | ConvertTo-Json -Depth 5 -Compress | ForEach-Object { $_ -replace '(\{|\})', '' })
        } }
    $COMBINED_JSON = $COMBINED_JSON + '}'
    # Remove the trailing comma
    $COMBINED_JSON = $COMBINED_JSON -replace '\],\}$', ']}'
    # Validate the JSON
    $COMBINED_JSON | ConvertFrom-Json
    $COMBINED_JSON | ConvertFrom-Json | ConvertTo-Csv


    #Write-Debug $COMBINED_JSON
    $COMBINED_JSON | ConvertFrom-Json | Export-Excel -Path $XLSX_FILE -WorksheetName 'Combined' -AutoSize -FreezeTopRow -BoldTopRow -Show  -MoveToStart
}
