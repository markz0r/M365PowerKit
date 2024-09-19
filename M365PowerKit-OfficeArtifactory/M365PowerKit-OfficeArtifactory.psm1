<#
.SYNOPSIS
This module contains functions for:
 - Creating docx files from Markdown files
 - Applying a templates to docx files
 - Converting docx files to PDF
    
.DESCRIPTION
Used as part of the M365PowerKit, this module contains functions for:
 - Creating docx files from Markdown files
 - Applying a templates to docx files
 - Converting docx files to PDF

.PARAMETER SoureFile
The source file to be converted.

.PARAMETER DestinationFile
The destination file to be saved.

.PARAMETER TemplateFile
The template file to be applied.

.PARAMETER ArtifactOp
The operation to be performed on the artifact.
 saved.

.PARAMETER DisableDebug
Disables debug output.

.PARAMETER InstallDepsOnly
Installs the required dependencies only.

.PARAMETER SkipModules
Skips importing the required modules.


.EXAMPLE
# To install the module, run the following commands:

$PARAM_HASH = @{
    SourceFile = "C:\temp\example.md"
    DestinationFile = "C:\temp\example.docx"
    TemplateFile = "C:\temp\template.docx"
    ArtifactOp = "Convert-MarkdownToDocxWithTemplate"
}

Import-Module M365PowerKit -Force
# Install-M365PowerKitDependencies -InstallDepsOnly
Invoke-M365PowerKitFunction -FunctionName "Invoke-OfficeArtifactConversion" -FunctionParameters $PARAM_HASH


.LINK
GitHub: https://github.com/microsoft/M365PowerKit
#>

$ErrorActionPreference = 'Stop'; $DebugPreference = 'Continue'
# Start transcript logging
$TranscriptPath = "$PSScriptRoot\Trans\$(Get-Date -Format 'yyyyMMdd_hhmmss')-Transcript.log"

$DEFAULT_DOCX_TEMPLATE = "$env:USERPROFILE\B42\Zoak Solutions - Documents\Zoak_Document_Template.dotx"

function Validate-DocxFile {
    param (
        [Parameter(Mandatory = $true)]
        [string]$DocxFile
    )
    if (-not (Test-Path $DocxFile)) {
        throw "Docx file not found: $DocxFile"
    } else {
        Write-Output "Docx file found: $DocxFile"
        # Attempt to open the docx file
        $Word = New-Object -ComObject Word.Application
        $Word.Visible = $false
        # Ensure we print any errors to the console as Debug
        try {
            Write-Debug "Opening docx file: $DocxFile"
            $Doc = $Word.Documents.Open($DocxFile)
            Write-Debug "Docx file opened: $DocxFile"
        } catch {
            Write-Debug "Error opening docx file: $DocxFile"
            Write-Debug $($_ | Format-List | Out-String)
            Write-Error "Error opening docx file: $DocxFile"
        }
        finally {
            $Word.Quit()
        }
    }
}

function Convert-MarkdownToDocxWithTemplate {
    param (
        [Parameter(Mandatory = $true)]
        [string]$SourceFile,
        [Parameter(Mandatory = $false)]
        [string]$DestinationFile,
        [Parameter(Mandatory = $false)]
        [string]$TemplateFile
    )
    # Check if the source file exists and is a markdown
    if (-not (Test-Path $SourceFile)) {
        throw "Source file not found: $SourceFile"
    } 
    if (-not $TemplateFile) {
        # Default template to USERPROFILE\B42\Zoak Solutions - Documents
        Write-Debug "Using default template: $DEFAULT_DOCX_TEMPLATE"
        $TemplateFile = $DEFAULT_DOCX_TEMPLATE

    }
    if (-not $DestinationFile) {
        $DestinationFile = $SourceFile -replace '\.md$', '.docx'
    }
    # If destination file exists, ask user to confirm overwrite
    if (Test-Path $DestinationFile) {
        $Confirm = Read-Host "Destination file already exists. Overwrite? (Y/N)"
        if ($Confirm -ne 'Y') {
            throw "Destination file already exists: $DestinationFile"
        }
    }
    # First convert markdown to HTML
    $HTMLFile = $SourceFile -replace '\.md$', '.html'
    # If ConvertFrom-Markdown is not available, install the required module
    ConvertFrom-Markdown -Path $SourceFile -DestinationPath $HTMLFile

    [ref]$SaveFormat = "microsoft.office.interop.word.WdSaveFormat" -as [type] 
    $Word = New-Object -ComObject Word.Application
    $Word.Visible = $false
    try {
        # Convert HTML to docx
        $Doc = $Word.Documents.Open($HTMLFile)
        $Doc.saveas([ref] $docx, [ref]$SaveFormat::wdFormatDocumentDefault) 
        # Apply template
        $Doc.AttachedTemplate = $TemplateFile
        # Save after applying template
        $Doc.SaveAs($DestinationFile)
        Write-Output "Docx file saved: $DestinationFile"
        $Doc.Close()
    }
    finally {
        $Word.Quit()
    }
    Write-Debug "Validating docx file: $DestinationFile"
    Validate-DocxFile -DocxFile $DestinationFile
}

function Invoke-OfficeArtifactory {
    param (
        [Parameter(Mandatory = $false)]
        [string]$SourceFile,
        [Parameter(Mandatory = $false)]
        [string]$DestinationFile,
        [Parameter(Mandatory = $false)]
        [string]$TemplateFile = $DEFAULT_DOCX_TEMPLATE,
        [Parameter(Mandatory = $true)]
        [string]$ArtifactOp,
        [Parameter(Mandatory = $false)]
        [switch]$DisableDebug
    )
    # Start transcript logging
    $TranscriptPath = "$PSScriptRoot\Trans\$(Get-Date -Format 'yyyyMMdd_hhmmss')-Transcript.log"

    if ($DisableDebug) {
        $DebugPreference = 'SilentlyContinue'
    }

    if ($ArtifactOp -eq "Convert-MarkdownToDocxWithTemplate") {
        Convert-MarkdownToDocxWithTemplate -SourceFile $SourceFile -DestinationFile $DestinationFile -TemplateFile $TemplateFile
    }
    elseif ($ArtifactOp -eq "Convert-DocxToPdf") {
        Convert-DocxToPdf -SourceFile $SourceFile -DestinationFile $DestinationFile
    }
    else {
        throw "Invalid operation: $ArtifactOp"
    }
}
