# Set exit on error
$ErrorActionPreference = 'Stop'; $DebugPreference = 'Continue'

# Function to import contacts from CSV file to Exchange Online given the CSV file path
function Import-EACContactsFromCSV {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]
        [string]$CsvFilePath,
        [Parameter(Mandatory = $false)]
        [string]$UPN
    )
    if (-not $CsvFilePath) {
        $CsvFilePath = Read-Host 'Enter the path to the CSV file containing the contacts [required]'
    }
    # Test for presence of the CSV file
    if (-not (Test-Path -Path $CsvFilePath)) {
        Write-Error "The CSV file '$CsvFilePath' does not exist"
    }
    if (-not $UPN) {
        $UPN = Read-Host 'Enter the User Principal Name (UPN) of the user running the script (e.g.: admin@onmicrosoft.com) [required]'
    }
    Install-SharedDependencies
    New-IPPSSession -UPN $UPN

    # Import contacts from CSV file
    Import-Csv -Path $CsvFilePath | ForEach-Object {
        $contact = $_
        New-Contact -Name $contact.Name -EmailAddress $contact.EmailAddress -ExternalEmailAddress $contact.ExternalEmailAddress -FirstName $contact.FirstName -LastName $contact.LastName -DisplayName $contact.DisplayName -Department $contact.Department -Title $contact.Title -Office $contact.Office -PhoneNumber $contact.PhoneNumber -MobilePhone $contact.MobilePhone -Fax $contact.Fax -StreetAddress $contact.StreetAddress -City $contact.City -StateOrProvince $contact.StateOrProvince -PostalCode $contact.PostalCode -CountryOrRegion $contact.CountryOrRegion -Notes $contact.Notes -Company $contact.Company -Manager $contact.Manager -Assistant $contact.Assistant -BusinessHomePage $contact.BusinessHomePage -OtherTelephone $contact.OtherTelephone -OtherMobile $contact.OtherMobile -OtherHomePhone $contact.OtherHomePhone -OtherFax $contact.OtherFax -OtherPager $contact.OtherPager -OtherCity $contact.OtherCity -OtherStateOrProvince $contact.OtherStateOrProvince -OtherPostalCode $contact.OtherPostalCode -OtherCountryOrRegion $contact.OtherCountryOrRegion -OtherStreetAddress $contact.OtherStreetAddress -OtherPOBox $contact.OtherPOBox -OtherCompany $contact.OtherCompany -OtherManager $contact.OtherManager -OtherAssistant $contact.OtherAssistant -OtherBusinessHomePage $contact.OtherBusinessHomePage -Initials $contact.Initials -Photo $contact.Photo -UserPrincipalName $contact.UserPrincipalName -CustomAttribute1 $contact.CustomAttribute1 -CustomAttribute2 $contact.CustomAttribute2 -CustomAttribute3 $contact.CustomAttribute3 -CustomAttribute4 $contact.CustomAttribute4 -CustomAttribute5 $contact.CustomAttribute5 -CustomAttribute6 $contact.CustomAttribute6 -CustomAttribute7 $contact.CustomAttribute7 -CustomAttribute8 $contact.CustomAttribute8 -CustomAttribute9 $contact.CustomAttribute9 -CustomAttribute10 $contact.CustomAttribute10 -CustomAttribute11 $contact.CustomAttribute11 -CustomAttribute12 $contact.CustomAttribute12 -CustomAttribute13 $contact.CustomAttribute13 -CustomAttribute14 $contact.Custom        
    }
}