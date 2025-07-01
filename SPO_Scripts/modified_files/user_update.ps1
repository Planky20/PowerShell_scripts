############################################
# Author
# Adam C - V1.0
# Wladyslaw L - V2.0

# Requirments
# AD      powershell module
# SPOPNP  powershell module
# MSGraph powershell module

# For mailing to be working correctly - Microsoft.Graph.Core v.3.1.13 is required!!!

# Azure app registrations
# app reg 1 - app reg with ClientId1 and ClientSecret for SharePoint Online connection
# app reg 2 - app reg with ClientId2 and ClientCert for Microsoft Graph connection

# Modules
# Install-Module -Name "PnP.PowerShell"
# Install-Module -Name "Microsoft.Graph"
############################################


#region Functions definition

#region Function to convert date to string
function date-string () {
    Param(
        [DateTime]$Data 
    )

    $DateStr = $Data.ToString("dd\/MM\/yyyy")
    return $DateStr

}
#endregion

#region Function to check if record is correct and add it to array
function check-record() {
    param(
        [string]$key,
        [string]$value
    )

    [bool]$result = $false

    if (($value -eq $key)) {
        $result = $true
    }
    elseif (($key -eq "Error")) {
    }
    return $result
}
#endregion

#region Function to add national local group
function add-national-local-group() {
    Param(
        [string]$country
    )

    switch ($country) {
        "POLAND"         { $group = "Domain PL ALL" }
        "UNITED KINGDOM" { $group = "Domain UK ALL" }
        "SWITZERLAND"    { $group = "Domain CH ALL" }
    }

    return $group
}
#endregion

#region Function to match country to country code
function match-country () {
    param(
        [string]$country
    )
    
    switch ($country) {
        "POLAND"         { $result = "PL" }
        "AUSTRIA"        { $result = "AT" }
        "UNITED KINGDOM" { $result = "GB" }
        "SWITZERLAND"    { $result = "CH" }
        "ROMANIA"        { $result = "RO" }
    }
    
    return $result
}
#endregion

#region Function to match country to DC OU path
function match-country-DC () {
    param(
        [string]$country,
        [string]$city
    )
    $city = $city.Replace(" 2", "")
    if ($country -like "SWITZERLAND") {
        $result = 'OU=' + $city + ',OU=' + $country + ',OU=Users,OU=ADSync,DC=your,DC=domain,DC=FQDN'
    }
    else {
        $result = 'OU=' + $city + ',OU=' + $country + ',OU=Users,OU=CloudSync,DC=your,DC=domain,DC=FQDN'
    }

    return $result
}
#endregion

#region Function to match city to street address
function match-address () {
    param(
        [string]$city
    )
    switch ($city) {
        "Warsaw"    { $result = "some street address for Warsaw"    }
        "Lublin"    { $result = "some street address for Lublin"    }
        "Lodz"      { $result = "some street address for Lodz"      }
        "Poznan"    { $result = "some street address for Poznan"    }
        "Katowice"  { $result = "some street address for Katowice"  }
        "Rzeszow"   { $result = "some street address for Rzeszow"   }
        "Zurich"    { $result = "some street address for Zurich"    }
        "Geneva"    { $result = "some street address for Geneva"    }
        "Kingston"  { $result = "some street address for Kingston"  }
        "Vienna"    { $result = "some street address for Vienna"    }
        "Bucharest" { $result = "some street address for Bucharest" }
        "Remote"    { $result = "Remote" }
    }
    
    return $result
}
#endregion

#region Function to match city to postal code
function match-address-code () {
    param(
        [string]$city
    )
    
    switch ($city) {
        "Warsaw"    { $result = "Warsaw postal code"    }
        "Lublin"    { $result = "Lublin postal code"    }
        "Lodz"      { $result = "Lodz postal code"      }
        "Poznan"    { $result = "Poznan postal code"    }
        "Rzeszow"   { $result = "Rzeszow postal code"   }
        "Katowice"  { $result = "Katowice postal code"  }
        "Zurich"    { $result = "Zurich postal code"    }
        "Geneva"    { $result = "Geneva postal code"    }
        "Kingston"  { $result = "Kingston postal code"  }
        "Vienna"    { $result = "Vienna postal code"    }
        "Bucharest" { $result = "Bucharest postal code" }
        "Remote"    { $result = " " }
    }

    return $result
}
#endregion

#region Function to get city
function get-city () {
    param(
        [string]$city
    )
    $result = $city.Replace(" 2", "")
    return $result
}
#endregion

#region Function to remove Polish characters from name and surname
function Remove-StringLatinCharacters {
    PARAM ([string]$String)
    [Text.Encoding]::ASCII.GetString([Text.Encoding]::GetEncoding("Cyrillic").GetBytes($String))
}
#endregion

#region Function to check if User Principal Name is unique
function check-UPN () {
    param(
        [string]$UPN,
        [string]$domain
    )
    $test = $UPN
    $prefix = $UPN.Remove(($UPN.indexof("@")), ($UPN.Length - ($UPN.indexof("@"))))
    $i = 1
    do {   
        $tmp = Get-ADUser -Filter { UserPrincipalName -eq $test }
        if ($tmp) {
            $probe = $true
            $test = $prefix + $i + "@" + $domain
            $i++
        }
        else {
            if ($i -gt 1) { $test = $prefix + $i + "@" + $domain }
            else { $test = $prefix + "@" + $domain }
            $probe = $false
        }
    } while ($probe -eq $true )
    $results = $test
    return $results
}
#endregion

#region Function to get manager name
function get-manager-name () {
    Param(
        [Microsoft.Sharepoint.Client.FieldUserValue]$Object
    )
    if (!$Object.Email) {
        $test = $Object.LookupValue
        $resultsx = (Get-ADUser -Filter { Name -eq $test }).SamAccountName
        $results = $resultsx
    }
    else {
        $test = $Object.Email
        $reversecheck = $Object.Email
        if ($reversecheck) { $new_main_alias = ($reversecheck.Remove($reversecheck.IndexOf("@"), $reversecheck.Length - $reversecheck.IndexOf("@")) + "@domain.com") }
        else { $new_main_alias = $reversecheck }
        $results = (Get-ADUser -Filter { UserPrincipalName -eq $new_main_alias }).SamAccountName
    }
    if (!$results) {
        $test = $Object.LookupValue
        $resultsx = (Get-ADUser -Filter { Name -eq $test }).SamAccountName
        $results = $resultsx
    }
    return $results
}
#endregion

#region Function to get techleader name and add it to pager attribute
function get-techleader-name () {
    Param(
        [Microsoft.Sharepoint.Client.FieldUserValue]$Object
    )
    $reversecheck = $Object.Email
    if ($reversecheck) {}

    if (!$Object.Email) {
        $test = $Object.LookupValue
        $results = (Get-ADUser -Filter { Name -eq $test }).SamAccountName
    }
    else {
        $test = $Object.Email
        $reversecheck = $Object.Email
        if ($reversecheck) { $new_main_alias = ($reversecheck.Remove($reversecheck.IndexOf("@"), $reversecheck.Length - $reversecheck.IndexOf("@")) + "@domain.com") }
        else { $new_main_alias = $reversecheck }
        $results = (Get-ADUser -Filter { UserPrincipalName -eq $new_main_alias }).userprincipalname
    }
    return $results
}
#endregion

#region Function to get MPK
function get-mpk () {
    Param(
        [string]$Department
    )
    switch ($Department) {
        
        "Dept1"  { $MPK = "100001" }
        "Dept2"  { $MPK = "100002" }
        "Dept3"  { $MPK = "100003" }
        "Dept4"  { $MPK = "100004" }
        "Dept5"  { $MPK = "100005" }
        "Dept6"  { $MPK = "100006" }
        "Dept7"  { $MPK = "100007" }
        "Dept8"  { $MPK = "100008" }
        "Dept9"  { $MPK = "100009" }
        "Dept10" { $MPK = "100010" }
        default  { $MPK = "100000" }
    }

    return $MPK
}
#endregion

#region Function to get sex
function get-sex () {
    Param(
        [string]$Sex
    )
    switch ($Sex) {
        "Female" { $results = "F" }
        "Male"   { $results = "M" }
    }

    return $results
}
#endregion

#region Function to check if record is different
function check-record2 () {
    Param(
        [Parameter(mandatory = $false)]
        [string]$OldValue,
        [string]$NewValue
    )
    if ($NewValue -ne $OldValue) {
        $results = $true
    }
    else { $results = $false }
    return $results 
}
#endregion

#region Function to change record
function change-record () {
    Param(
        [Parameter(mandatory = $false)]
        [object]$Value,
        [string]$Attribute,
        [string]$Account,
        [string]$ID
    )
    switch ($Attribute) {
        Title               { if ((check-record2 -OldValue ((get-aduser -Identity $Account -Properties *).surname) -NewValue (Remove-StringLatinCharacters -String $Value))) { $surname = (Remove-StringLatinCharacters -String $Value); Set-ADUser -Identity $Account -Surname $surname ; $results = $Attribute } }
        Nazwisko            { if ((check-record2 -OldValue ((get-aduser -Identity $Account -Properties *).givenname) -NewValue (Remove-StringLatinCharacters -String $Value))) { $givenname = (Remove-StringLatinCharacters -String $Value); Set-ADUser -Identity $Account -GivenName $Value ; $results = $Attribute } }
        Sex                 { if ((check-record2 -OldValue ((get-aduser -Identity $Account -Properties *).extensionAttribute5) -NewValue $Value)) { Set-ADUser -Identity $Account -Replace @{ extensionAttribute5 = (get-sex -Sex $Value) }; $results = $Attribute } }
        Licences            { if ((check-record2 -OldValue ((get-aduser -Identity $Account -Properties *).extensionAttribute2) -NewValue $Value)) { Set-ADUser -Identity $Account -Replace @{ extensionAttribute2 = $Value }; $results = $Attribute } }
        Contract_x0020_type { if ((check-record2 -OldValue ((get-aduser -Identity $Account -Properties *).extensionAttribute7) -NewValue $Value)) { Set-ADUser -Identity $Account -Replace @{ extensionAttribute7 = $Value }; $results = $Attribute } }
        Job_x0020_Title     { if ((check-record2 -OldValue ((get-aduser -Identity $Account -Properties *).Description) -NewValue $Value)) { Set-ADUser -Identity $Account -Description $Value -Title $Value ; $results = $Attribute } }
        Stanowisko          { if ((check-record2 -OldValue ((get-aduser -Identity $Account -Properties *).extensionAttribute3) -NewValue $Value)) { Set-ADUser -Identity $Account -Replace @{ extensionAttribute3 = $Value }; $results = $Attribute } }
        Poziom              { if ((check-record2 -OldValue ((get-aduser -Identity $Account -Properties *).extensionAttribute4) -NewValue $Value)) { Set-ADUser -Identity $Account -Replace @{ extensionAttribute4 = $Value }; $results = $Attribute } }
        Manager_x0020_Name  { if ((check-record2 -OldValue ((get-aduser -Identity $Account -Properties *).manager) -NewValue $Value)) { Set-ADUser -Identity $Account -Manager (get-manager-name -Object $Value); $results = $Attribute } }
        Dzia_x0142_         { Set-ADUser -Identity $Account -Department $Value ; $results = $Attribute }
        Dzia_x0142_         { Set-ADUser -Identity $Account -Replace @{extensionAttribute12 = (get-mpk -Department $Value) } ; $results = $Attribute }
        Kraj                { if ((check-record2 -OldValue ((get-aduser -Identity $Account -Properties *).Country) -NewValue $Value)) { Set-ADUser -Identity $Account -Country (match-country -country $Value); $results = $Attribute } }
        Lokalizacja         { Set-ADUser -Identity $Account -StreetAddress (match-address -city $Value ) -PostalCode (match-address-code -city $Value ) -Replace @{extensionAttribute6 = $Value } -City (get-city -city $Value) -Office (get-city -city $Value) ; $results = $Attribute }
        AD_login            { if ((check-record2 -OldValue ((get-aduser -Identity $Account -Properties *).SamAccountName) -NewValue $Value)) { Set-ADUser -Identity $Account -SamAccountName $Value; $results = $Attribute } }
        Email               { if ((check-record2 -OldValue ((get-aduser -Identity $Account -Properties *).UserPrincipalName) -NewValue $Value)) { Set-ADUser -Identity $Account -UserPrincipalName (check-UPN -UPN $Value -domain "domain.com") ; $results = $Attribute } }
        NRIDNEW             { if ((check-record2 -OldValue ((get-aduser -Identity $Account -Properties *).extensionAttribute9) -NewValue $Value)) { Set-ADUser -Identity $Account -Replace @{ extensionAttribute9 = $Value }; $results = $Attribute } }
        NRIDNEW             { if ((check-record2 -OldValue ((get-aduser -Identity $Account -Properties *).extensionAttribute11) -NewValue $Value)) { Set-ADUser -Identity $Account -Replace @{ extensionAttribute11 = $Value }; $results = $Attribute } }
        JobTimeSize         { if ((check-record2 -OldValue ((get-aduser -Identity $Account -Properties *).extensionAttribute1) -NewValue $Value)) { Set-ADUser -Identity $Account -Replace @{ extensionAttribute1 = $Value }; $results = $Attribute } }
        Billable            { if ((check-record2 -OldValue ((get-aduser -Identity $Account -Properties *).extensionAttribute14) -NewValue $Value)) { Set-ADUser -Identity $Account -Replace @{ extensionAttribute14 = $Value }; $results = $Attribute } }    
    }

    return $results
}
#endregion

#region Function to get company
function get-company () {
    Param(
        [string]$ContractType,
        [string]$Company
    )
    switch ($ContractType) {
        "External"  { $result = "External"; break      }
        "Partner"   { $result = "Partner"; break       }
        default     { $result = "Default Value"; break }
    }
    if ($Company -eq "Other") {
        $result = "Other Sp. z o.o."
    }

    return $result
}
#endregion

#region Function to check if string contains white signs
function check-white-signs () {
    Param(
        [string]$String
    )
    $result = $false
    if ($String.IndexOfAny(" ") -eq (-1))
    { $result = $true }
    return $result
}
#endregion
#endregion Functions definition

#region Variables definition
$SiteUrl        = "SP site URL"
$listName       = "SP list name"
$tenantID       = "tenant ID"
$ClientId1      = "app reg 1 ID"
$ClientId2      = "app reg 2 ID"
$KeyVaultName   = "KV name"
$company        = "company name"
$from           = "sender email address"
$To             = "recipient email address"
$SubjectSuccess = "Update was completed with success"
$attributes     = "Title", "Nazwisko", "Contract_x0020_type", "Job_x0020_Title", "Stanowisko", "Poziom", "Dzia_x0142_", "Kraj", "Lokalizacja", "Company", "Licences", "JobTimeSize", "Billable", "TechLeader", "Technology", "NRIDNEW", "Manager_x0020_Name"
$date_ad        = Get-Date -Format "dd/MM/yyyy HH:mm"
#endregion Variables definition

#region Environment connection
$Response = Invoke-RestMethod -Uri 'http://169.254.169.254/metadata/identity/oauth2/token?api-version=2018-02-01&resource=https%3A%2F%2Fvault.azure.net' -Method GET -Headers @{Metadata="true"}
$KeyVaultToken = $Response.access_token
$ClientCert = (Invoke-RestMethod -Uri https://$KeyVaultName.vault.azure.net/secrets/<CERTNAME>?api-version=2016-10-01 -Method GET -Headers @{Authorization="Bearer $KeyVaultToken"}).value
$ClientSecret = (Invoke-RestMethod -Uri https://$KeyVaultName.vault.azure.net/secrets/<SECRETNAME>?api-version=2016-10-01 -Method GET -Headers @{Authorization="Bearer $KeyVaultToken"}).value
Import-Module Microsoft.Graph.Authentication
Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId1 -ClientSecret $ClientSecret
Connect-MgGraph -ClientId $ClientId2 -CertificateThumbprint $ClientCert -TenantId $tenantID
#endregion Environment connection

#region Main Code
$List = Get-PnPList -Identity $listName
$records_spos = (Get-PnPListItem -List $List | Select-Object id, @{label = "Filename"; expression = { $_.FieldValues } }).filename
$records_spos.Count

$Lists = @()
foreach ($records_spo in $records_spos) {
    if (check-record -value ($records_spo.Ready_x0020_for_x0020_deploy) -key "True") {
        if (check-record -value $records_spo.Zmiana -key "Update" ) {
            if (check-record -value $records_spo.Disabled -key "No" ) {
                if (check-record -value $records_spo.Deployed -key "Yes" ) {
                    $Lists += $records_spo
                }
            }
        }
    }
}

$Lists.Count

if ($SAM) {
    Remove-Variable SAM 
}
if ($UPN) {
    Remove-Variable UPN 
}
if ($record) {
    Remove-Variable record 
}

foreach ($record in $lists) {
    $User_created               = $record.Title + " " + $record.Nazwisko                      # string
    $Title                      = Remove-StringLatinCharacters -String $record.Title          # string
    $Nazwisko                   = Remove-StringLatinCharacters -String $record.Nazwisko       # string
    $Sex                        = $record.Sex                                                 # string
    $Od_x0020_kiedy             = $record.Od_x0020_kiedy                                      # date
    $NRIDNEW                    = $record.NRIDNEW                                             # string
    $Contract_x0020_type        = $record.Contract_x0020_type                                 # string
    $Job_x0020_title            = $record.Job_x0020_Title                                     # string
    $Stanowisko                 = $record.Stanowisko                                          # string
    $Poziom                     = $record.Poziom                                              # string
    $Manager_x0020_Name         = $record.Manager_x0020_Name                                  # FieldUserValue
    $Dzia_x0142_                = $record.Dzia_x0142_                                         # string
    $Kraj                       = $record.Kraj                                                # string
    $Lokalizacja                = $record.Lokalizacja                                         # string
    $Dodatkowe_x0020_informacje = $record.Dodatkowe_x0020_Informacje                          # string
    $Ad_login                   = $record.AD_login                                            # string
    $Email                      = $record.Email                                               # string
    $Po                         = $record.PO                                                  # string
    $Company                    = $record.Company                                             # string
    $Licences                   = $record.Licences                                            # string
    $ID                         = $record.ID                                                  # string
    $Manager                    = $Manager_x0020_Name.LookupValue                             # string
    $JobTimeSize                = $record.JobTimeSize                                         # string
    $Billable                   = $record.Billable                                            # string
    $TechLeader                 = $record.TechLeader                                          # FieldUserValue
    $Technology                 = $record.Technology                                          # string
    $CompanyCode0               = $record.CompanyCode0                                        # string
    $ProfitCenter               = $record.PO                                                  # string

    if (check-white-signs -String $Title)    { $Error.Add("White signs detected in variable = Title")   }
    if (check-white-signs -String $Nazwisko) { $Error.Add("White signs detected in variable = Surname") }
    if (check-white-signs -String $Ad_login) { $Error.Add("White signs detected in variable = Adlogin") }
    if (check-white-signs -String $Email)    { $Error.Add("White signs detected in variable = Email")   }

    foreach ($attribute in $attributes) {
        change-record -Value $record.$attribute -Attribute $attribute -Account $Ad_login -ID $ID
    }
    Add-ADGroupMember -Identity (add-national-local-group -country $Kraj) -Members $Ad_login

    if ($TechLeader) { Set-ADUser -Identity $Ad_login -Replace @{ pager = (get-techleader-name -Object $TechLeader) } }
    else { Set-ADUser -Identity $Ad_login -Clear pager }

    if ($Technology) { Set-ADUser -Identity $Ad_login -Replace @{ personalPager = $Technology } }
    else { Set-ADUser -Identity $Ad_login -Clear personalPager }

    Set-ADUser -Identity $Ad_login -Add @{carlicense = (date-string -Data $Od_x0020_kiedy) }
    Set-ADUser -Identity $Ad_login -Replace @{extensionAttribute12 = $ProfitCenter }
    Set-ADUser -Identity $Ad_login -Replace @{comment = $CompanyCode0 }
    Set-ADUser -Identity $Ad_login -Department $Dzia_x0142_
    Set-ADUser -Identity $Ad_login -Company (get-company -ContractType $Contract_x0020_type -Company $company )
    $User_DN = (get-aduser -Identity $Ad_login).DistinguishedName
    Move-ADObject -Identity $User_DN -TargetPath (match-country-DC -country $Kraj -city $Lokalizacja )

    if ($Dzia_x0142_ -eq "Talent Acquisition" -or $Dzia_x0142_ -eq "Talent Management" ) {
        Add-ADGroupMember -Identity 'Group Name' -Members $SamAccountName 
    }
    #endregion Main Code

    #region Email notification
    $BodySuccess = "Dear Support Team,

Please be informed that account was updated with success ( $User_created ). 

Basic information

UPN Name:     $Email
SAM Name:     $Ad_login
ID number:    $NR_x0020_ID
Contract:     $Contract_x0020_type
Job EN:       $Job_x0020_title
Job PL:       $Stanowisko
Level:        $Poziom
Manager:      $Manager
Department:   $Dzia_x0142_
Country:      $Kraj
Location:     $Lokalizacja
Po:           $Po
Company:      $Company
Licence ID:   $Licences
SP ID:        $ID


Best Regards

Orchestrator
"
    $params = @{
        Message         = @{
            Subject      = $SubjectSuccess
            Body         = @{
                ContentType = "Text"
                Content     = $BodySuccess
            }
            ToRecipients = @(
                @{
                    EmailAddress = @{
                        Address = $To
                    }
                }
            )
        }
        SaveToSentItems = "false"
    }
    Send-MgUserMail -UserId $from -BodyParameter $params
    Set-PnPListItem -List $List -Identity $ID -Values @{"Dodatkowe_x0020_Informacje" = "$Dodatkowe_x0020_informacje; User update at $date_ad" ; "Zmiana" = "Auto-Update" }
}
#endregion Email notification