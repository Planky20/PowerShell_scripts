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

#region Functions to convert date to string
function date-string () {
    Param(
        [DateTime]$Data 
    )

    $DateStr = $Data.ToString("dd\/MM\/yyyy")
    return $DateStr

}
#endregion

#region Functions to check if record is correct and add it to array
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
    $result = 'OU=' + $city + ',OU=' + $country + ',OU=Users,OU=CloudSync,DC=your,DC=domain,DC=FQDN'

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

#region Function to get city name
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

#region Function to build SAM account name (short login)
function Build-SAM () {
    param(
        [string]$Givenname,
        [string]$Surname
    )

    if ($Surname.Length -ge 3) {
        $Givenname = $Givenname.ToLower()
        $results = $Givenname[0] + ($Surname.remove(3, ($Surname.Length - 3))).ToLower()
    }
    else {
        $Givenname = $Givenname.ToLower()
        $results = $Givenname[0] + $Surname.ToLower()
    }
    return $results
}
#endregion

#region Function to check if SAM account name is unique
function check-SAM () {
    param(
        [string]$SAM
    )
    $i = 1
    $test = $SAM 
    do {   
        $tmp = Get-ADUser -Filter { SamAccountName -eq $test }
        if ($tmp) {
            $probe = $true
            $test = $SAM + $i
            $i++
        }
        else {
            if ($i -gt 1) { $test = $SAM + $i }
            else { $test = $SAM }
            $probe = $false
        }
    } while ($probe -eq $true )
    $results = $test

    return $results
}
#endregion

#region Function to build User Principal Name (email)
function Build-UPN () {
    param(
        [string]$Givenname,
        [string]$Surname,
        [string]$domain
    )

    $Givenname = $Givenname.ToLower()
    $Surname = $Surname.ToLower()
    $results = $Givenname + "." + $Surname + "@" + $domain

    return $results
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

#region Function to change string to bool
function Change-tobool() {
    Param(
        [Parameter(mandatory = $true)]
        [string]$String
    )

    [bool]$result = $false
    switch ($String) {
        "yes"   { $result = $true  }
        "no"    { $result = $false }
        "TRUE"  { $result = $true  }
        "FALSE" { $result = $false }
    }
    return $result
}
#endregion

#region Function to generate password for returning user
function password-generator() {
    Param(
        [int]$Strong = 10
    )
    $Password = "!@#$%^&*0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz".tochararray() 
    $Password1 = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ_abcdefghijklmnopqrstuvwxyz".tochararray() 
    $Password2 = "!@#$%^&*".tochararray() 
    $prefix = (($Password1 | Get-Random -count 2 ) -join '')
    $result = (($Password | Get-Random -count $Strong ) -join '')
    $sufix = (($Password2 | Get-Random -count 1 ) -join '')
    $results = (( $sufix + $result | Get-Random -count ($Strong + 1)) -join '')
    return $prefix + $results
}
#endregion

#region Function to get manager name
function get-manager-name () {
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
        $results = (Get-ADUser -Filter { UserPrincipalName -eq $new_main_alias }).SamAccountName
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

#region Function to add national sufix to proxy addresses
    function add-national-sufix() {
    Param(
        [string]$country,
        [string]$UserPrincipalName
    )

    switch ($country) {
        "POLAND"          { $sufix = "@domain.pl"    }
        "AUSTRIA"         { $sufix = "@domain.at"    }
        "UNITED KINGDOM"  { $sufix = "@domain.co.uk" }
        "SWITZERLAND"     { $sufix = "@domain.ch"    }
        "ROMANIA"         { $sufix = "@domain.ro"    }
    }
    $prefix  = ($UserPrincipalName.Remove($UserPrincipalName.IndexOf("@"), ($UserPrincipalName.Length - $UserPrincipalName.IndexOf("@"))))
    $results = $prefix + $sufix

    return $results
}
#endregion

#region Function to define sex
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

#region Function to check if strings with name and surname are without white signs
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
$SubjectSuccess = "Enabling account was completed with success"
$date_ad        = Get-Date -Format "dd/MM/yyyy HH:mm"
#endregion Variables definition

#region Environment connection
$Response = Invoke-RestMethod -Uri 'http://169.254.169.254/metadata/identity/oauth2/token?api-version=2018-02-01&resource=https%3A%2F%2Fvault.azure.net' -Method GET -Headers @{Metadata = "true" }
$KeyVaultToken = $Response.access_token
$ClientCert = (Invoke-RestMethod -Uri https://$KeyVaultName.vault.azure.net/secrets/<CERTNAME>?api-version=2016-10-01 -Method GET -Headers @{Authorization = "Bearer $KeyVaultToken" }).value
$ClientSecret = (Invoke-RestMethod -Uri https://$KeyVaultName.vault.azure.net/secrets/<SECRETNAME>?api-version=2016-10-01 -Method GET -Headers @{Authorization = "Bearer $KeyVaultToken" }).value
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
        if (check-record -value $records_spo.Zmiana -key "Start over" ) {
            if (check-record -value $records_spo.Disabled -key "Yes" ) {
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

foreach ($record in $Lists) {


    $User_created               = $record.Title + " " + $record.Nazwisko                   # string
    $Title                      = Remove-StringLatinCharacters -String $record.Title       # string
    $Nazwisko                   = Remove-StringLatinCharacters -String $record.Nazwisko    # string
    $Sex                        = $record.Sex                                              # string
    $Od_x0020_kiedy             = $record.Od_x0020_kiedy                                   # date
    $NRIDNEW                    = $record.NRIDNEW                                          # string
    $Contract_x0020_type        = $record.Contract_x0020_type                              # string
    $Job_x0020_title            = $record.Job_x0020_Title                                  # string
    $Stanowisko                 = $record.Stanowisko                                       # string
    $Poziom                     = $record.Poziom                                           # string
    $Manager_x0020_Name         = $record.Manager_x0020_Name                               # FieldUserValue
    $Dzia_x0142_                = $record.Dzia_x0142_                                      # string
    $Kraj                       = $record.Kraj                                             # string
    $Lokalizacja                = $record.Lokalizacja                                      # string
    $Dodatkowe_x0020_informacje = $record.Dodatkowe_x0020_Informacje                       # string
    $Ad_login                   = $record.AD_login                                         # string
    $Email                      = $record.Email                                            # string
    $Po                         = $record.PO                                               # string
    $Company                    = $record.Company                                          # string
    $Licences                   = $record.Licences                                         # string
    $ID                         = $record.ID                                               # string
    $Manager                    = $Manager_x0020_Name.LookupValue                          # string
    $JobTimeSize                = $record.JobTimeSize                                      # string
    $Billable                   = $record.Billable                                         # string
    $TechLeader                 = $record.TechLeader                                       # FieldUserValue
    $Technology                 = $record.Technology                                       # string

    if (!$NRIDNEW) {
        [int]$last_value = Get-Content -Path "path to memory file" -Tail 1

        $last_value = $last_value + 1
        $NumberIDNEW = $last_value.ToString()
        Set-PnPListItem -List $List -Identity $ID -Values @{"NRIDNEW" = "$NumberIDNEW" }
        set-Content -Value $NumberIDNEW -Path "path\to\memory_file"
    }

    if (!(check-white-signs -String $Title))    { $Error.Add("White signs detected in variable = Title") }
    if (!(check-white-signs -String $Nazwisko)) { $Error.Add("White signs detected in variable = Surname") }
    if (!(check-white-signs -String $Ad_login)) { $Error.Add("White signs detected in variable = Adlogin") }
    if (!(check-white-signs -String $Email))    { $Error.Add("White signs detected in variable = Email") }

    $User_DN = (get-aduser -Identity $Ad_login).DistinguishedName
    Move-ADObject -Identity $User_DN -TargetPath (match-country-DC -country $Kraj -city $Lokalizacja )
    $passx = password-generator
    $pass = ConvertTo-SecureString -AsPlainText ($passx) -Force
    Set-ADAccountPassword -Identity $Ad_login -NewPassword $pass
    Set-ADUser -Enabled $true -Identity $Ad_login -SmartcardLogonRequired $false
    Set-ADUser -Identity $Ad_login -Replace @{extensionAttribute2 = "$Licences" }
    #endregion Main Code

    #region Email notification
    $BodySuccess = "Dear Support Team,

Please be informed that account was enabled with success ( $User_created ). 

Basic information

UPN Name:     $Email
SAM Name:     $Ad_login
ID number:    $NRIDNEW
New Pass:     $passx



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
    Set-PnPListItem -List $List -Identity $ID -Values @{"Dodatkowe_x0020_Informacje" = "$Dodatkowe_x0020_informacje; User enabled at $date_ad" ; "Zmiana" = "Update"; "Disabled" = "No" }
}
#endregion Email notification