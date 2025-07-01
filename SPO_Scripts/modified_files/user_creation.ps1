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
      $result = 'OU=' + $city + ',OU=' + $country + ',OU=Users,OU=ADSync,DC=your,DC=tenant,DC=FQDN'
   }
   else {
      $result = 'OU=' + $city + ',OU=' + $country + ',OU=Users,OU=CloudSync,DC=your,DC=tenant,DC=FQDN'
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
   $probe = $false
   if ($SAM.Length -gt 4) {
      $number = [int]$SAM.Remove(0, 4)
      $base = $SAM.Remove(4, 1)
   }
   else {
      $number = 2
      $base = $SAM
      if (!(Get-ADUser -Filter { SamAccountName -eq $SAM }).SamAccountName) {
         $tmp = $SAM
         $probe = $true
      }
   }
   do { 
      if ((Get-ADUser -Filter { SamAccountName -eq $SAM }).SamAccountName) {
         $tmp = $base + $number
      }
      if ((Get-ADUser -Filter { SamAccountName -eq $tmp }).SamAccountName) {
         $number++
         $probe = $false
      }
      else {
         $probe = $true
      }
   } until ($probe -eq $true )

   return $tmp
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
   $upn_1 = $UPN
   $probe = $false
   $base = $UPN.Remove(($UPN.indexof("@")), ($UPN.Length - ($UPN.indexof("@"))))
   $number = 2
   if (!(Get-ADUser -Filter { UserPrincipalName -eq $UPN })) { $probe = $true }
   else {
      do {   
         $upn_1 = $base + $number + "@" + $domain
         if (Get-ADUser -Filter { UserPrincipalName -eq $upn_1 }) {
            $probe = $false
            $number++
         }
         else {
            $probe = $true
         }
      } until ($probe -eq $true )
   }
   return $upn_1 
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
      "yes"    { $result = $true  }
      "no"     { $result = $false }
      "TRUE"   { $result = $true  }
      "FALSE"  { $result = $false }
   }
   return $result
}
#endregion

#region Function to generate password for new user
function password-generator() {
   Param(
      [int]$Strong = 10
   )
   $Password = "!@#$%^&*0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz".tochararray() 
   $Password1 = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ_abcdefghijkmnopqrstuvwxyz".tochararray() 
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

#region Function to add national code to CountryCode attribute
function add-national-code() {
   Param(
      [string]$country
   )

   switch ($country) {
      "POLAND"          { $countrycode = "616" }
      "UNITED KINGDOM"  { $countrycode = "826" }
      "SWITZERLAND"     { $countrycode = "756" }
   }

   return $countrycode
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
   $prefix = ($UserPrincipalName.Remove($UserPrincipalName.IndexOf("@"), ($UserPrincipalName.Length - $UserPrincipalName.IndexOf("@"))))
   $results = $prefix + $sufix

   return $results
}
#endregion

#region Function to define sec
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

#region Function to populate company attribute
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

#region Function to check if DisplayName is unique
function check-name () {
   Param(
      [string]$FirstName,
      [string]$LastName
   )
   $probe = $false
   $base = "$LastName $FirstName"
   $number = 2
   if (!(get-aduser -Filter { Name -eq $base })) {
      $result = "$LastName $FirstName"
   }
   else {
      do {   
         $line_1 = $base + " " + $number
         if (get-aduser -Filter { Name -eq $line_1 }) {
            $probe = $false
            $number++
         }
         else {
            $probe = $true
         }
      } until ($probe -eq $true )
      $result = $line_1 
   }

   return $result
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

#region Function to get city name
function get-city () {
   param(
      [string]$city
   )
   $result = $city.Replace("2", "")

   return $result
}
#endregion
#endregion Functions definition

#region Variables definition
$path_report    = "\\path\mgmt$" # Path to shared folder with reports on FS
$SiteUrl        = "SP site URL"
$listName       = "SP list name"
$tenantID       = "tenant ID"
$ClientId1      = "app reg 1 ID"
$ClientId2      = "app reg 2 ID"
$KeyVaultName   = "KV name"
$domain_main    = "domain.com"
$company        = "Company name"
$from           = "sender email address"
$To             = "recipient email address"
$SubjectFault   = "Operation of adding new user failed"
$SubjectSuccess = "Adding new user was completed with success"
$SubjectOrder   = "New device request appeared"
$date_ad        = Get-Date -Format "dd/MM/yyyy HH:mm"
#endregion Variables definition

#region Start of the transcript
$transcriptFolder = Join-Path -Path $path_report -ChildPath "user_creation_transcripts"
if (!(Test-Path -Path $transcriptFolder)) {
   New-Item -Path $transcriptFolder -ItemType Directory
}
$transcriptPath = Join-Path -Path $transcriptFolder -ChildPath "user_creation_V3_transcript_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
Start-Transcript -Path $transcriptPath -Force
#endregion Start of the transcript

#region Environment connection
$managedIdentityMetadata = Invoke-RestMethod -Headers @{Metadata = "true" } -Method GET -Uri 'http://169.254.169.254/metadata/identity/oauth2/token?api-version=2018-02-01&resource=https%3A%2F%2Fvault.azure.net'
$KeyVaultToken = $managedIdentityMetadata.access_token
$ClientCert = (Invoke-RestMethod -Uri https://$KeyVaultName.vault.azure.net/secrets/<CERTNAME>?api-version=2016-10-01 -Method GET -Headers @{Authorization = "Bearer $KeyVaultToken" }).value
$ClientSecret = (Invoke-RestMethod -Uri https://$KeyVaultName.vault.azure.net/secrets/<SECRETNAME>?api-version=2016-10-01 -Method GET -Headers @{Authorization = "Bearer $KeyVaultToken" }).value
Import-Module Microsoft.Graph.Authentication
Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId1 -ClientSecret $ClientSecret
Connect-MgGraph -ClientId $ClientId2 -CertificateThumbprint $ClientCert -TenantId $tenantID
#endregion Environment connection

#region Main Code
$List = Get-PnPList -Identity $listName
$records_spos = (Get-PnPListItem -List $List | Select-Object id, @{label = "Filename"; expression = { $_.FieldValues } }).filename


$Lists = @()
foreach ($records_spo in $records_spos) {
   if (check-record -value ($records_spo.Ready_x0020_for_x0020_deploy) -key "True") {
      if (check-record -value $records_spo.Zmiana -key "Start work" ) {
         if (check-record -value $records_spo.Disabled -key "No" ) {
            if (check-record -value $records_spo.Deployed -key "No" ) {
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
   $User_created                 = $record.Title + " " + $record.Nazwisko                      # string
   $Title                        = Remove-StringLatinCharacters -String $record.Title          # string
   $Nazwisko                     = Remove-StringLatinCharacters -String $record.Nazwisko       # string
   $Sex                          = $record.Sex                                                 # string
   $NR_x0020_ID                  = $record.NR_x0020_ID                                         # string
   $Contract_x0020_type          = $record.Contract_x0020_type                                 # string
   $Job_x0020_title              = $record.Job_x0020_Title                                     # string
   $Stanowisko                   = $record.Stanowisko                                          # string
   $Poziom                       = $record.Poziom                                              # string
   $Manager_x0020_Name           = $record.Manager_x0020_Name                                  # FieldUserValue
   $Zmiana                       = $record.Zmiana                                              # string
   $Dzia_x0142_                  = $record.Dzia_x0142_                                         # string
   $Od_x0020_kiedy               = $record.Od_x0020_kiedy                                      # date
   $Do_x0020_kiedy               = $record.Do_x0020_kiedy                                      # date
   $Kraj                         = $record.Kraj                                                # string
   $Lokalizacja                  = $record.Lokalizacja                                         # string
   $Zapotrzebowanie              = $record.Zapotrzebowanie                                     # string[]
   $Dodatkowe_x0020_informacje   = $record.Dodatkowe_x0020_Informacje                          # string
   $Ready_x0020_for_x0020_deploy = Change-tobool -string $record.Ready_x0020_for_x0020_deploy  # string to bool
   $Deployed                     = Change-tobool -string $record.Deployed                      # string to bool
   $Disabled                     = Change-tobool -string $record.Disabled                      # string to bool
   $Ad_login                     = $record.AD_login                                            # string
   $Email                        = $record.Email                                               # string
   $Po                           = $record.PO                                                  # string
   $Company                      = $record.Company                                             # string
   $password                     = password-generator                                          # string
   $Licences                     = $record.Licences                                            # string
   $ID                           = $record.ID                                                  # int
   $Manager                      = $Manager_x0020_Name.LookupValue                             # string
   $JobTimeSize                  = $record.JobTimeSize                                         # string
   $Billable                     = $record.Billable                                            # string
   $CompanyCode0                 = $record.CompanyCode0                                        # string
   $TechLeader                   = $record.TechLeader                                          # FieldUserValue
   $Technology                   = $record.Technology                                          # string
   $Hardwarerequirements         = $record.Hardwarerequirements                                # string
   $ProfitCenter                 = $record.PO                                                  # string


   #region NRIDNEW assignment if null
   if ($record.NRIDNEW -eq $null) {
      [int]$last_value = Get-Content -Path "path to memory file" -Tail 1

      $last_value = $last_value + 1
      $NumberIDNEW = $last_value.ToString()
      Set-PnPListItem -List $List -Identity $ID -Values @{"NRIDNEW" = "$NumberIDNEW" }
      set-Content -Value $NumberIDNEW -Path "path to memory file"

      $NRIDNEW = $NumberIDNEW
   }
   else {
      $NRIDNEW = $record.NRIDNEW 
   }
   #endregion NRINDEW assignment if null

   #region Date condition check
   $start_date = $Od_x0020_kiedy.date
   $two_weeks_before_start_date = $start_date.AddDays(-14)
   $today = Get-Date
   $today_date = $today.Date
   #endregion Date condition check

   if ($today_date -ge $two_weeks_before_start_date) {
      # Proceed with account creation
      $errors = @()
      if (check-white-signs -String $Title) { $errors.Add("White signs detected in variable = Title") }
      if (check-white-signs -String $Nazwisko) { $errors.Add("White signs detected in variable = Surname") }

      #region Try and catch block for sending email
      try {
         # SAM and UPN definition
         $SAM = (Build-SAM -Givenname $Nazwisko -Surname $Title)
         $SamAccountName = check-SAM -SAM $SAM
         $UPN = (Build-UPN -Givenname $Nazwisko -Surname $Title -domain $domain_main )
         $UserPrincipalName = check-UPN -UPN $UPN -domain $domain_main

         $DisplaY_user_name = check-name -FirstName $Nazwisko -LastName $Title
         $pass = ConvertTo-SecureString -AsPlainText $password -Force
         New-ADUser -AccountPassword $pass -GivenName $Nazwisko -Surname $Title -DisplayName $DisplaY_user_name -Name $DisplaY_user_name -SamAccountName $SamAccountName -UserPrincipalName $UserPrincipalName -EmailAddress $UserPrincipalName -City (get-city -city $Lokalizacja) -Company (get-company -ContractType $Contract_x0020_type -Company $company ) -Country (match-country -country $Kraj) -Department $Dzia_x0142_ -Enabled $true -Manager (get-manager-name -Object $Manager_x0020_Name) -EmployeeID $NRIDNEW -Title $Job_x0020_title -StreetAddress (match-address -city $Lokalizacja ) -Office (get-city -city $Lokalizacja) -PostalCode (match-address-code -city $Lokalizacja ) -path (match-country-DC -country $Kraj -city $Lokalizacja )
         Start-Sleep -Seconds 10
         Enable-ADAccount -Identity $SamAccountName
         Set-ADUser -Identity $SamAccountName -Add @{extensionAttribute14 = $Billable; extensionAttribute1 = $JobTimeSize; extensionAttribute3 = $Stanowisko; extensionAttribute4 = $Poziom; extensionAttribute2 = $Licences; extensionAttribute6 = $Lokalizacja; extensionAttribute7 = $Contract_x0020_type; extensionAttribute5 = (get-sex -Sex $Sex) }

         if ($ProfitCenter) {
            Set-ADUser -Identity $SamAccountName -Add @{extensionAttribute12 = $ProfitCenter }
         }

         if ($CompanyCode0) {
            Set-ADUser -Identity $SamAccountName -Add @{comment = $CompanyCode0 }
         }

         Set-ADUser -Identity $SamAccountName -Add @{extensionAttribute11 = $NRIDNEW }
         Set-ADUser -Identity $SamAccountName -Add @{carlicense = (date-string -Data $Od_x0020_kiedy) }
         Set-ADUser -Identity $SamAccountName -Add @{co = $Kraj }
         Set-ADUser -Identity $SamAccountName -Replace @{countrycode = (add-national-code -country $Kraj) }
         Set-ADUser -Identity $SamAccountName -Add @{msExchUsageLocation = (match-country -country $Kraj) }

         if ($TechLeader) {
            Set-ADUser -Identity $SamAccountName -Add @{pager = (get-techleader-name -Object $TechLeader) }
         }

         if ($Technology) {
            Set-ADUser -Identity $SamAccountName -Add @{personalPager = $Technology }
         }

         $National = add-national-sufix -country $Kraj -UserPrincipalName $UserPrincipalName

         if ($National) {

            Set-ADUser -Identity $SamAccountName -add @{ProxyAddresses = "SMTP:$UserPrincipalName" }
            Set-ADUser -Identity $SamAccountName -add @{ProxyAddresses = "smtp:$National" }

         }

         Set-ADUser -Identity $SamAccountName -PasswordNotRequired $false -Description $Job_x0020_title
         Add-ADGroupMember -Identity (add-national-local-group -country $Kraj) -Members $SamAccountName
      }

      catch {
         $errors.Add($_.Exception.Message)
      }

      if ($errors.Count -gt 0) {
         $BodyFault = "

Dear Support Team,

Please be informed that creation of new account was finished with error ( $User_created ). 
Please investigate what happened. You can find all details below. 

$($errors -join '')

Best Regards

Orchestrator
"

         $params = @{
            Message         = @{
               Subject      = $SubjectFault
               Body         = @{
                  ContentType = "Text"
                  Content     = $BodyFault
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
         Set-PnPListItem -List $List -Identity $ID -Values @{"Zmiana" = "Error" }
      }
      else {
         # Send success email only if no errors
         Set-PnPListItem -List $List -Identity $ID -Values @{"Email" = "$UserPrincipalName"; "AD_login" = "$SamAccountName"; "Zmiana" = "Auto-Update"; "Dodatkowe_x0020_Informacje" = "User created at $date_ad"; "Deployed" = "Yes" }
         $BodySuccess = "Dear Support Team,

Please be informed that creation of new account was completed with success ( $User_created ).

Basic information

UPN Name:           $UserPrincipalName
SAM Name:           $SamAccountName
Password:              $password


Other information

Department:         $Dzia_x0142_
Manager:              $Manager
Start working:      $Od_x0020_kiedy
Office:                $Lokalizacja





Best Regards,
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
      }
   }
   else {
      continue  # Skip to next record
   }
}
#endregion Try and catch block for sending email

#endregion Main Code

#region IT Support email notification about new user hardware requirements
$jira = "IT Support email address"

# Retrieve items where NotificationSent is not set or is set to 'No'
$pendingNotifications = Get-PnPListItem -List $listName | Where-Object {
   ($_.FieldValues.NotificationSent -eq $null) -or ($_.FieldValues.NotificationSent -eq $false)
}

foreach ($record in $pendingNotifications) {
   $User_created         = $record.Title + " " + $record.Nazwisko
   $Poziom               = $record.Poziom
   $Lokalizacja          = $record.Lokalizacja
   $Od_x0020_kiedy       = $record.Od_x0020_kiedy
   $Zapotrzebowanie      = $record.Zapotrzebowanie
   $Hardwarerequirements = $record.Hardwarerequirements
   $ID                   = $record.ID

   if ($Zapotrzebowanie) {
      $BodyOrder = "
Dear Support Team,

Please be informed that HR department has requested a device for new employee ( $User_created ). 
You can find details below. 

User competency level is : $Poziom
Location                 : $Lokalizacja
Start from               : $Od_x0020_kiedy

$Zapotrzebowanie
$Hardwarerequirements


Best Regards

Orchestrator
"
      $params = @{
         Message         = @{
            Subject      = $SubjectOrder
            Body         = @{
               ContentType = "Text"
               Content     = $BodyOrder
            }
            ToRecipients = @(
               @{
                  EmailAddress = @{
                     Address = $jira
                  }
               }
            )
         }
         SaveToSentItems = "false"
      }
      try {
         Send-MgUserMail -UserId $from -BodyParameter $params
         # Update the NotificationSent field to 'Yes' after successful email sending
         Set-PnPListItem -List $listName -Identity $ID -Values @{"NotificationSent" = $true }
      }
      catch {
         Write-Host "Failed to send notification email for $User_created. Error: $_"
         # Optionally handle the error, e.g., log it or set a different flag
      }
   }
}
#endregion IT Support email notification about new user hardware requirements
Stop-Transcript