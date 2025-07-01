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

#region Function to remove Polish characters from name and surname
function Remove-StringLatinCharacters {
	PARAM ([string]$String)
	[Text.Encoding]::ASCII.GetString([Text.Encoding]::GetEncoding("Cyrillic").GetBytes($String))
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
		"yes" { $result = $true }
		"no" { $result = $false }
		"TRUE" { $result = $true }
		"FALSE" { $result = $false }
	}
	return $result
}
#endregion

#region Function to change password for blocked user
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
	if (!$Object.Email) {
		$test = $Object.LookupValue
		$results = (Get-ADUser -Filter { Name -eq $test }).SamAccountName
	}
	else {
		$test = $Object.Email
		$results = (Get-ADUser -Filter { UserPrincipalName -eq $test }).SamAccountName
	}
	return $results
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
$from           = "sender email address"
$To             = "recipient email address"
$SubjectFault   = "Operation of disabling account failed"
$SubjectSuccess = "Account was disabled successfully"
$date_ad        = Get-Date -Format "dd/MM/yyyy HH:mm"
$time_date      = get-date -Hour 0 -Minute 0 -Second 0 -Millisecond 0
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


$Lists = @()
foreach ($records_spo in $records_spos) {
	if (check-record -value ($records_spo.Ready_x0020_for_x0020_deploy) -key "True") {
		if (check-record -value $records_spo.Zmiana -key "Finish work" ) {
			if (check-record -value $records_spo.Disabled -key "No" ) {
				$Lists += $records_spo
			}
		}
	}
}

$Lists.Count

foreach ($record in $lists) {
	$User_created                 = $record.Title + " " + $record.Nazwisko                      		# string
	$Title                        = Remove-StringLatinCharacters -String $record.Title          		# string
	$Nazwisko                     = Remove-StringLatinCharacters -String $record.Nazwisko       		# string
	$Do_x0020_kiedy               = $record.Do_x0020_kiedy                                    				# date
	$Dodatkowe_x0020_informacje   = $record.Dodatkowe_x0020_Informacje            																# string
	$Temporary                    = Change-tobool -string $record.Temporary                     		# string to bool
	$Ready_x0020_for_x0020_deploy = Change-tobool -string $record.Ready_x0020_for_x0020_deploy  		# string to bool
	$Deployed                     = Change-tobool -string $record.Deployed                      		# string to bool
	$Disabled                     = Change-tobool -string $record.Disabled                      		# string to bool
	$Ad_login                     = $record.AD_login                                            		# string
	$Email                        = $record.Email                                               		# string
	$password                     = password-generator                                          		# string
	$Licences                     = $record.Licences                                            		# string
	$ID                           = $record.ID                                                    # string
	$Do_x0020_kiedy               = $Do_x0020_kiedy.AddHours(-$Do_x0020_kiedy.Hour)               # date
	$Do_x0020_kiedy               = $Do_x0020_kiedy.AddMilliseconds(-$Do_x0020_kiedy.Millisecond) # date

	if (  $Do_x0020_kiedy -ge $time_date ) {}
	else {
		$Error.Clear()
		$pass = ConvertTo-SecureString -AsPlainText $password -Force
		try {
			Get-ADUser -Identity $Ad_login -Properties MemberOf | ForEach-Object {
				$_.MemberOf | Remove-ADGroupMember -Members $_.DistinguishedName -Confirm:$false
			}
		}
		catch { $Error.Add("Unable to remove User membership group") }

		Set-ADAccountPassword -Identity $Ad_login -NewPassword $pass
		Set-ADUser -Identity $Ad_login -Clear manager
		Set-ADUser -Identity $Ad_login -Clear Department
		Set-ADUser -Enabled $false -Identity $Ad_login -SmartcardLogonRequired $true
		Set-ADUser -Identity $Ad_login -Replace @{extensionAttribute2 = "EMPTY" }
		$User_DN = (get-aduser -Identity $Ad_login).DistinguishedName
		Start-Sleep -Seconds 4
		Move-ADObject -Identity $User_DN -TargetPath "OU=Disabled,OU=Users,OU=CloudSync,DC=your,DC=domain,DC=fqdn"

		Set-PnPListItem -List $List -Identity $ID -Values @{"Dodatkowe_x0020_Informacje" = "$Dodatkowe_x0020_informacje ;User disabled at $date_ad"; "Disabled" = "Yes"; "Licences" = "EMPTY" }
		#endregion Main Code

		#region Email notification
		if ($Error) {
			$BodyFault = "
Dear Support Team,

Please be informed that disabling of account was finished with error ( $User_created ). 
Please investigate what happened. You can find all details below. 

$Error


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
			$BodySuccess = "Dear Support Team,

Please be informed that account was disabled successfully ( $User_created ). 


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
		}
	}
}
#endregion Email notification