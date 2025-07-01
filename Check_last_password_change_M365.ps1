#Connect to Microsoft Graph
Connect-MgGraph -Scopes "User.Read.All"

Import-Module MGgraph

#Get the User
$User = Get-MgUser -UserId "UPN" -Property UserPrincipalName, PasswordPolicies, lastPasswordChangeDateTime
 
#Get the user's last password change date and time
$User | Select UserPrincipalName, PasswordPolicies, lastPasswordChangeDateTime
