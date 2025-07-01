$csvPath = "path to csv"
 
Import-CSV $csvPath | ForEach-Object {  
	Try {
		$username = $_.username              # Assuming the CSV has a 'username' column.
		$extensionAttribute2 = $_.extensionAttribute2   # Assuming the CSV has a 'extensionAttribute2' column.
 
  $adUser = Get-ADUser -Filter 'SamAccountName -eq $username' -Properties extensionAttribute2, manager

  if ($adUser -and ($adUser.extensionAttribute2 -ne $extensionAttribute2)) {
            
   Write-Host "ExtensionAttribute2 changed for user $username from $($adUser.extensionAttribute2) to $extensionAttribute2"
   Set-ADUser -Identity $username -Replace @{extensionAttribute2 = $extensionAttribute2 }
  
  }
  elseif ($adUser -and ($adUser.extensionAttribute2 -eq $extensionAttribute2)) {
   Write-Host "No changes for user $username (actual ExtensionAttribute2 is: $($adUser.extensionAttribute2))"
  }
  else {
   Write-Host "User $username has not been found in Active Directory"
  }
 }

 Catch {
  $ErrorOccured = $true
  "Errors"
  Write-Error -Message "Error: $($_.Exception.Message) - Line Number: $($_.InvocationInfo.ScriptLineNumber)"
 }
 $ErrorOccured = $false

}