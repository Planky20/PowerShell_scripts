$csvPath = "path"

Import-CSV $csvPath | % { 
 Try {
  $username = $_.username  # Assuming the CSV has a 'username' column.
  $city = $_.city  # Assuming the CSV has a 'city' column.

  # Get the AD user with the necessary property 'city'
  $adUser = Get-ADUser -Filter "SamAccountName -eq '$username'" -Properties l

  if ($adUser -and ($adUser.l -ne $city)) {
            
   Write-Host "City changed for user $username from $($adUser.l) to $city"
   Set-ADUser -Identity $username -Replace @{l = $city }
  }
  elseif ($adUser -and ($adUser.l -eq $city)) {
   Write-Host "No changes for user $username (actual city is: $($adUser.l))"
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