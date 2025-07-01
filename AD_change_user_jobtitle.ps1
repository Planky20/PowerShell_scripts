$csvPath = "path to CSV"

Import-CSV $csvPath | ForEach-Object { 
 Try {
  $username = $_.username  # Assuming the CSV has a 'username' column.
  $jobtitle = $_.jobtitle  # Assuming the CSV has a 'jobtitle' column.

  $adUser = Get-ADUser -Filter "SamAccountName -eq '$username'" -Properties title

  if ($adUser -and ($adUser.title -ne $jobtitle)) {
            
   Write-Host "Job title changed for user $username from $($adUser.title) to $jobtitle"
   Set-ADUser -Identity $username -Title "$jobtitle"

  }
  elseif ($adUser -and ($adUser.title -eq $jobtitle)) {
   Write-Host "No changes for user $username (actual job title is: $($adUser.title))"
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