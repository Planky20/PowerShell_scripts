$exportCsvPath = "path"
$ouPath = "OU=Poznan,OU=Poland,OU=Users,OU=CloudSync,DC=your,DC=domain,DC=FQDN"  # Specify your target OU
$streetAddress = "new street address"  # Specify the new street address

# Step 1: Export users from the specified OU to a CSV file
Try {
    Write-Host "Exporting users from OU: $ouPath..."
    Get-ADUser -SearchBase $ouPath -Filter * -Properties SamAccountName | 
    Select-Object SamAccountName | 
    Export-Csv -Path $exportCsvPath -NoTypeInformation
    Write-Host "Export completed: $exportCsvPath"
}
Catch {
    Write-Error -Message "Error exporting users: $($_.Exception.Message)"
    Exit
}

# Step 2: Update street address for all users in the exported CSV
Import-CSV $exportCsvPath | % { 
    Try {
        $username = $_.SamAccountName  # Using the exported usernames

        # Get the AD user with the current 'streetAddress'
        $adUser = Get-ADUser -Filter "SamAccountName -eq '$username'" -Properties streetAddress

        if ($adUser -and ($adUser.streetAddress -ne $streetAddress)) {
            Write-Host "Street address changed for user $username from $($adUser.streetAddress) to $streetAddress"
            Set-ADUser -Identity $username -Replace @{streetAddress = $streetAddress }
        }
        elseif ($adUser -and ($adUser.streetAddress -eq $streetAddress)) {
            Write-Host "No changes for user $username (actual street address is already: $($adUser.streetAddress))"
        }
        else {
            Write-Host "User $username has not been found in Active Directory"
        }
    }
    Catch {
        Write-Error -Message "Error updating user $username $($_.Exception.Message) - Line Number: $($_.InvocationInfo.ScriptLineNumber)"
    }
}
