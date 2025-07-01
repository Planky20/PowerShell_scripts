$csvPath = "path"

Import-CSV $csvPath | % {  
    Try {
        $username = $_.username
        $extensionAttribute12 = $_.extensionAttribute12
        $manager = $_.manager

        #$adUser = Get-ADUser -Filter "SamAccountName -eq '$username'" -Properties extensionAttribute12, manager
        $adUser = Get-ADUser -Filter 'SamAccountName -eq $username' -Properties extensionAttribute12, manager


        if ($adUser -and (($adUser.extensionAttribute12 -ne $extensionAttribute12) -or ($adUser.manager -ne $manager))) {
            
            Write-Host "Zmieniono ExtAttribute12 dla uzytkownika $username z $($adUser.extensionAttribute12) na $extensionAttribute12"
            Set-ADUser -Identity $username -Add @{extensionAttribute12 = $extensionAttribute12 }

            Write-Host "Zmieniono managera dla uzytkownika $username z $($adUser.manager) na $manager"
            Set-ADUser -Identity $username -Manager $manager

        }
        else {
            Write-Host "Uzytkownik $username nie zostal znaleziony w Active Directory"
        }
    }
    Catch {
        $ErrorOccured = $true
        "Errors"
        Write-Error -Message "Error: $($_.Exception.Message) - Line Number: $($_.InvocationInfo.ScriptLineNumber)"
    }

    $ErrorOccured = $false
}