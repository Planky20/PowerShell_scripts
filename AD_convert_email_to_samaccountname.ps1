# Import the Active Directory module
Import-Module ActiveDirectory

# Read the CSV file
$users = Import-Csv -Path "C:\temp\PLemailconvert.csv" -Encoding utf8 -Delimiter ";"

# Create an empty array to hold the results
$results = @()

foreach ($user in $users) {
    # Get the user from Active Directory
    $adUser = Get-ADUser -Filter "EmailAddress -eq '$($user.Email)'" -Properties EmailAddress

    # If the user was found, add their SamAccountName to the results
    if ($adUser) {
        $results += New-Object PSObject -Property @{
            "Email"          = $user.Email
            "SamAccountName" = $adUser.SamAccountName
        }
    }
}

# Export the results to a new CSV file
$results | Export-Csv -Path "C:\temp\output.csv" -NoTypeInformation