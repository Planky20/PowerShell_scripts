# Connect to Microsoft Graph with user authentication
Connect-MgGraph -Scopes "User.Read.All"

# Define the usage location you are filtering by
$usageLocation = "PL"  # Replace with your specific usage location

# Initialize an array to hold user data
$userList = @()

# Get all users with the specific usage location
$users = Get-MgUser -Filter "usageLocation eq '$usageLocation' and accountEnabled eq true" -Property "id,displayName,userPrincipalName,usageLocation,accountEnabled" -All

# Process each user and add to the array
foreach ($user in $users) {
    $userList += [PSCustomObject]@{
        Id                = $user.Id
        DisplayName       = $user.DisplayName
        UserPrincipalName = $user.UserPrincipalName
        UsageLocation     = $user.UsageLocation
        AccountEnabled    = $user.AccountEnabled
    }
}

# Export the user list to a CSV file
$userList | Export-Csv -Path "path to CSV" -NoTypeInformation -Encoding UTF8