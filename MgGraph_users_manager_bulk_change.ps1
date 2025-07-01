# Import the Microsoft Graph module
Import-Module Microsoft.Graph

# Define required scopes
$scopes = @("User.ReadWrite.All")

# Connect to Microsoft Graph
Connect-MgGraph -Scopes $scopes

# Define the CSV file path (Update this accordingly)
$csvPath = "path"

# Import CSV data
$users = Import-Csv -Path $csvPath

# Iterate through each user in the CSV
foreach ($user in $users) {
    $userUPN = $user.UserPrincipalName
    $managerUPN = $user.ManagerUserPrincipalName

    try {
        # Get the Manager's Object ID
        $manager = Get-MgUser -UserId $managerUPN
        if ($manager -ne $null) {
            $managerId = $manager.Id

            # Create the manager reference object properly
            $managerRef = @{
                "@odata.id" = "https://graph.microsoft.com/v1.0/users/$managerId"
            }

            # Update the user's manager
            Set-MgUserManagerByRef -UserId $userUPN -BodyParameter $managerRef

            Write-Host "Updated manager for $userUPN -> $managerUPN"
        }
        else {
            Write-Host "Manager $managerUPN not found for user $userUPN" -ForegroundColor Yellow
        }
    }
    catch {
        Write-Host "Failed to update manager for $userUPN $_" -ForegroundColor Red
    }
}

# Disconnect from Microsoft Graph
Disconnect-MgGraph