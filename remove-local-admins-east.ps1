# Define the names of the admin accounts to keep
$accountsToKeep = @("AdminEAST", "Administrator")
$localMachineName = $env:COMPUTERNAME

# Import the .NET namespace for working with user accounts
Add-Type -AssemblyName System.DirectoryServices.AccountManagement

# Get the context for the local machine
$principalContext = New-Object System.DirectoryServices.AccountManagement.PrincipalContext('Machine', $localMachineName)

# Get all members of the Administrators group
$adminGroup = [System.DirectoryServices.AccountManagement.GroupPrincipal]::FindByIdentity($principalContext, 'Administrators')
$adminMembers = $adminGroup.GetMembers()

# Loop through each member of the Administrators group
foreach ($member in $adminMembers) {
    $username = $member.SamAccountName

    # Check if the account is a local account
    $isLocalAccount = $member.ContextType -eq 'Machine'

    # Output for debugging
    Write-Output "Processing user/group: $username"

    # Check if the account is a non-local account (e.g., Azure AD)
    $isNonLocalAccount = -not $isLocalAccount

    # Skip processing if the account is non-local
    if ($isNonLocalAccount) {
        Write-Output "Skipping non-local account: $username"
        Continue
    }
    $userAzure = $member.DisplayName -match 'AzureAD'
    if ($userAzure) {
        Write-Output "Skipping azure account: $username"
        Continue
    }
    $groupAzure = $member.Sid -match 'S-1-12-'
    if ($groupAzure) {
        Write-Output "Skipping azure group: $username"
        Continue
    }

    # Check if the account should be kept
    if ($accountsToKeep -notcontains $username) {
        Write-Output "Removing local user from Administrators group: $username"
        try {
            # Remove the user from the Administrators group
            $adminGroup.Members.Remove($member)
            $adminGroup.Save()

            Write-Output "Successfully removed $username from Administrators group"

            # Delete the user account completely
            try {
                $user = [System.DirectoryServices.AccountManagement.UserPrincipal]::FindByIdentity($principalContext, $username)
                if ($user -ne $null) {
                    $user.Delete()
                    Write-Output "Successfully deleted local user: $username"
                }
                else {
                    Write-Output "Local user $username not found for deletion."
                }
            }
            catch {
                Write-Output "Failed to delete local user: $username - $_"
            }
        }
        catch {
            Write-Output "Failed to remove $username from Administrators group - $_"
        }
    }
    else {
        Write-Output "Keeping user/group: $username"
    }
}