#region Variables definition
$siteURL = "SP site URL"
$listName = "SP list name"
$tenantID = "tenant ID"
$ClientId1 = "app reg 1 ID"
$ClientId2 = "app reg 2 ID"
$KeyVaultName = "KV name"
$path_report = "\\path\mgmt$" # Path to shared folder with reports on FS
$child_path = "Child folder name"
#endregeon Variables definition

#region Start of the transcript
$transcriptFolder = Join-Path -Path $path_report -ChildPath $child_path
if (!(Test-Path -Path $transcriptFolder)) {
    New-Item -Path $transcriptFolder -ItemType Directory
}
$transcriptPath = Join-Path -Path $transcriptFolder -ChildPath "LogFileName_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
Start-Transcript -Path $transcriptPath -Force
#endregion Start of the transcript

#region Environment connection
$Response = Invoke-RestMethod -Uri 'http://169.254.169.254/metadata/identity/oauth2/token?api-version=2018-02-01&resource=https%3A%2F%2Fvault.azure.net' -Method GET -Headers @{Metadata = "true" }
$KeyVaultToken = $Response.access_token

$ClientSecretSPO = (Invoke-RestMethod -Uri https://$KeyVaultName.vault.azure.net/secrets/<SPOKVSecretName>?api-version=2016-10-01 -Method GET -Headers @{Authorization = "Bearer $KeyVaultToken" }).value
$ClientCertEXO = (Invoke-RestMethod -Uri https://$KeyVaultName.vault.azure.net/secrets<EXOKVCertName>?api-version=2016-10-01 -Method GET -Headers @{Authorization = "Bearer $KeyVaultToken" }).value
$ClientCertMgGraph = (Invoke-RestMethod -Uri https://$KeyVaultName.vault.azure.net/secrets/<MgGraphKVCertName>?api-version=2016-10-01 -Method GET -Headers @{Authorization = "Bearer $KeyVaultToken" }).value

Import-Module Microsoft.Graph.Authentication
Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId1 -ClientSecret $ClientSecretSPO

Add-Type -Path "C:\Program Files\PowerShell\Modules\ExchangeOnlineManagement\3.8.0\netCore\Microsoft.Identity.Client.dll"
Add-Type -Path "C:\Program Files\PowerShell\Modules\ExchangeOnlineManagement\3.8.0\netCore\Microsoft.IdentityModel.JsonWebTokens.dll"
Add-Type -Path "C:\Program Files\PowerShell\Modules\ExchangeOnlineManagement\3.8.0\netCore\Microsoft.IdentityModel.Tokens.dll"
Connect-ExchangeOnline -AppId $ClientId2 -Organization "tenant full name" -CertificateThumbprint $ClientCertEXO

Connect-MgGraph -ClientId $ClientId2 -CertificateThumbprint $ClientCertMgGraph -TenantId $tenantID
#endregion Environment connection

#region Collect EXO Data
$allGroups = @()
$nameFilter = "*<FILTER>*" # Name filter for groups

# Distribution Lists
$DistributionLists = Get-DistributionGroup -RecipientTypeDetails MailUniversalDistributionGroup |
Where-Object { $_.DisplayName -like $nameFilter } |
Select-Object DisplayName, PrimarySmtpAddress, WhenCreated, ManagedBy, @{Name = "GroupType"; Expression = { "Distribution List" } }
$allGroups += $DistributionLists

# Dynamic Distribution Lists
$DynamicLists = Get-DynamicDistributionGroup |
Where-Object { $_.DisplayName -like $nameFilter } |
Select-Object DisplayName, PrimarySmtpAddress, WhenCreated, ManagedBy, @{Name = "GroupType"; Expression = { "Dynamic Distribution List" } }
$allGroups += $DynamicLists

# Mail-enabled Security Groups
$MailSecGroups = Get-DistributionGroup -RecipientTypeDetails MailUniversalSecurityGroup |
Where-Object { $_.DisplayName -like $nameFilter } |
Select-Object DisplayName, PrimarySmtpAddress, WhenCreated, ManagedBy, @{Name = "GroupType"; Expression = { "Mail-Enabled Security" } }
$allGroups += $MailSecGroups

# Shared Mailboxes
$SharedMailboxes = Get-Mailbox -RecipientTypeDetails SharedMailbox |
Where-Object { $_.DisplayName -like $nameFilter } |
Select-Object DisplayName, PrimarySmtpAddress, WhenCreated, ManagedBy, @{Name = "GroupType"; Expression = { "Shared Mailbox" } }
$allGroups += $SharedMailboxes

# M365 Groups from Microsoft Graph
$M365Groups = @()
$MgGraphM365Groups = Get-MgGroup -Filter "groupTypes/any(c:c eq 'Unified')" -All | Where-Object { $_.DisplayName -like $nameFilter }

foreach ($group in $MgGraphM365Groups) {
    if ($group.DisplayName -like $nameFilter) {
        $owners = Get-MgGroupOwner -GroupId $group.Id
        $ownerNames = $owners | ForEach-Object {
            if ($_.UserPrincipalName -and $_.UserPrincipalName -ne "") {
                $_.UserPrincipalName
            }
            elseif ($_.DisplayName -and $_.DisplayName -ne "") {
                $_.DisplayName
            }
        } | Where-Object { $_ -ne $null -and $_ -ne "" }

        $M365Groups += [PSCustomObject]@{
            DisplayName        = $group.DisplayName
            PrimarySmtpAddress = $group.Mail
            WhenCreated        = $group.CreatedDateTime
            ManagedBy          = if ($ownerNames.Count -gt 0) { $ownerNames -join ", " } else { "" }
            GroupType          = "M365 Group"
        }
    }
}
$allGroups += $M365Groups
#endregion Collect EXO Data

#region Load existing SP list items
$existingItems = Get-PnPListItem -List $listName | ForEach-Object {
    @{
        ID        = $_.Id
        Title     = $_["Title"]
        Email     = $_["Email"]
        GroupType = $_["GroupType"]
        ManagedBy = $_["ManagedBy"]
    }
}
#endregion

#region Sync SharePoint List
foreach ($group in $allGroups) {
    $email = $group.PrimarySmtpAddress.ToString().ToLower()
    $match = $existingItems | Where-Object { $_.Email -eq $email }

    $managedBy = if ($group.ManagedBy -ne $null) {
        ($group.ManagedBy -join ", ")
    }
    else {
        ""
    }

    if ($match) {
        # Update if something changed
        if ($match.Title -ne $group.DisplayName -or $match.GroupType -ne $group.GroupType -or $match.ManagedBy -ne $managedBy) {
            Set-PnPListItem -List $listName -Identity $match.ID -Values @{
                Title      = $group.DisplayName
                GroupType  = $group.GroupType
                ManagedBy  = $managedBy
                LastSynced = (Get-Date)
            }
        }
    }
    else {
        # Add new item
        Add-PnPListItem -List $listName -Values @{
            Title       = $group.DisplayName
            Email       = $email
            GroupType   = $group.GroupType
            CreatedDate = $group.WhenCreated
            ManagedBy   = $managedBy
            LastSynced  = (Get-Date)
        }
    }
}

# Delete irrelevant entries from list
$groupEmails = $allGroups.PrimarySmtpAddress | ForEach-Object { $_.ToString().ToLower() }
$toRemove = $existingItems | Where-Object { $_.Email -notin $groupEmails }

foreach ($item in $toRemove) {
    Remove-PnPListItem -List $listName -Identity $item.ID -Force
}
#endregion

#region Disconnect
Disconnect-ExchangeOnline -Confirm:$false
#endregion

Stop-Transcript