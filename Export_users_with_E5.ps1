Connect-MsolService


Get-MsolUser -All | Where-Object { ($_.licenses).AccountSkuId -match "SPE_F1" -and $_.UsageLocation -eq "PL" } | Select-Object UserPrincipalName, Title | Export-Csv -Path "path to csv"

Get-MsolUser -All | Where-Object { $_.UsageLocation -eq "RO" -and ($_.licenses).AccountSkuId -match "EnterprisePremium" } | Select-Object UserPrincipalName, Title | Export-Csv -Path "path to csv"


Get-MsolUser -UserPrincipalName "UPN" | select *


Get-MsolSubscription # - lists all available subscriptions in tenant.


# SPE_F1 - Microsoft 365 F3
# EnterprisePremium - Office 365 E5