<# exchange dynamic distribution list group troubleshoot commands
#>


Connect-ExchangeOnline

Get-DynamicDistributionGroup -Identity "full email" | Select-Object *recipi*

Get-DynamicDistributionGroupMember -Identity "full email" # - sprawdza członków

$FTE = Get-DynamicDistributionGroup -Identity "full email"   # - preview członków grupy, można wykonać przed Get-DynamicDistributionGroupMember, nie pokazuje to faktycznych członków
Get-Recipient -RecipientPreviewFilter ($FTE.RecipientFilter) | ft displayname, recipienttype, title, primarysmtpaddress  # - preview członków grupy, można wykonać przed Get-DynamicDistributionGroupMember, nie pokazuje to faktycznych członków

Remove-DynamicDistributionGroup -Identity 'full email' -Confirm:$false # - usunięcie grupy

New-DynamicDistributionGroup -Name "Group Name" -DisplayName "Group DisplayName" -PrimarySmtpAddress "full email" -RecipientFilter $Filter11 # - utworzenie nowej grupy

Set-DynamicDistributionGroup -Identity "full email" -RecipientFilter $PLTLDataEngineers2  # - zmiana recipient filtru

Set-DynamicDistributionGroup -Identity "full email" -ForceMembershipRefresh  # - wymuszenie odświeżenia listy członków

Get-DistributionGroupMember -Identity "Room Resource Name" # - sprawdza członków tzw. Room list w Outlook
Add-DistributionGroupMember -Identity "Room Resource Name" -Member "NewRoom@yourdomain.com" # - dodanie nowego członka do Room list w Outlook