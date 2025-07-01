Connect-MgGraph -Scopes "User.Read.All"

$userId = "UPN"
Get-MgUser -UserId $userId -Property onPremisesExtensionAttributes | select -ExpandProperty onPremisesExtensionAttributes #| select -ExpandProperty ExtensionAttribute12

$newAttribute12Value = "new value for ExtensionAttribute12"
$extensionAttributes = @{ExtensionAttribute12 = $newAttribute12Value}
Update-MgUser -UserId $userId -OnPremisesExtensionAttributes $extensionAttributes