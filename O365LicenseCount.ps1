$pass =  Get-Content -Path "C:\Users\ppatel-adm\Desktop\encrypt.txt" | ConvertTo-SecureString
$cred = New-Object System.Management.Automation.PsCredential("ppatel@takarabelmont.com", $pass)

Connect-MsolService -Credential $cred
Get-MsolAccountSku | where {$_.AccountSkuID -eq "takarabelmont:ENTERPRISEPACK"}


<#
$licenses = Get-MsolAccountSku | where {$_.AccountSkuID -eq "takarabelmont:ENTERPRISEPACK"}
$active = $licenses | select -Expand activeunits
$consumed = $licenses | select -Expand consumedunits
$available = $active - $consumed
write-host "Available Licenses" $available
#>