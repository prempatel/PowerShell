#Select the group you want
$group = get-adgroup WF_NJ_SYNC_Users-Desktop

#Get all the users with memberof property, and filter using Where-Object DistinguishedName not in MemberOf

get-aduser -SearchBase "<OU>" -Properties memberof -filter 'enabled -eq $true' | Where-Object {$group.DistinguishedName -notin $_.memberof} | select name | Export-Csv "C:\Users\ppatel-adm\Desktop\WorkFolders.csv"
