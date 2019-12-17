# Reverse DNS Lookup in our domain for A records
 
$searchIP = Read-Host "Enter IP Address"
$record = Get-DnsServerResourceRecord -ZoneName "ZONE_NAME" -ComputerName "DOMAINCONTROLLER" -RRType "A" | where {$_.RecordData.IPV4Address -eq $searchIP} | select hostname, @{ Name = 'IPAddress'; Expression = { $_.RecordData.IPV4Address } }

$hostname = $record | select -Expand hostname
$ip = $record | select -Expand IPAddress
$description = Get-ADComputer $hostname -Properties Description | select -Expand Description

Write-Host "Computer: " $hostname
Write-Host "IP: " $ip
Write-Host "User: " $description

#Gridview Version of All DNS Records
#Get-DnsServerResourceRecord -ZoneName "takarabelmont.com" -ComputerName nj-vdc-01 -RRType "A" | select Hostname, @{ Name = 'IP Address'; Expression = { $_.RecordData.IPV4Address } } | Out-GridView
