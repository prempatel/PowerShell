#Get Computer IP assigned via DHCP
#Issues when DNS not updated for WiFi/Ethernet IPs

$computer = Read-Host "Enter Computer Name"
while([string]::IsNullOrWhitespace($computer)){
    $computer = Read-Host "Enter Computer Name"
}

$computer = $computer + '.takarabelmont.com'

# Location Selection
[int]$LocationChoice = 0
while ( $LocationChoice -lt 1 -or $LocationChoice -gt 11 ) {
    Write-Host "#### REGION SELECTION ####" -ForegroundColor green
    Write-Host ""
    Write-host "1. New Jersey"
    Write-host "2. Illinois"
    Write-host "3. Kansas"
    Write-host "4. Missouri"
    Write-host "5. New Jersey"
    Write-host "6. New York"
    Write-host "7. Texas"

    Write-Host ""
    [Int]$LocationChoice = Read-Host "Enter region option 1 - 7" 
    Write-Host ""
}

#Assign selected location to variable
Switch ( $LocationChoice ) {
    1 {$Location = "NJ"}
    2 {$Location = "IL"}
    3 {$Location = "KS"}
    4 {$Location = "MO"}
    5 {$Location = "CA"}
    6 {$Location = "NY"}
    7 {$Location = "TX"}
}


Switch ( $LocationChoice ) {
    1 {$Scope = "10.0.110.0"}
    2 {$Scope = "10.30.0.0"}
    3 {$Scope = "10.60.0.0"}
    4 {$Scope = "10.10.110.0"}
    5 {$Scope = "10.70.0.0"}
    6 {$Scope = "10.20.0.0"}
    7 {$Scope = "10.50.0.0"}
}

Get-DhcpServerv4Lease -ComputerName "$Location-VDC-01" -ScopeId $Scope | Select-Object HostName, IPAddress | Where-Object {$_.HostName -match $computer}

