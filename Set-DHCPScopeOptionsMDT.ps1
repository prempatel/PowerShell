#Set DHCP Scope Option 67 (bootfile name) depending if PC is BIOS/UEFI for MDT deployments

[int]$choice = 0
while ( $choice -lt 1 -or $choice -gt 2 ) {
    Write-Host "#### SET MDT DHCP SCOPE OPTIONS ####" -ForegroundColor green
    Write-Host ""
    Write-host "1. UEFI (boot\x64\wdsmgfw.efi)"
    Write-host "2. BIOS (boot\x64\wdsnbp.com)"

    Write-Host ""
    [Int]$choice = Read-Host "Please enter an option (1/2)"
    Write-Host ""
}

#Assign selected department to variable
Switch ( $choice ) {
    1 {$ScopeValue = "boot\x64\wdsmgfw.efi"}
    2 {$ScopeValue = "boot\x64\wdsnbp.com"}
}


Set-DhcpServerv4OptionValue -ComputerName <DHCP_SERVER> -ScopeId <SCOPE_ID> -OptionId 67 -Value $ScopeValue

# To confirm if Scope Option was set
# Get-DhcpServerv4OptionValue -ComputerName <DHCP_SERVER> -ScopeId <SCOPE_ID> -OptionId 67
