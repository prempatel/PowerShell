Write-Host "Deleting listener that accepts requests..."
winrm enumerate winrm/config/listener
winrm delete winrm/config/listener?address=*+transport=HTTP

Write-Host "Disabling Windows RM firewall exceptions"
netsh advfirewall firewall set rule group="Windows Remote Management" new enable=no

Write-Host "Stopping Windows Remote Management service"
Stop-Service winrm
Set-Service -Name winrm -StartupType Disabled

Set-ItemProperty -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System -Name LocalAccountTokenFilterPolicy -Value 0 -Type DWord

Write-Host "Done cleaning PSRemoting leftovers."