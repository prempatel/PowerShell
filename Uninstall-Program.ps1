$Application = Read-Host -Prompt 'Program Name?'

$uninstall32 = gci "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall" | foreach { gp $_.PSPath } | ? { $_ -match $Application } | select -ExpandProperty UninstallString
$uninstall64 = gci "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall" | foreach { gp $_.PSPath } | ? { $_ -match $Application } | select -ExpandProperty UninstallString

if ($uninstall64) {
$uninstall64 = $uninstall64.UninstallString -Replace "msiexec.exe","" -Replace "/I","" -Replace "/X",""
$uninstall64 = $uninstall64.Trim()

Write "Uninstalling x64..."
start-process "msiexec.exe" -arg "/X $uninstall64 /qb" -Wait}

if ($uninstall32) {
$uninstall32 = $uninstall32.UninstallString -Replace "msiexec.exe","" -Replace "/I","" -Replace "/X",""
$uninstall32 = $uninstall32.Trim()

Write "Uninstalling x32..."
start-process "msiexec.exe" -arg "/X $uninstall32 /qb" -Wait}