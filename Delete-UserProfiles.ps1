<#
Deletes user profiles for all local users.
#>

#Parameters for ComputerName and Excluding users from deletion
Param (
    [Parameter(Mandatory=$true)][String[]]$ComputerName,
    [String[]]$ExcludeUsers = @('administrator','admintb')
)

foreach( $computer in $ComputerName){

    # Get a list of all user profiles
    $users = Get-WmiObject Win32_UserProfile -ComputerName $computer;

    Write-Host "Cleaning up user profiles..." -ForegroundColor Magenta

    foreach( $user in $users ) {
    
        # Normalize profile name.
        $userPath = (Split-Path $user.LocalPath -Leaf).ToLower()
    
        # If the profile name was found in the excluded list, skip it
        if( $ExcludeUsers -contains $userPath ) {
            Write-Host "Skipping $userPath from ignore list.";
            continue;
        }

        # Skip account if it is currently in use
        if( $userPath -eq $env:username ) {
            Write-Host "Skipping $userPath because it belongs to the current user.";
            continue;
        }

        # If the profile belongs to a "special" account (network/system), skip it.
        if( $user.Special -eq $true ) {
            Write-Host "Skipping $userPath because it is a special system account.";
            continue;
        }

        Write-Host "Deleting profile for $userPath..." -ForegroundColor Green;
        $user.Delete();
    }
}