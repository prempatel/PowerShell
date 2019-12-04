#Auto refresh URL and reopen URL if closed

$link = 'http://nj-vsql-01/Reports/Pages/Report.aspx?ItemPath=%2fGP+Custom+Reports%2fMO+Production+Schedule'

Start-Process iexplore.exe -ArgumentList '-k ', $link
$wshell = New-Object -ComObject wscript.shell

while($true){
    do{
        Start-Sleep 5400
        $wshell.AppActivate('Internet Explorer')
        $wshell.SendKeys('{F5}')
    } while ((Get-Process iexplore -ErrorAction SilentlyContinue).Count -ge 1)

    Start-Process iexplore.exe -ArgumentList '-k', $link
}