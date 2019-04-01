$gateway = "10.0.120.1"
$cidr = 24
$dns1 = "10.0.120.131"
$dns2 = "10.0.120.132"
$type = "IPv4"

$adapters = Get-NetAdapter
$totalAdapters = $adapters | Measure-Object | select -ExpandProperty Count -Wait

    if($totalAdapters -gt 1){
        Write-Host "There are $totalAdapters Network Adapters available" -ForegroundColor Magenta
        $title = "Network Adapters"
        $message = "SELECT A NETWORK ADAPTER"
        $default = 0

        $options = @()
        for($i=0; $i -lt $totalAdapters; $i++){
            $choice = [System.Management.Automation.Host.ChoiceDescription]::new("&$i - " + $adapters[$i].Name + " | " + $adapters[$i].InterfaceDescription)
            $options+=$choice
            Clear-Variable choice
        }

        $prompt = $host.ui.PromptForChoice($title, $message, $options, 0)
        $index = $adapters[$prompt].InterfaceIndex
    }
    else{
        $index = $adapters.InterfaceIndex
    }

Read-Host -Prompt "Enter IP Address"
$adapter = Get-NetAdapter -InterfaceIndex $index
$adapter | New-NetIPAddress -AddressFamily $type -IPAddress $ip -PrefixLength $cidr -DefaultGateway $gateway