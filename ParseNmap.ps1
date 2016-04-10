function Nmap-Report
{
Param(
    $Ruta
)

Begin{}

Process{    

[xml]$Nmap = Get-Content $Ruta

$discoveredhosts = $nmap.nmaprun.host | Where-Object {$_.status.state -eq "up"}

foreach($hosts in $discoveredhosts)
    {
    foreach ($address in $hosts.address | ForEach-Object {if($_.addrtype -eq "ipv4"){$_.addr}})
        {
        Write-Host "----------------------------------------------------------------------------"
        Write-Host "Remote Hosts" $address
            Write-Host ""
            foreach($os in $hosts.os.osmatch | Format-Table Name,accuracy -AutoSize)
            {
                $os
            }
            foreach($ports in $hosts.ports.port)
            {
                $ports | Where-Object {$_.state.state -contains "open"} |
                Select-Object -Property Protocol,PortId, @{Name="Status"; Expression={$_.state.state}},@{Name="Servicio"; Expression={$_.service.name}}

            } 
    
        }
    }
}

End{}

} 