Function Get-AntiVirusProduct { 

    [CmdletBinding()] 
    Param ( 
        [parameter(ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)] 
        [Alias(‘name’)] 
        $computername=$env:computername 
    )

    $AntiVirusProduct = Get-WmiObject -Namespace root\SecurityCenter2 -Class AntiVirusProduct  -ComputerName $computername

     
    $ht = [ordered]@{
    'ComputerName' = $computername
    'Name' = $AntiVirusProduct.displayName
    'ProductExecutable' = $AntiVirusProduct.pathToSignedProductExe
    'GUID' = $AntiVirusProduct.instanceGuid
    'Status' = $AntiVirusProduct.productstate
    } 

    
    $obj = New-Object -TypeName PSObject -Property $ht
    Write-Output $obj


}


