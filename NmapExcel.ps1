Function Export-NmapExcel
{
    [CmdletBinding()]
    Param
    (
        # Nmap XML output file.
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0,
                   ParameterSetName = "File")]
        [ValidateScript({Test-Path $_})] 
        $NmapXML,

        # XML Object containing Nmap XML information
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0,
                   ParameterSetName = "XMLDoc")]
        [xml]$InputObject
    )

    Begin{}

    Process{
        
        if ($NmapXML)
        {
            $file = Get-ChildItem $NmapXML
            [xml]$nmap = [System.IO.File]::ReadAllText($file.FullName)
        }
        else
        {
            [xml]$nmap = $InputObject
        }

            #Creacion de la Aplicacion Excel
             $excel = New-Object -ComObject Excel.Application
             $excel.Visible = $true
             $excel.DisplayAlerts = $false
             $Book = $excel.Workbooks.Add()

            #Empieza el Parse del arbol XML
            $discoveredhosts = $nmap.nmaprun.host | Where-Object {$_.status.state -eq "up"}

            #Pestaña de ScanInfo            
            $PrimerHoja = $Book.WorkSheets.item(1)
            $PrimerHoja.Name = "ScanInfo"
            $PrimerHoja.Activate() | Out-Null
              
            #Titulos para ScanInfo
            $PrimerHoja.Cells.item(1,1) = "NmapVersion"
            $PrimerHoja.Cells.item(2,1) = "Command"
            $PrimerHoja.Cells.item(3,1) = "StartTime"
            $PrimerHoja.Cells.item(4,1) = "EndTime"
            $PrimerHoja.Cells.item(5,1) = "RunTime"
            $PrimerHoja.Cells.item(6,1) = "ScanType"
            $PrimerHoja.Cells.item(7,1) = "ScanProtocol"
            $PrimerHoja.Cells.item(8,1) = "NumberofServices"
            $PrimerHoja.Cells.item(9,1) = "Services"
            $PrimerHoja.Cells.item(10,1) = "DebugLevel"
            $PrimerHoja.Cells.item(11,1) = "VerboseLevel"
            $PrimerHoja.Cells.item(12,1) = "Summary"
            $PrimerHoja.Cells.item(13,1) = "ExitStatus"
                        
            #Asignacion de informacion en ScanInfo
            $PrimerHoja.Cells.item(1,2) = $nmap.nmaprun.version
            $PrimerHoja.Cells.item(2,2) = $nmap.nmaprun.args
            $PrimerHoja.Cells.item(3,2) = $scanstart
            $PrimerHoja.Cells.item(4,2) = $scanend
            $PrimerHoja.Cells.item(5,2) = $nmap.nmaprun.runstats.finished.elapsed
            $PrimerHoja.Cells.item(6,2) = $nmap.nmaprun.scaninfo.type
            $PrimerHoja.Cells.item(7,2) = $nmap.nmaprun.scaninfo.protocol
            $PrimerHoja.Cells.item(8,2) = $nmap.nmaprun.scaninfo.numservices
            $PrimerHoja.Cells.item(9,2) = $nmap.nmaprun.scaninfo.services
            $PrimerHoja.Cells.item(10,2) = $nmap.nmaprun.debugging.level
            $PrimerHoja.Cells.item(11,2) = $nmap.nmaprun.verbose.level
            $PrimerHoja.Cells.item(12,2) = $nmap.nmaprun.runstats.finished.summary
            $PrimerHoja.Cells.item(13,2) = $nmap.nmaprun.runstats.finished.exit
                                        
            #Creamos las Hojas
            $WorksheetCount = $discoveredhosts.Count
            $MissingType = [System.Type]::Missing
            $null = $Excel.Worksheets.Add($MissingType, $Excel.Worksheets.Item($Excel.Worksheets.Count), 
                    $WorksheetCount - $Excel.Worksheets.Count, $Excel.Worksheets.Item(1).Type)
            
            #Contador para las hojas
            $x = 2

            #Contador para Write-Progress
            $i = 0
   
                    ForEach($dischost in $discoveredhosts)
                    {
                    
                    #Control de Write-Progress
                    $i++

                    Write-Progress -Activity "Espera un momento" -Status "Progreso..." -PercentComplete ($i/$WorksheetCount*100)
                        
                        $NombreHoja = $dischost.address.addr
                        $Hoja = $Book.Worksheets.Item($x)
                        $Hoja.Name = $NombreHoja
                        $Hoja.Cells.Item(1,1) = "Puerto"
                        $Hoja.Cells.Item(1,2) = "Servicio"
                        $Hoja.Cells.Item(1,3) = "Banner"


                        
                        #Seleccionamos los puertos abiertos
                        $puertos = $dischost.ports.port | Where-Object {$_.state.state -eq "open"}
                        
                        #Contador de las filas
                        $p = 2

                        foreach($port in $puertos)
                        {
                            $Hoja.Cells.Item($p,1) = $port.portid
                            $Hoja.Cells.Item($p,2) = $port.service.name
                            $Hoja.Cells.Item($p,3) = $port.service.Product
                            $p++
                        }
                        $x++
                  
                    }

        
        }
    
    
    

    End{}

}