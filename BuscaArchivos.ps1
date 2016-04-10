    #Hecho por Mario Alberto Sanchez Badillo
    #Oficina de seguridad KIONETWORKS
Function Search-Files{
    
    <#
      .SYNOPSIS
      Busca Archivos

      .DESCRIPTION

      Funcion que nos apoya en la busqueda de archivos, por fecha de inicio y fin; en una
      determinada Ruta que como usuario le asignemos.
      .EXAMPLE

      Search-Files -Inicio 02/20/2016 -Fin 02/26/2016 -Ruta C:\Users\Usuario\Desktop -Resultados C:\Users\Usuario\Desktop\Resultados
    #>


    [CmdletBinding()]

    Param(
    [Parameter(Mandatory=$True,
                ValueFromPipeline=$True,
                ValueFromPipelineByPropertyName=$True)]
    [datetime]$Inicio,

    [Parameter(Mandatory=$True,
                ValueFromPipeline=$True,
                ValueFromPipelineByPropertyName=$True)]
    [datetime]$Fin,
    
    
    [Parameter(Mandatory=$True,
                ValueFromPipeline=$True,
                ValueFromPipelineByPropertyName=$True)]
    $Ruta,

    [Parameter(Mandatory=$True,
                ValueFromPipeline=$True,
                ValueFromPipelineByPropertyName=$True)]
    $Resultados,

    [Parameter(Mandatory=$false,
                ValueFromPipeline=$True,
                ValueFromPipelineByPropertyName=$True)]
    [String]$Extencion
     )

     Begin{}

     Process{
     Get-ChildItem -Path $Ruta $Extencion -Recurse -Force  | where {($_.LastWriteTime -gt $Inicio) -and ($_.LastWriteTime -lt $Fin)} | Out-File -Width 255 $Resultados
     $Directorios =  Get-ChildItem -Path $Ruta -Recurse -Force $Extencion | where {($_.LastWriteTime -gt $Inicio) -and ($_.LastWriteTime -lt $Fin)} 
     $Directorios.DirectoryName
     }

     End{
        notepad.exe $Resultados
     }
    
}
