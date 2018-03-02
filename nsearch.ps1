
#Instalamos Nsearch
if(Test-Path -Path 'C:\Nsearch')
{
    if(Test-Path -Path 'C:\Nsearch\PSSQLite')
    {
        if(Test-Path -Path 'C:\Nsearch\db')
        {
        }
    }
}
else
{
    [string]$location = (Get-Location).Path
    $location = "$location" + "\*"
    New-Item -Path 'C:\' -Name Nsearch -ItemType Directory | Out-Null
    Copy-Item -Path $location -Destination 'C:\Nsearch' -Force
}

#Verificamos si Nmap esta instalado
$envpathnmap = $env:Path.Split(";")
$envpathnmap = $envpathnmap | Select-String -Pattern "Nmap$"
if(Test-Path -Path "$envpathnmap")
{
    Clear-Host
}
else
{
    Write-Host "No se encontro Nmap" -ForegroundColor Red -BackgroundColor Black
    switch($opc){
        1{
            $global:envpathnmap = Read-Host -Prompt "Path Nmap"    
        }
        2{
            Write-Host "Download Nmap and WinPcap" -ForegroundColor Green
            $download = New-Object System.Net.WebClient    
            $a.DownloadFile("https://nmap.org/dist/nmap-7.40-setup.exe","C:\test\nmap-7.40-setup.exe")
            $a.DownloadFile("https://www.winpcap.org/install/bin/WinPcap_4_1_3.exe","C:\test\WinPcap_4_1_3.exe")
            Write-Host "Descargos" -ForegroundColor Green
            Write-Host
        }
    }
    
    
    
}


Write-Host "========================================"
Write-Host " _   _  _____                     _     "
Write-Host "| \ | |/ ____|                   | |    "
Write-Host "|  \| | (___   ___  __ _ _ __ ___| |__  "
Write-Host "| . ` |\___ \ / _ \/ _` | '__/ __| '_ \ "
Write-Host "| |\  |____) |  __/ (_| | | | (__| | | |"
Write-Host "|_| \_|_____/ \___|\__,_|_|  \___|_| |_|"
Write-Host "========================================"
Write-Host "Version 0.1              @kcah27"
Write-Host ""

#Funcion para cambiar el Prompt
Function Global:prompt{
        Write-Host ("NS>") -nonewline -foregroundcolor Magenta
        return " "
    }

Function ayuda{
    param(
        [switch]$search,
        [switch]$doc,
        [switch]$templates,
        [switch]$run,
        [switch]$update
    )

    Begin{}
    
    Process{

        if($search)
        {
            Write-Host ""
            Write-Host "Use Example: " -ForegroundColor Cyan
            Write-Host "search -name http" -ForegroundColor Green
            Write-Host "search -category ex..[autocomplete]" -ForegroundColor Green
            Write-Host "search -author Pau..[autocomplete]" -ForegroundColor Green
        }
        elseif($doc)
        {
            Write-Host ""
            Write-Host "Use Example: " -ForegroundColor Cyan
            Write-Host "doc -name ht..[autocomplete]"  -ForegroundColor Green
        }
        elseif($templates){
            Write-Host ""
            Write-Host "Use Example: " -ForegroundColor Cyan
            Write-Host "Template" -ForegroundColor Green
        }
        elseif($run)
        {
            Write-Host ""
            Write-Host "Use Example: " -ForegroundColor Cyan
            Write-Host "run" -ForegroundColor Green
        }
        elseif($update)
        {
            Write-Host ""            
            Write-Host "Update nmap scripts" -ForegroundColor Green
        }
        else{
            Write-Host ""
            Write-Host "Comandos disponibles: " -ForegroundColor Cyan
            Write-Host "search     doc     templates" -ForegroundColor Green
            Write-Host "run        update" -ForegroundColor Green
            Write-Host ""   
        }
    }
    End{}    
}

Function search{
    
    param(
        [string]$nombre,
        [ValidateSet("auth","broadcast","brute","default","discovery","dos","exploit","external","fuzzer","intrusive","malware","safe","version","vuln")]
        [string]$categoria,
        [string]$autor
    )

    Begin{
        $envpathnmap = $env:Path.Split(";")
        $envpathnmap = $envpathnmap | Select-String -Pattern "Nmap$"
        $envpathnmap = $envpathnmap.ToString() + "\scripts"
        $db = Get-Content -Path ("$envpathnmap" + "\script.db")    
    }
    
    Process{
        foreach($line in $db)
        {
            $line = $line.replace('Entry { filename = "',"").replace('", categories = { "',',"').replace('", } }','"').replace('", "','","')
            $line = $line.Split(",")
    
            $pathdoc = ("$envpathnmap" + "\" + "$($line[0])")
            $doc = Get-Content -Path $pathdoc
            
            foreach($l in $doc)
            {
                if($l.StartsWith("author"))
                    {
                        $author = $l.replace('author = "',"").replace('"',',"').replace('[[',"").replace(',"',"").replace("author =","Brandon Enright <bmenrigh@ucsd.edu>, Duane Wessels <wessels@dns-oarc.net>")    
                        #$author
                    }
            }
            if(($line[0] -match "$nombre") -and ($line[0] -match "$categoria") -and ($author -match "$autor"))
            {
                $line[0]
            }
        }
    
    }

    End{}

}

Function doc{
    [CmdletBinding()]
    param()
    DynamicParam {
        $envpathnmap = $env:Path.Split(";")
        $envpathnmap = $envpathnmap | Select-String -Pattern "Nmap$"
        $envpathnmap = $envpathnmap.ToString() + "\scripts"
        
        # Set the dynamic parameters' name
        $ParameterName = 'namescript'
            
        # Create the dictionary 
        $RuntimeParameterDictionary = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary

        # Create the collection of attributes
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
            
        # Create and set the parameters' attributes
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttribute.Mandatory = $true
        $ParameterAttribute.Position = 1

        # Add the attributes to the attributes collection
        $AttributeCollection.Add($ParameterAttribute)

        # Generate and set the ValidateSet 
        $arrSet = Get-ChildItem -Path $envpathnmap | Select-Object -ExpandProperty Name
        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)

        # Add the ValidateSet to the attributes collection
        $AttributeCollection.Add($ValidateSetAttribute)

        # Create and return the dynamic parameter
        $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($ParameterName, [string], $AttributeCollection)
        $RuntimeParameterDictionary.Add($ParameterName, $RuntimeParameter)
        return $RuntimeParameterDictionary
    }

    Begin{
        $namescript = $PsBoundParameters[$ParameterName]
        $pathscript = "$envpathnmap" +"\" + "$namescript"
    }

    Process{
        $script = Get-Content -Path "$pathscript"
        foreach ($line in $script)
        {
            if($line.StartsWith("license"))
            {
                break
            }
        Write-Host $line -ForegroundColor Green
        }
    }

    End{}

}

Function templates{

    Begin{
    
    }

    Process{
    
    }

    End{
    
    }
}

Function update{
    Begin{}
    Process{
        if(([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))
        {
            nmap --script-updatedb
        }
        else
        {
            Write-Host "Please run this script with admin priviliges" -ForegroundColor Red -BackgroundColor Black
        }
    }
    End{}
}

#Funcion para regresar el prompt original
Function quit{
            
    Begin{
        #Clear-Host
    }

    Process
    {
        Function Global:prompt{"PS " + "$(Get-Location)" + "> "}
    }

    End{
        Write-Host "Saliendo Nsearch :D ..." -ForegroundColor Yellow
        Start-Sleep 2
        Clear-Host

        #Removemos las funciones construidas
        Remove-Item -Path Function:\ayuda
        Remove-Item -Path Function:\search
        Remove-Item -Path Function:\doc
        Remove-Item -Path Function:\update
        Remove-Item -Path Function:\quit
    }
}

