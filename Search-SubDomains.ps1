Function Search-SubDomains{
    <#

    .SYNOPSIS
        Script que automatiza la busqueda de subdominios en fuentes publicas

    .DESCRIPTION
        Script que automatiza la busqueda de subdominios en fuentes publicas como:
        www.google.com, www.bing.com, www.virustotal sin la necesidad de un APIKEY

    .EXAMPLE
        Search-SubDomains -D "dominio" -google
        Busqueda de subdominios en www.google.com
    
    .EXAMPLE
        Search-SubDomains -D "dominio" -bing
        Busqueda de subdominios en el buscador de Microsoft Bing

    .EXAMPLE
        Search-SubDomains -D "dominio" -virustotal
        Busqueda de subdominios en www.VirusTotal.com
        Warning: Solo acepta 3 busquedas, ya que se realiza sin APIKEY

    .EXAMPLE
        Search-SubDomains -D "dominio" -all
        Busqueda en Google, Bing y VirusTotal

    .EXAMPLE
        Search-SubDomains -D "dominio" -google -bing
        Busqueda multiple

    .EXAMPLE
        Search-SubDomains -D "dominio" -google -bing | Out-GridView
        Search-SubDomains -D "dominio" -bing | Out-File -FilePath
        Search-SubDomains -D "dominio" -virustotal | Export-Csv -Path
        $variable = Search-SubDomains -D "dominio" -all
        Salida en multiples formatos y variables
    
    .NOTES
        Twitter: @kcah27
        mail: albertokcah27@outlook.com
    #>

    param(
        [Parameter(Mandatory=$true)]
        [string[]]$D,
        [switch]$google,
        [switch]$bing,
        [switch]$virustotal,
        [switch]$all
    )

    Begin{

        #Variables Globales
        $Global:UserAgent = "Mozilla/5.0 (Windows NT 10.0; WOW64; Trident/7.0; rv:11.0) like Gecko"

        #Funciones
        Function Search-Google{

        param(

            [string]$domain

        )

        Begin{
        
            $dominio = $domain
            $www = "www."+$dominio
            $MaxSearchPages = 5
            $googlesearch = "https://www.google.com/search?q=site:$dominio+-$www&num=100"
            $pagelinks = (Invoke-WebRequest -Uri $googlesearch -UserAgent $UserAgent -UseBasicParsing).Links
            $validlinks = @()
            $subdomains = @()
            $otherpages = @()
            $hrefs = @()
            $otherpageslimit = @()
            $morepagelinks = @()
            $morehrefs = @()
        }

        Process{

            foreach($links in $pagelinks)
            {
                $hrefs += $links.href
            }
            foreach ($url in $hrefs)
            {
                if(($url -like "*$dominio*") -and ($url -notlike "*https://www.google.com*") -and ($url -notlike "https://maps.google.com*") -and ($url -notlike "http://webcache.googleusercontent.com*") -and ($url -notlike "https://webcache.googleusercontent.com*") -and ($url -notlike "https://translate.google.com*"))
                {
                    if($url -like "*http://*")
                    {
                        $strippedurl = [regex]::Match($url,"http([^\)]+)").value
                        $validlinks += $strippedurl 
                    }
                    elseif($url -like "*https://*")
                    {
                        $strippedurl = [regex]::Match($url,"https([^\)]+)").value
                        $validlinks += $strippedurl 
                    }
                }
            }

        
            foreach($url in $hrefs) 
            {
                if ($url -like "*search?q*start=*")
                {
                    $otherpages += "https://www.google.com$url" + "&num=100"
                }
            }

            $otherpages = $otherpages | sort | unique
            $pagecount = $otherpages.count

            if ($pagecount -gt $MaxSearchPages)
            {
                $totalpagelimit = $MaxSearchPages
                for($j=0; $j -lt $totalpagelimit; $j++)
                {                                    
                    $otherpageslimit += $otherpages[$j]
                }
            }

            $otherpageslimit = $otherpageslimit -replace "&amp;","&"
            foreach($page in $otherpageslimit)
            {
                $i++
                $morepagelinks = (Invoke-WebRequest -Uri $page -UserAgent $UserAgent -UseBasicParsing).Links
                $morehrefs = $morepagelinks.href
                foreach($url in $morehrefs)
                {
                    if (($url -like "*$dominio*"))
                    {
                        if ($url -like "*http://*")
                        {
                            $strippedurl = [regex]::match($url,"https([^\)]+)").Value
                            $validlinks += $strippedurl
                        }
                        elseif ($url -like "*https://*")
                        {
                            $strippedurl = [regex]::match($url,"https([^\)]+)").Value
                            $validlinks += $strippedurl
                        }
                    }     
                }
            }
        }
        End{
            foreach ($valid in $validlinks)
            {
                $subdomain = [regex]::Split($valid,"\/")[2]
                $subdomains += $subdomain
            }
            #$subdomains | select -Unique

            Set-Variable -Name resultgoogle -Value $subdomains -Scope Global
        }

    }#Fin funcion Google
        
        Function Search-Bing{
    
        param(
        
            [string]$domain

        )

        Begin{

            $dominio = $domain
            $www = "www."+$dominio
            $bingsearch = "http://www.bing.com/search?q=site:$dominio&count=30"
            $bingpagelinks = @()
            $bingpagelinks = (Invoke-WebRequest -Uri $bingsearch -UserAgent $UserAgent -UseBasicParsing).Links
            $binghrefs = @()
            $bingotherpages = @()
            $morepagelinks = @()
            $morehrefs = @()
            $validlinks = @()
            $subdomains = @()

        }

        Process{

            foreach($links in $bingpagelinks)
            {
                $binghrefs += $links.href
            }

            foreach($url in $binghrefs)
            {
                if($url -like "*$dominio*")
                {
                    if($url -like "*http://*")
                    {
                        $strippedurl = [regex]::Match($url,"http([^\)]+)").value
                        $validlinks += $strippedurl 
                    }
                    elseif($url -like "*https://*")
                    {
                        $strippedurl = [regex]::Match($url,"https([^\)]+)").value
                        $validlinks += $strippedurl 
                    }
                }
            }
                

            foreach($url in $binghrefs)
            {
                if ($url -like "*search?q*first=*")
                {
                    $bingotherpages += "https://www.bing.com$url" + "&count=30"
                }
            }
                
            $bingotherpages = $bingotherpages | sort | unique
            $bingotherpages = $bingotherpages -replace "&amp;","&"
            foreach($page in $bingotherpages)
            {
                $morepagelinks = (Invoke-WebRequest -Uri $page -UserAgent $UserAgent -UseBasicParsing).Links
                $morehrefs = $morepagelinks.href
                foreach($url in $morehrefs)
                {
                    if (($url -like "*$dominio*"))
                    {
                        if ($url -like "*http://*")
                        {
                            $strippedurl = [regex]::Match($url,"http([^\)]+)").value
                            $validlinks += $strippedurl
                        }
                        elseif ($url -like "*https://*")
                        {
                            $strippedurl = [regex]::Match($url,"https([^\)]+)").value
                            $validlinks += $strippedurl
                        }
                    }
                }
            }
        }#Fin Process

        End{
        
            foreach ($valid in $validlinks)
            {
                $subdomain = [regex]::Split($valid,"\/")[2]
                $subdomains += $subdomain
            }
            #$subdomains | select -Unique

            Set-Variable -Name resultbing -Value $subdomains -Scope Global
        }

    }#Fin funcion Bing
        
        Function Search-VirusTotal{
    
        param(

            [string]$domain

        )

        Begin{

            $dominio = $domain
            $uri = "https://www.virustotal.com/en/domain/" + $dominio +"/information/"
            $UserAgent = "Mozilla/5.0 (Windows NT 10.0; WOW64; Trident/7.0; rv:11.0) like Gecko"
            $response = Invoke-WebRequest -Uri $uri -Method Get -UserAgent $UserAgent
            $term = $dominio.Split(".")
            $term = $term[1].ToString()
            $term = "." + "$term"
            $hrefs = @()
            $validlinks = @()
            $subdomains = @()
        }

        Process{
        
            foreach($links in $response.Links)
            {
                $hrefs += ($links.href)
            }

            foreach($url in $hrefs)
            {
                if($url -like "*$dominio*")
                {
                    if($url -like "*$term*")
                    {
                        $validlinks += $url.Replace('/information/','').Replace('/ca/domain/','').Replace('/en/domain/','').Replace('/da/domain/','').Replace('/de/domain/','').Replace('/es/domain/','').Replace('/fr/domain/','').Replace('/hr/domain/','').Replace('/it/domain/','').Replace('/hu/domain/','').Replace('/nl/domain/','').Replace('/nb/domain/','').Replace('/pt/domain/','').Replace('/pl/domain/','').Replace('/sk/domain/','').Replace('/uk/domain/','').Replace('/vi/domain/','').Replace('/tr/domain/','').Replace('/ru/domain/','').Replace('/sr/domain/','').Replace('/bg/domain/','').Replace('/he/domain/','').Replace('/ka/domain/','').Replace('/ar/domain/','').Replace('/fa/domain/','').Replace('/zh-cn/domain/','').Replace('/zh-tw/domain/','').Replace('/ja/domain/','').Replace('/ko/domain/','')     
                    }
                }
            }



        }#fin Process
    
        End{
            #$validlinks | select -Unique
            
            foreach($l in $validlinks){
                $subdomains += $l
            }

            Set-Variable -Name resultvt -Value $subdomains -Scope Global
        }

    } #Fin funcion Virus Total

    }#Fin Begin Funcion principal

    Process{
        if(($google -eq $false) -and ($bing -eq $false) -and ($virustotal -eq $false) -and ($all -eq $false)){
            Write-Host "Favor de indicar parametro de busqueda" -ForegroundColor Red
        }

        if($google){
            Search-Google -domain $D
            $resultgoogle | Select -Unique
        }

        elseif($bing){
            Search-Bing -domain $D
            $resultbing | Select -Unique
        }
    
        elseif($virustotal){
            Search-VirusTotal -domain $D
            $resultvt | Select -Unique
        }
        elseif(($all) -or (($google -eq $true) -and ($bing -eq $true) -and ($virustotal -eq $true))){
            Search-Google -domain $D
            Search-Bing -domain $D
            Search-VirusTotal -domain $D
            
            $resultall = @()
            foreach($z in $resultgoogle)
            {
                $resultall +=$z
            }
            
            foreach($z in $resultbing)
            {
                $resultall +=$z
            }
            
            foreach($z in $resultvt)
            {
                $resultall +=$z
            }
            $resultall | select -Unique
        }
        elseif(($google -eq $true) -and ($bing -eq $true)){
            Search-Google -domain $D
            Search-Bing -domain $D
            $resultall = @()
            foreach($z in $resultgoogle)
            {
                $resultall +=$z
            }
            
            foreach($z in $resultbing)
            {
                $resultall +=$z
            }
            $resultall | select -Unique
        }
        elseif(($google -eq $true) -and ($virustotal -eq $true)){
            Search-Google -domain $D
            Search-VirusTotal -domain $D
            $resultall = @()
            foreach($z in $resultgoogle)
            {
                $resultall +=$z
            }

            foreach($z in $resultvt)
            {
                $resultall +=$z
            }
            $resultall | select -Unique

        }
        elseif(($bing -eq $true) -and ($virustotal -eq $true)){
            Search-Bing -domain $D
            Search-VirusTotal -domain $D
            $resultall = @()
            foreach($z in $resultbing)
            {
                $resultall +=$z
            }
            
            foreach($z in $resultvt)
            {
                $resultall +=$z
            }
            $resultall | select -Unique
        }

    }#Fin Process Funcion principal

    End{
    
    }#End de la Funcion principal

}