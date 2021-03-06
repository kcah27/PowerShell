
function Get-Hardening {
    [cmdletBinding()]	
    
    param (
 
        [Parameter(Mandatory=$false,
               ValueFromPipelineByPropertyName=$true,
               Position=0)]
        [ValidateSet("local","remoto")] 
        [string]
        $Entorno = "local",
        
        [Parameter(Mandatory=$false,
        ValueFromPipelineByPropertyName=$true,
        Position=1)] 
        [switch]
        $Reporte
    
    )
    
	    BEGIN 
		    {#Inicia BEGIN
                if( $Entorno -eq "local")
                {
                    $nameComputer= "$env:COMPUTERNAME" 
                }
                else
                {
                $nameComputer = $(Read-Host "Nombre de equipo o IP")
                $credential = $(Get-Credential)
			    New-PSSession -ComputerName $nameComputer -Credential $credential
                $session = Get-PSSession
                $s = $session.id
                $session = Get-PSSession -Id $s
                }
             #-------------------------------------------------------------------Reporte Local------------------------------------------------------------------------------------
             
             $scriptReportetxt = { 

Write-Host ============================================================================ -ForegroundColor green
Write-Host "                          INFORMACION GENERAL                              " -ForegroundColor green
Write-Host ============================================================================ -ForegroundColor green
Write-Host ""

Write-Host TIMESTAMP -ForegroundColor blue
Write-Host ========= -ForegroundColor blue
    Get-Date
Write-Host ""

Write-Host SO / SERVICE PACK / ARQUITECTURA -ForegroundColor blue
Write-Host ================================ -ForegroundColor blue
    $sOS =Get-WmiObject -class Win32_OperatingSystem
    $sOS | Select-Object Description, Caption, OSArchitecture, ServicePackMajorVersion | Format-List | Out-Host
Write-Host ""

Write-Host INFORMACION DEL SERVIDOR -ForegroundColor blue
Write-Host ======================== -ForegroundColor blue
    $cim = Get-Command -Name *Get-Cim*
    if($cim -eq $null){Write-Host "Informacion no disponible" -ForegroundColor red -BackgroundColor black}
    else {Get-CimInstance Win32_OperatingSystem | FL * | Out-Host}
Write-Host ""

Write-Host INFORMACION INTERNET EXPLORER -ForegroundColor blue
Write-Host ============================= -ForegroundColor blue
    (Get-ItemProperty "HKLM:\Software\Microsoft\Internet Explorer").SvcVersion
Write-Host ""

Write-Host CONFIGURACION DEL PROXY LOCAL -ForegroundColor blue
Write-Host ============================= -ForegroundColor blue
    $Proxy = (Get-ItemProperty "Registry::HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings").ProxyEnable
    if( $Proxy -eq 0 ){Write-Host "Proxy no configurado"}
    else {Write-Host "Proxy configurado"}
Write-Host ""

Write-Host SISTEMA DE ARCHIVOS DE LAS UNIDADES DE DISCO -ForegroundColor blue
Write-Host ============================================ -ForegroundColor blue
    [System.IO.DriveInfo]::getdrives()
Write-Host ""

Write-Host ANTIMALWARE / ANTIVIRUS -ForegroundColor blue   
Write-Host ======================= -ForegroundColor blue
$wmi = gwmi -namespace "root" -class "__Namespace" | where {$_.name -eq "SecurityCenter2" } | select name
if($wmi -eq $null){Write-Host "No es posible verificar" -ForegroundColor red -BackgroundColor black}
else{
$wmiQuery = "SELECT * FROM AntiVirusProduct"
$AntivirusProduct = Get-WmiObject -Namespace "root\SecurityCenter2" -Query $wmiQuery	   	 
Write-host $AntivirusProduct.displayName -ForegroundColor Cyan
}
Write-Host ""

Write-Host OPCIONES DE SEGURIDAD -ForegroundColor blue
Write-Host ===================== -ForegroundColor blue
    Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters" | select NullSessionPipes, autodisconnect, enableforcedlogoff, enablesecuritysignature, requiresecuritysignature, restrictnullsessaccess, AdjustedNullSessionPipes, EnableAuthenticateUserSharing | Out-Host
Write-Host ""

Write-Host ULTIMOS PARCHES DE SEGURIDAD -ForegroundColor blue
Write-Host ============================ -ForegroundColor blue
Write-Host ""
    #Get-HotFix -Description "Security*"
Write-Host ""

Write-Host CONFIGURACION DEL CLIENTE NTP -ForegroundColor blue
Write-Host ============================= -ForegroundColor blue
Write-Host "La direccion NTP del servidor esta establecida en:  " -NoNewline
    (Get-ItemProperty "HKLM:\SYSTEM\Currentcontrolset\Services\W32time\Parameters\").Ntpserver
Write-Host ""

Write-Host CONFIGURACION DEL FIREwALL LOCAL -ForegroundColor blue
Write-Host =============================== -ForegroundColor blue
    netsh advfirewall show domainprofile | Out-Host
    netsh advfirewall show privateprofile | Out-Host
    netsh advfirewall show publicprofile | Out-Host
Write-Host ""


Write-Host CONFIGURACION DE LA RED -ForegroundColor blue
Write-Host ===================== -ForegroundColor blue
    ipconfig /all | Out-Host
Write-Host ""

Write-Host Group Policies -ForegroundColor blue
Write-Host ============== -ForegroundColor blue
    gpresult /R
Write-Host ""

Write-Host SECURE BOOT -ForegroundColor blue
Write-Host =========== -ForegroundColor blue
    bcdedit
Write-Host ""
Write-Host "Verifica si el Boot del systema es por BIOS o UEFI"
Write-Host ""
if( (Test-Path -Path "HKLM:\System\CurrentControlSet\Control\SecureBoot\State") -eq $true)
{
    $SecureBoot = (Get-ItemProperty "HKLM:\System\CurrentControlSet\Control\SecureBoot\State").UEFISecureBootEnabled
    If ($SecureBoot -eq "1")
        {
            Write-Host "UEFI Secure Boot esta habilitado"
        }
    If ($SecureBoot -eq "0")
        {
            Write-Host "UEFI Secure Boot esta deshabilitado"
        }
}
else
{
    Write-Host "No existe llave del registro SecureBoot" -ForegroundColor Yellow -BackgroundColor Black
}
Write-Host ""

Write-Host LEGAL -ForegroundColor blue
Write-Host ===== -ForegroundColor blue
    Get-ItemProperty "HKLM:/SOFTWARE/Microsoft/Windows NT/CurrentVersion/Winlogon" | Select-Object LegalNoticeCaption, LegalNoticeText | Format-List | Out-Host
Write-Host ""

sleep 5

Write-Host ============================================================================ -ForegroundColor green
Write-Host "                                    USUARIOS                              " -ForegroundColor green
Write-Host ============================================================================ -ForegroundColor green
Write-Host ""

Write-Host INFORMACION DE CUENTAS LOCALES -ForegroundColor blue
Write-Host ============================== -ForegroundColor blue
[string]$variable = $true
    Get-WmiObject -Class Win32_UserAccount -Filter  "LocalAccount=$variable" | 
        Select-Object PSComputerName, Status, Caption, PasswordExpires, AccountType, Description, Disabled, Domain, FullName, InstallDate, LocalAccount, Lockout, Name, PasswordChangeable, PasswordRequired, SID, SIDType | Out-Host
Write-Host ""

Write-Host USUARIOS DE CADA GRUPO LOCAL -ForegroundColor blue
Write-Host ============================ -ForegroundColor blue

Write-Host "A continuacion se muestra la lista de todos los grupos locales y sus miebros en el servidor: "$server" server."
    $server = "$env:COMPUTERNAME"
    $computer = [ADSI]"WinNT://$server,computer"

    $computer.psbase.children | where { $_.psbase.schemaClassName -eq "group" } | foreach {
        write-host $_.name
        write-host "------"
        $group =[ADSI]$_.psbase.Path
        $group.psbase.Invoke("Members") | foreach {$_."GetType".Invoke().InvokeMember("Name", "GetProperty", $null, $_, $null)}
        write-host
    }
Write-Host ""

Write-Host ULTIMO ACCESO DE CUENTAS  -ForegroundColor blue 
Write-Host ================ -ForegroundColor blue
Write-Host "Note: Verifica las cuentas de los usuarios que no se han logueado en 90 dias."
Write-Host ""
    $([ADSI]"WinNT://$env:COMPUTERNAME").Children | where {$_.SchemaClassName -eq "user"} | Select-Object name, lastlogin | Out-Host
Write-Host ""

sleep 5

Write-Host ============================================================================ -ForegroundColor green
Write-Host "                                    PASSWORDS                              " -ForegroundColor green
Write-Host ============================================================================ -ForegroundColor green
Write-Host ""

Write-Host POLITICA DE PASSWORDS LOCAL -ForegroundColor blue
Write-Host =========================== -ForegroundColor blue
    net accounts | Out-Host
Write-Host ""

Write-Host POLITICA DE PASSWORDS DEL DOMINIOO -ForegroundColor blue
Write-Host ================================== -ForegroundColor blue
Write-Host "Note: Esto funcionara solo si el servidor esta conectado al Directorio Activo"
Write-Host ""   
    net accounts /domain | Out-Host

sleep 5
 
Write-Host ============================================================================ -ForegroundColor green
Write-Host "                                    RDP                                   " -ForegroundColor green
Write-Host ============================================================================ -ForegroundColor green
Write-Host ""

Write-Host CONEXIONES PERMITIDAS
Write-Host =====================
Write-Host ""

Function Get-RemoteDesktopConfig
    {if ((Get-ItemProperty -Path "HKLM:\System\CurrentControlSet\Control\Terminal Server").fDenyTSConnections -eq 1)

              {"Las conexiones RDP no estan permitidas"}

     elseif ((Get-ItemProperty -Path "HKLM:\System\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp").UserAuthentication -eq 1)
             {"Solo se permiten conexciones RDP seguras"} 

     else     {"Todas las conexiones RDP son permitidas"}

    } 
    
    Get-RemoteDesktopConfig
Write-Host ""

Write-Host CONFIGURACIONES DE TERMINAL SERVICE
Write-Host ====================================
Write-Host ""
    Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server"
Write-Host ""
    Get-ItemProperty "HKCU:\Software\Microsoft\Terminal Server Client\"
Write-Host ""

Write-Host NIVEL DE CIFRADO TERMINAL SERVICE
Write-Host ==================================
Write-Host ""
Write-Host "El nivel minimo de cifrado esta establecido en: " -NoNewline
    (Get-ItemProperty "HKLM:\System\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp\").MinEncryptionLevel
Write-Output "
1 = low
2 = client compatible
3 = high
4 = fips"
Write-Host ""

Write-Host DURACION DE LA SESION TERMINAL SERVICE      
Write-Host ====================================== 
Write-Host "Note: El campo MaxIdleTime debe estar establecido en: 30 Minutos(1800000 Milisegundos) y MaxDisconnectionTime debe estar establecido en  60 Minutos (3600000 Milisegundos)."
Write-Host ""
    Get-ItemProperty "HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services" | Format-List MaxDisconnectionTime, MaxIdleTime | Out-Host
Write-Host ""
Write-Host "Note: Si no se muestra algun resultado es porque los camposs inidcados no estan configurados"
Write-Host ""


sleep 5


Write-Host ============================================================================ -ForegroundColor green
Write-Host "                         AUDITORIA / LOGGING / MONITOREO                  " -ForegroundColor green
Write-Host ============================================================================ -ForegroundColor green
Write-Host ""

#Funcion para validar si las rutas existen
function Get-Exists{
    param(
        [string]$ruta
    )

    $existe = Test-Path -Path $ruta
    if($existe -eq $true){return 1}
    else {Write-Host "La ruta $ruta no existe" -ForegroundColor Red -BackgroundColor Black}
}

Write-Host =========================
Write-Host TAMAÑO DEL LOG DE EVENTOS
Write-Host =========================
Write-Host ""

Write-Host APPLICATION
Write-Host ===========
    $existe = Get-Exists -ruta "HKLM:\SYSTEM\CurrentControlSet\Services\Eventlog\Application"
    if($existe -eq 1)
    {
    Write-Host Maximum-Size in Bytes: -NoNewline
    (Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Services\Eventlog\Application").MaxSize
    }
    else
    {
        Write-Host $existe
    }
Write-Host ""

Write-Host SYSTEM
Write-Host ======
    $existe = Get-Exists -ruta "HKLM:\SYSTEM\CurrentControlSet\Services\Eventlog\Application"
    if($existe -eq 1)
    {
        Write-Host Maximum-Size in Bytes: -NoNewline
        (Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Services\Eventlog\System").MaxSize
    }
    else
    {
        Write-Host $existe
    }
Write-Host ""

Write-Host SECURITY
Write-Host ========
    $existe = Get-Exists -ruta "HKLM:\SYSTEM\CurrentControlSet\Services\Eventlog\Security"
    if($existe -eq 1)
    {
        Write-Host Maximum-Size in Bytes: -NoNewline
        (Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Services\Eventlog\Security").MaxSize
    }
    else
    {
        Write-Host $existe
    }
Write-Host ""
 
Write-Host ================================ 
Write-Host PERMISOS SOBRE LOGS DE EVENTOS
Write-Host ================================
Write-Host "Note: Restrict Guest Access value in registry should be set to 1."
Write-Host ""

Write-Host APPLICATION
Write-Host ===========
    Write-Host "Restrict Guest Access Value in Registry is set to " -NoNewline
    (Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Services\Eventlog\Application").RestrictGuestAccess
Write-Host ""

Write-Host SYSTEM
Write-Host ======
    Write-Host "Restrict Guest Access Value in Registry is set to " -NoNewline
    (Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Services\Eventlog\System").RestrictGuestAccess
Write-Host ""

Write-Host SECURITY
Write-Host ========
    Write-Host "Restrict Guest Access Value in Registry is set to " -NoNewline
    (Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Services\Eventlog\Security").RestrictGuestAccess
Write-Host ""

Write-Host ======================================
Write-Host ACLS SOBRE ARCHIVOS DE LOGS DE EVENTOS
Write-Host ======================================
Write-Host ""

Write-Host "APPLICATION"
Write-Host ================================================================================
    cacls "C:\WINDOWS\system32\winevt\Logs\Application.evtx"
Write-Host ""

Write-Host "SYSTEM"
Write-Host ===========================================================================
    cacls "C:\WINDOWS\system32\winevt\Logs\System.evtx"
Write-Host ""

Write-Host "SECURITY"
Write-Host =============================================================================
    cacls "C:\WINDOWS\system32\winevt\Logs\Security.evtx"
Write-Host ""

Write-Host ==========================
Write-Host ACLS DE CARPETAS SENSIBLES
Write-Host ==========================
Write-Host ""

Write-Host SYSTEM ROOT
Write-Host ====================
    Get-Acl "$env:SystemRoot" |Format-List | Out-Host
Write-Host SYSTEM32
Write-Host ========================
    Get-Acl "$env:SystemRoot\system32" |Format-List | Out-Host
Write-Host DRIVERS
Write-Host =======================
    Get-Acl "$env:SystemRoot\system32\drivers" |Format-List | Out-Host
Write-Host CONFIG
Write-Host ======================
    Get-Acl "$env:SystemRoot\System32\config" |Format-List | Out-Host
Write-Host SPOOL
Write-Host =====================
    Get-Acl "$env:SystemRoot\System32\spool" |Format-List | Out-Host
Write-Host ""

Write-Host =================
Write-Host ACLS DEL REGISTRO
Write-Host =================
Write-Host ""

Write-Host SYSTEM KEY
Write-Host ===============================
    $existe = Get-Exists -ruta "HKLM:\SYSTEM"
    if($existe -eq 1)
    {
        Get-Acl "HKLM:\SYSTEM" |Format-List | Out-Host
    }
    else
    {
        Write-Host $existe
    }
Write-Host PERFLIB KEY
Write-Host ================================
    $existe = Get-Exists -ruta "HKLM:\Software\Microsoft\Windows NT\CurrentVersion\Perflib"
    if($existe -eq 1)
    {
        Get-Acl "HKLM:\Software\Microsoft\Windows NT\CurrentVersion\Perflib" |Format-List | Out-Host
    }
    else
    {
        Write-Host $existe
    }
Write-Host WINLOGON KEY
Write-Host =================================
    $existe = Get-Exists -ruta "HKLM:\Software\Microsoft\Windows NT\CurrentVersion\Winlogon"
    if($existe -eq 1)
    {
        Get-Acl "HKLM:\Software\Microsoft\Windows NT\CurrentVersion\Winlogon" |Format-List | Out-Host
    }
    else
    {
        Write-Host $existe
    }
Write-Host LSA KEY
Write-Host ============================
    $existe = Get-Exists -ruta "HKLM:\SYSTEM\CurrentControlSet\Control\Lsa"
    if($existe -eq 1)
    {
        Get-Acl "HKLM:\SYSTEM\CurrentControlSet\Control\Lsa" |Format-List | Out-Host   
    }
    else
    {
        Write-Host $existe
    }
Write-Host SECURE PIPE SERVERS KEY
Write-Host ============================================
    $existe = Get-Exists -ruta "HKLM:\SYSTEM\CurrentControlSet\Control\SecurePipeServers"
    if($existe -eq 1)
    {
        Get-Acl "HKLM:\SYSTEM\CurrentControlSet\Control\SecurePipeServers" |Format-List | Out-Host   
    }
    else
    {
        Write-Host $existe
    }
Write-Host KNOWNDLLS KEY
Write-Host ==================================
    $existe = Get-Exists -ruta "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\KnownDLLs"
    if($existe -eq 1)
    {
        Get-Acl "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\KnownDLLs" |Format-List | Out-Host  
    }
    else
    {
        Write-Host $existe
    }
Write-Host ALLOWEDPATHS KEY
Write-Host =====================================
    $existe = Get-Exists -ruta "HKLM:\SYSTEM\CurrentControlSet\Control\SecurePipeServers\winreg\AllowedPaths"
    if($existe -eq 1)
    {
        Get-Acl "HKLM:\SYSTEM\CurrentControlSet\Control\SecurePipeServers\winreg\AllowedPaths" |Format-List | Out-Host
    }
    else
    {
        Write-Host $existe
    }
Write-Host SHARES KEY
Write-Host ===============================
    $existe = Get-Exists -ruta "HKLM:\SYSTEM\CurrentControlSet\Services\LanManServer\Shares"
    if($existe -eq 1)
    {
        Get-Acl "HKLM:\SYSTEM\CurrentControlSet\Services\LanManServer\Shares" |Format-List | Out-Host
    }
    else
    {
        Write-Host $existe
    }

Write-Host CURRENT VERSION KEYS
Write-Host =========================================
    Get-Acl "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run" |Format-List | Out-Host
    Get-Acl "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce" |Format-List | Out-Host
    Get-Acl "HKLM:\SOFTWARE\MICROSOFT\Windows NT\CurrentVersion" |Format-List | Out-Host
    Get-Acl "HKLM:\SOFTWARE\MICROSOFT\Windows NT\CurrentVersion\AeDebug" |Format-List | Out-Host
    Get-Acl "HKLM:\SOFTWARE\MICROSOFT\Windows NT\CurrentVersion\Fonts" |Format-List | Out-Host
    Get-Acl "HKLM:\SOFTWARE\MICROSOFT\Windows NT\CurrentVersion\FontSubstitutes" |Format-List | Out-Host
    Get-Acl "HKLM:\SOFTWARE\MICROSOFT\Windows NT\CurrentVersion\Font Drivers" |Format-List | Out-Host
    Get-Acl "HKLM:\SOFTWARE\MICROSOFT\Windows NT\CurrentVersion\FontMapper" |Format-List | Out-Host
    Get-Acl "HKLM:\SOFTWARE\MICROSOFT\Windows NT\CurrentVersion\GRE_Initialize" |Format-List | Out-Host
    Get-Acl "HKLM:\SOFTWARE\MICROSOFT\Windows NT\CurrentVersion\MCI Extensions" |Format-List | Out-Host
    Get-Acl "HKLM:\SOFTWARE\MICROSOFT\Windows NT\CurrentVersion\Ports" |Format-List | Out-Host
    Get-Acl "HKLM:\SOFTWARE\MICROSOFT\Windows NT\CurrentVersion\ProfileList" |Format-List | Out-Host
    Get-Acl "HKLM:\SOFTWARE\MICROSOFT\Windows NT\CurrentVersion\Compatibility32" |Format-List | Out-Host
    Get-Acl "HKLM:\SOFTWARE\MICROSOFT\Windows NT\CurrentVersion\Drivers32" |Format-List | Out-Host
    Get-Acl "HKLM:\SOFTWARE\MICROSOFT\Windows NT\CurrentVersion\MCI32" |Format-List | Out-Host

Write-Host =====================================
Write-Host ACLS DE ARCHIVOS EJECUTALES SENSIBLES
Write-Host =====================================
Write-Host "arp.exe"
    $existe = Get-Exists -ruta "C:\Windows\system32\arp.exe"
    if($existe -eq 1)
    {
        cacls "C:\Windows\system32\arp.exe"
    }
    else
    {
        Write-Host $existe
    }
Write-Host "at.exe"
    $existe = Get-Exists -ruta "C:\Windows\system32\at.exe"
    if($existe -eq 1)
    {
        cacls "C:\Windows\system32\at.exe"
    }
    else
    {
        Write-Host $existe
    }
Write-Host "attrib.exe"
    $existe = Get-Exists -ruta "C:\Windows\system32\attrib.exe"
    if($existe -eq 1)
    {
        cacls "C:\Windows\system32\attrib.exe"
    }
    else
    {
        Write-Host $existe
    }
Write-Host "cacls.exe"
    $existe = Get-Exists -ruta "C:\Windows\system32\cacls.exe"
    if($existe -eq 1)
    {
        cacls "C:\Windows\system32\cacls.exe"
    }
    else
    {
        Write-Host $existe
    }
Write-Host "cmd.exe"
    $existe = Get-Exists -ruta "C:\Windows\system32\cmd.exe"
    if($existe -eq 1)
    {
        cacls "C:\Windows\system32\cmd.exe"
    }
    else
    {
        Write-Host $existe
    }
Write-Host "dcpromo.exe"
    $existe = Get-Exists -ruta "C:\Windows\system32\dcpromo.exe"
    if($existe -eq 1)
    {
        cacls "C:\Windows\system32\dcpromo.exe"
    }
    else
    {
        Write-Host $existe
    }
Write-Host "eventcreate.exe"
    $existe = Get-Exists -ruta "C:\Windows\system32\eventcreate.exe"
    if($existe -eq 1)
    {
        cacls "C:\Windows\system32\eventcreate.exe"
    }
    else
    {
        Write-Host $existe
    }
Write-Host "finger.exe"
    $existe = Get-Exists -ruta "C:\Windows\system32\finger.exe"
    if($existe -eq 1)
    {
        cacls "C:\Windows\system32\finger.exe"
    }
    else
    {
        Write-Host $existe
    }
Write-Host "ftp.exe"
    $existe = Get-Exists -ruta "C:\Windows\system32\ftp.exe"
    if($existe -eq 1)
    {
        cacls "C:\Windows\system32\ftp.exe"
    }
    else
    {
        Write-Host $existe
    }
Write-Host "gpupdate.exe"
    $existe = Get-Exists -ruta "C:\Windows\system32\gpupdate.exe"
    if($existe -eq 1)
    {
        cacls "C:\Windows\system32\gpupdate.exe"
    }
    else
    {
        Write-Host $existe
    }
Write-Host "icacls.exe"
    $existe = Get-Exists -ruta "C:\Windows\system32\icacls.exe"
    if($existe -eq 1)
    {
        cacls "C:\Windows\system32\icacls.exe"
    }
    else
    {
        Write-Host $existe
    }
Write-Host "ipconfig.exe"
    $existe = Get-Exists -ruta "C:\Windows\system32\ipconfig.exe"
    if($existe -eq 1)
    {
        cacls "C:\Windows\system32\ipconfig.exe"
    }
    else
    {
        Write-Host $existe
    }
Write-Host "nbtstat.exe"
    $existe = Get-Exists -ruta "C:\Windows\system32\nbtstat.exe"
    if($existe -eq 1)
    {
        cacls "C:\Windows\system32\nbtstat.exe"
    }
    else
    {
        Write-Host $existe
    }
Write-Host "net.exe"
    $existe = Get-Exists -ruta "C:\Windows\system32\net.exe"
    if($existe -eq 1)
    {
        cacls "C:\Windows\system32\net.exe"
    }
    else
    {
        Write-Host $existe
    }
Write-Host "net1.exe"
    $existe = Get-Exists -ruta "C:\Windows\system32\net1.exe"
    if($existe -eq 1)
    {
       cacls "C:\Windows\system32\net1.exe"
    }
    else
    {
        Write-Host $existe
    }
Write-Host "netsh.exe"
    $existe = Get-Exists -ruta "C:\Windows\system32\netsh.exe"
    if($existe -eq 1)
    {
        cacls "C:\Windows\system32\netsh.exe"
    }
    else
    {
        Write-Host $existe
    }
Write-Host "netstat.exe"
    $existe = Get-Exists -ruta "C:\Windows\system32\netstat.exe"
    if($existe -eq 1)
    {
        cacls "C:\Windows\system32\netstat.exe"
    }
    else
    {
        Write-Host $existe
    }
Write-Host "nslookup.exe"
    $existe = Get-Exists -ruta "C:\Windows\system32\nslookup.exe"
    if($existe -eq 1)
    {
        cacls "C:\Windows\system32\nslookup.exe"
    }
    else
    {
        Write-Host $existe
    }
Write-Host "ping.exe"
    $existe = Get-Exists -ruta "C:\Windows\system32\ping.exe"
    if($existe -eq 1)
    {
        cacls "C:\Windows\system32\ping.exe"
    }
    else
    {
        Write-Host $existe
    }
Write-Host "reg.exe"
    $existe = Get-Exists -ruta "C:\Windows\system32\reg.exe"
    if($existe -eq 1)
    {
        cacls "C:\Windows\system32\reg.exe"
    }
    else
    {
        Write-Host $existe
    }
Write-Host "regedt32.exe"
    $existe = Get-Exists -ruta "C:\Windows\system32\regedt32.exe"
    if($existe -eq 1)
    {
        cacls "C:\Windows\system32\regedt32.exe"
    }
    else
    {
        Write-Host $existe
    }
Write-Host "regini.exe"
    $existe = Get-Exists -ruta "C:\Windows\system32\regini.exe"
    if($existe -eq 1)
    {
        cacls "C:\Windows\system32\regini.exe"
    }
    else
    {
        Write-Host $existe
    }
Write-Host "regsvr32.exe"
    $existe = Get-Exists -ruta "C:\Windows\system32\regsvr32.exe"
    if($existe -eq 1)
    {
        cacls "C:\Windows\system32\regsvr32.exe"
    }
    else
    {
        Write-Host $existe
    }
Write-Host "route.exe"
    $existe = Get-Exists -ruta "C:\Windows\system32\route.exe"
    if($existe -eq 1)
    {
        cacls "C:\Windows\system32\route.exe"
    }
    else
    {
        Write-Host $existe
    }
Write-Host "runonce.exe"
    $existe = Get-Exists -ruta "C:\Windows\system32\runonce.exe"
    if($existe -eq 1)
    {
        cacls "C:\Windows\system32\runonce.exe"
    }
    else
    {
        Write-Host $existe
    }
Write-Host "sc.exe"
    $existe = Get-Exists -ruta "C:\Windows\system32\sc.exe"
    if($existe -eq 1)
    {
        cacls "C:\Windows\system32\sc.exe"
    }
    else
    {
        Write-Host $existe
    }
Write-Host "secedit.exe"
    $existe = Get-Exists -ruta "C:\Windows\system32\secedit.exe"
    if($existe -eq 1)
    {
        cacls "C:\Windows\system32\secedit.exe"
    }
    else
    {
        Write-Host $existe
    }
Write-Host "subst.exe"
    $existe = Get-Exists -ruta "C:\Windows\system32\subst.exe"
    if($existe -eq 1)
    {
        cacls "C:\Windows\system32\subst.exe"
    }
    else
    {
        Write-Host $existe
    }
Write-Host "systeminfo.exe"
    $existe = Get-Exists -ruta "C:\Windows\system32\systeminfo.exe"
    if($existe -eq 1)
    {
        cacls "C:\Windows\system32\systeminfo.exe"
    }
    else
    {
        Write-Host $existe
    }
Write-Host "syskey.exe"
    $existe = Get-Exists -ruta "C:\Windows\system32\syskey.exe"
    if($existe -eq 1)
    {
        cacls "C:\Windows\system32\syskey.exe"
    }
    else
    {
        Write-Host $existe
    }
Write-Host "telnet.exe"
    $existe = Get-Exists -ruta "C:\Windows\system32\telnet.exe"
    if($existe -eq 1)
    {
        cacls "C:\Windows\system32\telnet.exe"
    }
    else
    {
        Write-Host $existe
    }
Write-Host "tftp.exe"
    $existe = Get-Exists -ruta "C:\Windows\system32\tftp.exe"
    if($existe -eq 1)
    {
        cacls "C:\Windows\system32\tftp.exe"
    }
    else
    {
        Write-Host $existe
    }
Write-Host "tlntsvr.exe"
    $existe = Get-Exists -ruta "C:\Windows\system32\tlntsvr.exe"
    if($existe -eq 1)
    {
        cacls "C:\Windows\system32\tlntsvr.exe"
    }
    else
    {
        Write-Host $existe
    }
Write-Host "tracert.exe"
    $existe = Get-Exists -ruta "C:\Windows\system32\tracert.exe"
    if($existe -eq 1)
    {
        cacls "C:\Windows\system32\tracert.exe"
    }
    else
    {
        Write-Host $existe
    }
Write-Host "xcopy.exe"
    $existe = Get-Exists -ruta "C:\Windows\system32\xcopy.exe"
    if($existe -eq 1)
    {
        cacls "C:\Windows\system32\xcopy.exe"
    }
    else
    {
        Write-Host $existe
    }
Write-Host ""

sleep 5


Write-Host ============================================================================ -ForegroundColor green
Write-Host "                         ACCESO                                           " -ForegroundColor green
Write-Host ============================================================================ -ForegroundColor green
Write-Host ""

Write-Host EVITAR EL ACCESO AL SERVIDOR POR USUARIOS QUE NO SE HALLAN LOGUEADO
Write-Host =====================================================================
Write-Host "Nota: El valor recomendado de la llave del registro debe estar establecido en 1"
Write-Host ""
Write-Host "El valor de restriccion de acceso por sesion nula esta establecido en:  " -NoNewline
    (Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Services\LanManServer\Parameters").restrictnullsessaccess
Write-Host ""

Write-Host ACCESO CON CREDENCIALES NULAS
Write-Host =============================
Write-Host "Note: The recommended value of following registry key should be set to 0." 
Write-Host ""
Write-Host "everyoneincludesanonymous value in Registry is set to " -NoNewline
    (Get-ItemProperty "HKLM:\System\CurrentControlSet\Control\Lsa").everyoneincludesanonymous
Write-Host ""
 
Write-Host EVITAR LA ENUMERACION DE USUARIOS Y CARPETAS DEL SISTEMA
Write-Host ========================================================
Write-Host "Note: The recommended value of RestrictAnonymous registry key should be set to 1"
Write-Host ""
Write-Host "RestrictAnonymous value in registry is set to " -NoNewline
    (Get-ItemProperty "HKLM:\System\CurrentControlSet\Control\Lsa\").RestrictAnonymous
Write-Host ""

Write-Host AUTENTICACION REQUERIDA DEL DOMINIO PARA EL DESBLOQUEO
Write-Host ======================================================
Write-Host ""
Write-Host "The ForceUnlockLogon value in the registry is set to " -NoNewline
    (Get-ItemProperty "HKLM:\Software\Microsoft\Windows NT\CurrentVersion\Winlogon").ForceUnlockLogon
Write-Host ""
Write-Output "
If Value:0 [hint=Domain controller authentication is not required to unlock the workstation]
If Value:1 [hint=Domain controller authentication is required to unlock the workstation]"
Write-Host ""

Write-Host =============================
Write-Host REPORTE DE ERRORES DE WINDOWS   
Write-Host =============================
Write-Host "Note: The recommended Value of following registry keys should be set to 1."
Write-Host ""
Write-Host "The value of Windows Error Reporting key in HKLM is set to " -NoNewline
    (Get-ItemProperty "HKLM:\Software\Microsoft\Windows\Windows Error Reporting\").Disabled
Write-Host ""
Write-Host "The value of Windows Error Reporting key in HKCU is set to " -NoNewline
    (Get-ItemProperty "HKCU:\Software\Microsoft\Windows\Windows Error Reporting\").Disabled
Write-Host ""

Write-Host CRASH DETECTION  
Write-Host ===============
Write-Host "Note: Gathering information about crash and shutdown events. This might take time please be patient."
Write-Host ""
Write-Host " Gathering information ..."  -ForegroundColor Yellow -BackgroundColor Black
    Start-Sleep -s 7
    $events1=get-eventlog system | where-object {$_.EventID -eq "41"} | Format-table -Wrap | Out-Host
        If ( $events1 )
	        {
	            Write-Host " Found Unexpected System Restarts. Thats Not Good. Please wait???" -ForegroundColor red -backgroundcolor black
	            Start-Sleep -s 7
	    
	            Write-Host " Here are the list of Events : " -ForegroundColor Green -BackgroundColor Black
	            Write-Output $events1
	        }
        Else
	        {
	        Write-Host " No Unexpected System Restarts Found." -ForegroundColor Green -BackgroundColor Black
	        }
Write-Host " Now Checking Normal Shutdown Events. " -ForegroundColor Yellow -BackgroundColor Black
    Start-Sleep -s 7
    $events2=get-eventlog system | where-object {$_.EventID -eq "1076"} | Format-Table -wrap | Out-Host
    $events3=get-eventlog system | where-object {$_.EventID -eq "1074"} | Format-Table -Wrap | Out-Host
    Start-Sleep 20
        If ( $events2 )
	        {
	
	        Write-Host " Found Normal Shutdown Event. Please wait???" -ForegroundColor Red -BackgroundColor Black
	        Start-Sleep -s 7
	
	        Write-Host " Here are the list of Events : " -ForegroundColor green -BackgroundColor Black
	
	        Write-Output $events2
	        }
        Else
	        { 
	        if ( $events3 )
	        {
		        Start-Sleep -s 7
		
		        Write-Output $events3
		
	        }
        else
	        {
	
	        Write-Host " No such events found from the available logs." -ForegroundColor Green -BackgroundColor Black
	
	        }
	        }
Write-Host ""
}                
            }#Cierre BEGIN--------------------------------------------------------------------------------------------------------------------------------------------------------

	    PROCESS 
		    {#INICIA PROCESS
                if($Entorno -eq "local")
                {
                    if($Reporte)
                    {
                        if( (Test-Path -Path "HKLM:\SOFTWARE\Microsoft\Office") -eq $false ){Write-Host "No existe Microsoft Office, Imposible generar el reporte" -ForegroundColor Red -BackgroundColor black}   
                        else
                        {
#-------------------------------------------------------VALIDACION CON REPORTE---------------------------------------------------------------------------------------------------

#Creacion de Documento 
$word = New-Object -ComObject word.application
$word.Visible = $True
$document = $word.Documents.Add()
$selection = $word.Selection

#Titulo
$selection.Style = "Título 1"
$selection.typetext("INFORMACION GENERAL")
$selection.TypeParaGraph()
$selection.Style = "Sin espaciado"

#-----------------------------------------------
$selection.TypeText("TIMESTAMP ")#--------------
#-----------------------------------------------
$Table = $Selection.Tables.add(
    $Selection.Range,1,1,
    [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
    [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent
)

$selection.TypeParaGraph()
$Table.Cell(1,1).range.Text = "$(Get-Date)"
$Word.Selection.Start= $Document.Content.End
$Selection.TypeParagraph()

$selection.TypeParaGraph()
#-----------------------------------------------
$selection.TypeText("SISTEMA OPERATIVO ")#------
#-----------------------------------------------
$selection.TypeParaGraph()
$OsType = @("Description","Caption","OSArchitecture","ServicePackMajorVersion")
$Range = @($Selection.Paragraphs)[-1].Range
$sOS = Get-WmiObject -class Win32_OperatingSystem | Select-Object Description, Caption, OSArchitecture, ServicePackMajorVersion
$Table = $Selection.Tables.add(
    $Selection.Range,2,($OsType.Count),
    [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
    [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent
)

#Columnas
$i = 1
foreach ($a in $OsType) {
$Table.Cell(1,$i).range.Bold = 1
$Table.Cell(1,$i).range.Text = "$a"
$i++
}
$description = $sOS.Description
$Table.Cell(2,1).range.Text = "$description"
$caption = $sOS.Caption
$Table.Cell(2,2).range.Text = "$caption"
$OSArchitecture = $sOS.OSArchitecture
$Table.Cell(2,3).range.Text = "$OSArchitecture"
$ServicePackMajorVersion = $sOS.ServicePackMajorVersion
$Table.Cell(2,4).range.Text = "$ServicePackMajorVersion"
$Word.Selection.Start= $Document.Content.End
$Selection.TypeParagraph()

#-------------------------------------------------
$selection.TypeText("INFORMACION DEL SERVIDOR ")#-
#-------------------------------------------------
$selection.TypeParaGraph()

$selection.TypeParaGraph()

#-----------------------------------------------
$selection.TypeText("INTERNET EXPLORER ")#------
#-----------------------------------------------
$selection.TypeParaGraph()
$Table = $Selection.Tables.add(
    $Selection.Range,2,2,
    [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
    [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent
)

$Table.Cell(1,1).range.Text = "Version"
$Table.Cell(1,2).range.Text = "Proxy Habilitado"
$Table.Cell(2,1).range.Text = "$((Get-ItemProperty "HKLM:\Software\Microsoft\Internet Explorer").SvcVersion)"
$Table.Cell(2,2).range.Text = " $((Get-ItemProperty "Registry::HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings").ProxyEnable)"
$Word.Selection.Start= $Document.Content.End
$Selection.TypeParagraph()

#-----------------------------------------------
$selection.TypeText("UNIDADES DE DISCO ")#------
#-----------------------------------------------
$selection.TypeParaGraph()
$Table = $Selection.Tables.add(
    $Selection.Range,4,3,
    [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
    [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent
)

$disco = [System.IO.DriveInfo]::getdrives()

$j = 2
for($i = 0; $i -lt ($disco.Count); $i++)
{
$Nombre = $disco[$i].Name
$Formato = $disco[$i].DriveFormat
$Tamano = $disco[$i].TotalSize 

$Table.Cell(1,1).range.Text = "Nombre"
$Table.Cell(1,2).range.Text = "Formato"
$Table.Cell(1,3).range.Text = "Tamaño"

$Table.Cell($j,1).range.Text = "$Nombre"
$Table.Cell($j,2).range.Text = "$Formato"
$Table.Cell($j,3).range.Text = "$Tamano"

$j++
}
$Word.Selection.Start= $Document.Content.End
$Selection.TypeParagraph()

#-----------------------------------------------
$selection.TypeText("CONFIGURACION DE RED")#------------------------<
#-----------------------------------------------
$Red = Get-WmiObject -Class Win32_NetworkAdapterConfiguration -Filter IPEnabled=TRUE
$Table = $Selection.Tables.add(
    $Selection.Range,($Red.Count + 1),5,
    [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
    [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent
)
$Table.Cell(1,1).range.Text = "Description"
$Table.Cell(1,2).range.Text = "IPAddress"
$Table.Cell(1,3).range.Text = "DHCPEnabled"
$Table.Cell(1,4).range.Text = "DNSDomain"
$Table.Cell(1,5).range.Text = "DefaultIPGateway"

$j=2
for($i=0; $i -lt ($Red.Count); $i++)
{
    $Table.Cell($j,1).range.Text = "$($Red[$i].Description)"
    $Table.Cell($j,2).range.Text = "$($Red[$i].IPAddress)"
    $Table.Cell($j,3).range.Text = "$($Red[$i].DHCPEnabled)"
    $Table.Cell($j,4).range.Text = "$($Red[$i].DNSDomain)"
    $Table.Cell($j,5).range.Text = "$($Red[$i].DefaultIPGateway)"
    $j++
}
$Word.Selection.Start= $Document.Content.End
$Selection.TypeParagraph()

#-----------------------------------------------
$selection.TypeText("NTP")#------------------------<
#-----------------------------------------------
$selection.TypeParaGraph()
$Table = $Selection.Tables.add(
    $Selection.Range,2,1,
    [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
    [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent
)
$Table.Cell(1,1).range.Text = "Name"
$Table.Cell(2,1).range.Text = "$($(Get-ItemProperty "HKLM:\SYSTEM\Currentcontrolset\Services\W32time\Parameters\").Ntpserver)" 
$Word.Selection.Start= $Document.Content.End
$Selection.TypeParagraph()

#-----------------------------------------------
$selection.TypeText("ANTIMALWARE / ANTIVIRUS")#------------------------<
#-----------------------------------------------
function Get-AntivirusName {
    [cmdletBinding()]	
    param (
    [string]$ComputerName = "$env:computername" ,
    $Credential
    )
	    BEGIN 
		    {
			    $wmiQuery = "SELECT * FROM AntiVirusProduct"
		    }

	    PROCESS 
		    {
                $AntivirusProduct = Get-WmiObject -Namespace "root\SecurityCenter2" -Query $wmiQuery  @psboundparameters # -ErrorVariable myError -ErrorAction "SilentlyContinue"	   	 
                $AntivirusProduct.displayName
		    }
	    END {
		    }
    }
    
$selection.TypeParaGraph()
$Table = $Selection.Tables.add(
    $Selection.Range,2,1,
    [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
    [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent
)
$Table.Cell(1,1).range.Text = "Name"
$Table.Cell(2,1).range.Text = "$(Get-AntivirusName)"
$Word.Selection.Start= $Document.Content.End
$Selection.TypeParagraph()

#-----------------------------------------------
$selection.TypeText("ULTIMOS PARCHES DE SEGURIDAD")#------------------------<
#-----------------------------------------------
$Selection.TypeParagraph()   
$Parches =  Get-HotFix -Description "Security*"
   $Table = $Selection.Tables.add(
    $Selection.Range,($parches.count +1 ),4,
    [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
    [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent
) 
$Table.Cell(1,1).range.Text = "Description"
$Table.Cell(1,2).range.Text = "HotFixID"
$Table.Cell(1,3).range.Text = "InstalledBy"
$Table.Cell(1,4).range.Text = "InstalledOn"

$j = 2
for($i = 0; $i -lt ($parches.Count); $i++ )
{
$Table.Cell($j,1).range.Text = "$($Parches[$i].Description)"
$Table.Cell($j,2).range.Text = "$($Parches[$i].HotFixID)"
$Table.Cell($j,3).range.Text = "$($Parches[$i].InstalledBy)"
$Table.Cell($j,4).range.Text = "$($Parches[$i].InstalledOn)"
$j++
}
$Word.Selection.Start= $Document.Content.End

$Selection.TypeParagraph()

#-----------------------------------------------
$selection.TypeText("FIREWALL")#------------------------<
#-----------------------------------------------
$Selection.TypeParagraph() 
$domain = netsh advfirewall show domainprofile 
$selection.TypeText("$domain")
$Selection.TypeParagraph() 
$Selection.TypeParagraph() 
$private = netsh advfirewall show privateprofile
$selection.TypeText("$private")
$Selection.TypeParagraph() 
$Selection.TypeParagraph() 
$public = netsh advfirewall show publicprofile
$selection.TypeText("$public")
$Selection.TypeParagraph() 
$Selection.TypeParagraph() 

#-----------------------------------------------
$selection.TypeText("OPCIONES DE SEGURIDAD")#------------------------<
#-----------------------------------------------
$security = Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters" | select NullSessionPipes, 
autodisconnect, enableforcedlogoff, enablesecuritysignature, requiresecuritysignature, restrictnullsessaccess, AdjustedNullSessionPipes, EnableAuthenticateUserSharing
$Table = $Selection.Tables.add(
    $Selection.Range,8,2,
    [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
    [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent
) 
$Table.Cell(1,1).range.Text = "NullSessionPipes"
$Table.Cell(2,1).range.Text = "autodisconnect"
$Table.Cell(3,1).range.Text = "enableforcedlogoff"
$Table.Cell(4,1).range.Text = "enablesecuritysignature"
$Table.Cell(5,1).range.Text = "requiresecuritysignature"
$Table.Cell(6,1).range.Text = "restrictnullsessaccess"
$Table.Cell(7,1).range.Text = "AdjustedNullSessionPipes"
$Table.Cell(8,1).range.Text = "EnableAuthenticateUserSharing"
$Table.Cell(1,2).range.Text = "$($security.NullSessionPipes)"
$Table.Cell(2,2).range.Text = "$($security.autodisconnect)"
$Table.Cell(3,2).range.Text = "$($security.enableforcedlogoff)"
$Table.Cell(4,2).range.Text = "$($security.enablesecuritysignature)"
$Table.Cell(5,2).range.Text = "$($security.requiresecuritysignature)"
$Table.Cell(6,2).range.Text = "$($security.restrictnullsessaccess)"
$Table.Cell(7,2).range.Text = "$($security.AdjustedNullSessionPipes)"
$Table.Cell(8,2).range.Text = "$($security.EnableAuthenticateUserSharing)"
$Word.Selection.Start= $Document.Content.End
$Selection.TypeParagraph()

#-----------------------------------------------
$selection.TypeText("GPO")#------------------------<
#-----------------------------------------------
#$gpo = Get-Module -All 
$Selection.TypeParagraph()

$Selection.TypeParagraph()


#-----------------------------------------------
$selection.TypeText("SECURE BOOT")#------------------------<
#-----------------------------------------------
$Selection.TypeParagraph()
#$SecureBoot = (Get-ItemProperty "HKLM:\System\CurrentControlSet\Control\SecureBoot\State").UEFISecureBootEnabled
#If ($SecureBoot -eq "1")
#        {
#            Write-Host "UEFI Secure Boot esta habilitado"
#        }
#    If ($SecureBoot -eq "0")
#        {
#            Write-Host "UEFI Secure Boot esta deshabilitado"
#        }

$Selection.TypeParagraph()

#-----------------------------------------------
$selection.TypeText("LEGAL")#------------------------<
#-----------------------------------------------
$Legal = Get-ItemProperty "HKLM:/SOFTWARE/Microsoft/Windows NT/CurrentVersion/Winlogon" | Select-Object LegalNoticeCaption, LegalNoticeText

$Table = $Selection.Tables.add(
    $Selection.Range,2,2,
    [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
    [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent
) 
$Table.Cell(1,1).range.Text = "LegalNoticeCaption"
$Table.Cell(1,2).range.Text = "LegalNoticeText"
$Table.Cell(2,1).range.Text = "$($Legal.LegalNoticeCaption)"
$Table.Cell(2,2).range.Text = "$($LegalNoticeText)"
$Word.Selection.Start= $Document.Content.End
$Selection.TypeParagraph()

$selection.Style = "Título 1"
$selection.typetext("USUARIOS")
$selection.TypeParaGraph()
$selection.Style = "Sin espaciado"
$selection.TypeParaGraph()

#-----------------------------------------------
$selection.TypeText("INFORMACION DE CUENTAS LOCALES")#------------------------<
#-----------------------------------------------
$account = Get-WmiObject -Class Win32_UserAccount -Filter  "LocalAccount="True"" | 
Select-Object PSComputerName, Status, Name, Domain, PasswordExpires, Description, Disabled, InstallDate, Lockout, PasswordChangeable, PasswordRequired

$Table = $Selection.Tables.add(
    $Selection.Range,11,($account.count + 1),
    [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
    [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent
) 
$Table.Cell(1,1).range.Text = "Name"
$Table.Cell(2,1).range.Text = "Domain"
$Table.Cell(3,1).range.Text = "Status"
$Table.Cell(4,1).range.Text = "Description"
$Table.Cell(5,1).range.Text = "Disabled"
$Table.Cell(6,1).range.Text = "InstallDate"
$Table.Cell(7,1).range.Text = "Lockout"
$Table.Cell(8,1).range.Text = "PSComputerName"
$Table.Cell(9,1).range.Text = "PasswordExpires"
$Table.Cell(10,1).range.Text = "PasswordChangeable"
$Table.Cell(11,1).range.Text = "PasswordRequired"

$j=2
for($i=0;$i -lt ($account.count); $i++)
{
    $Table.Cell(1,$j).range.Text = "$($account[$i].Name)"
    $Table.Cell(2,$j).range.Text = "$($account[$i].Domain)"
    $Table.Cell(3,$j).range.Text = "$($account[$i].Status)"
    $Table.Cell(4,$j).range.Text = "$($account[$i].Description)"
    $Table.Cell(5,$j).range.Text = "$($account[$i].Disabled)"
    $Table.Cell(6,$j).range.Text = "$($account[$i].InstallDate)"
    $Table.Cell(7,$j).range.Text = "$($account[$i].Lockout)"
    $Table.Cell(8,$j).range.Text = "$($account[$i].PSComputerName)"
    $Table.Cell(9,$j).range.Text = "$($account[$i].PasswordExpires)"
    $Table.Cell(10,$j).range.Text = "$($account[$i].PasswordChangeable)"
    $Table.Cell(11,$j).range.Text = "$($account[$i].PasswordRequired)"
    $j++
}
$Word.Selection.Start= $Document.Content.End
$Selection.TypeParagraph()

#-----------------------------------------------
$selection.TypeText("USUARIOS DE CADA GRUPO")#------------------------<
#-----------------------------------------------
$server = "$env:COMPUTERNAME"
$computer = [ADSI]"WinNT://$server,computer"

$Names = $computer.psbase.children | where { $_.psbase.schemaClassName -eq "group" } | foreach {$_.name}
    
 $Table = $Selection.Tables.add(
    $Selection.Range,($Names.count),2,
    [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
    [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent
)

for($i = 0;$i -lt ($names.count); $i++){
$Table.Cell(($i + 1),1).range.Text = "$($Names[$i])"
}

#-----------------------------------------------------
Function Get-LocalGroupMembership {
 
 [Cmdletbinding()]

 PARAM (
        [alias("DnsHostName","__SERVER","Computer","IPAddress")]
  [Parameter(ValueFromPipelineByPropertyName=$true,ValueFromPipeline=$true)]
  [string[]]$ComputerName = $env:COMPUTERNAME,
  
  [string]$GroupName = "Administrators"

  )
    BEGIN{
    }#BEGIN BLOCK

    PROCESS{
        foreach ($Computer in $ComputerName){
            TRY{
                $Everything_is_OK = $true

                # Testing the connection
                Write-Verbose -Message "$Computer - Testing connection..."
                Test-Connection -ComputerName $Computer -Count 1 -ErrorAction Stop |Out-Null
                     
                # Get the members for the group and computer specified
                Write-Verbose -Message "$Computer - Querying..."
             $Group = [ADSI]"WinNT://$Computer/$GroupName,group"
             $Members = @($group.psbase.Invoke("Members"))
            }#TRY
            CATCH{
                $Everything_is_OK = $false
                Write-Warning -Message "Something went wrong on $Computer"
                Write-Verbose -Message "Error on $Computer"
                }#Catch
        
            IF($Everything_is_OK){
             # Format the Output
                Write-Verbose -Message "$Computer - Formatting Data"
             $members | ForEach-Object {
              $name = $_.GetType().InvokeMember("Name", "GetProperty", $null, $_, $null)
              $class = $_.GetType().InvokeMember("Class", "GetProperty", $null, $_, $null)
              $path = $_.GetType().InvokeMember("ADsPath", "GetProperty", $null, $_, $null)
  
              # Find out if this is a local or domain object
              if ($path -like "*/$Computer/*"){
               $Type = "Local"
               }
              else {$Type = "Domain"
              }

              $Details = "" | Select-Object ComputerName,Account,Class,Group,Path,Type
              $Details.ComputerName = $Computer
              $Details.Account = $name
              $Details.Class = $class
                    $Details.Group = $GroupName
              $details.Path = $path
              $details.Type = $type
  
              # Show the Output
                    $Details
             }
            }#IF(Everything_is_OK)
        }#Foreach
    }#PROCESS BLOCK

    END{Write-Verbose -Message "Script Done"}#END BLOCK
}

$i=1
foreach ($name in $names){
    $cuentas = Get-LocalGroupMembership -ComputerName $env:COMPUTERNAME -GroupName $name
    $usuarios = $cuentas | select account
    $Table.Cell($i,2).range.Text = "." + $usuarios + "."
    $i++
}
$Word.Selection.Start= $Document.Content.End
$Selection.TypeParagraph()

#------------------------------------------------------
$selection.TypeText("ULTIMO LOGUIN CUENTAS DE USUARIO")#------------------------<
#------------------------------------------------------
$LOGUIN = $([ADSI]"WinNT://$env:COMPUTERNAME").Children | where {$_.SchemaClassName -eq "user"} | Select-Object name, lastlogin
$Table = $Selection.Tables.add(
    $Selection.Range,($LOGUIN.Count + 1),2,
    [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
    [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent
) 
$Table.Cell(1,1).range.Text = "Nombre"
$Table.Cell(1,2).range.Text = "Ultimo Loguin"

$j=2
for($i=0;$i -lt ($LOGUIN.Count); $i++)
{
    $Table.Cell($j,1).range.Text = "$($LOGUIN[$i].name)"
    $Table.Cell($j,2).range.Text = "$($LOGUIN[$i].lastlogin)"
    $j++
}
$Word.Selection.Start= $Document.Content.End
$Selection.TypeParagraph()

$selection.Style = "Título 1"
$selection.typetext("PASSWORDS")
$selection.TypeParaGraph()
$selection.Style = "Sin espaciado"

#------------------------------------------------------
$selection.TypeText("POLITICA DE PASSWORDS LOCAL")#------------------------<
#------------------------------------------------------
$selection.TypeParaGraph()
$politica = net accounts
$selection.TypeText("$politica")
$selection.TypeParaGraph()

$selection.TypeParaGraph()

#------------------------------------------------------
$selection.TypeText("POLITICA DE PASSWORDS DEL DOMINIO")#------------------------<
#------------------------------------------------------
$selection.TypeParaGraph()
$politica = net accounts /domain
$selection.TypeText("$politica")
$selection.TypeParaGraph()

$selection.TypeParaGraph()


$selection.Style = "Título 1"
$selection.typetext("RDP")
$selection.TypeParaGraph()
$selection.Style = "Sin espaciado"

#------------------------------------------------------
$selection.TypeText("CONEXIONES")#------------------------<
#------------------------------------------------------

Function Get-RemoteDesktopConfig
    {if ((Get-ItemProperty -Path "HKLM:\System\CurrentControlSet\Control\Terminal Server").fDenyTSConnections -eq 1)

              {"Las conexiones RDP no estan permitidas"}

     elseif ((Get-ItemProperty -Path "HKLM:\System\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp").UserAuthentication -eq 1)
             {"Solo se permiten conexciones RDP seguras"} 

     else     {"Todas las conexiones RDP son permitidas"}

    } 
    
$selection.TypeParaGraph()
$politica = Get-RemoteDesktopConfig
$selection.TypeText("$politica")
$selection.TypeParaGraph()

$selection.TypeParaGraph()


#------------------------------------------------------
$selection.TypeText("CONFIGURACIONES DE TERMINAL SERVICES")#------------------------<
#------------------------------------------------------
$TS = Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server" | 
select NotificationTimeOut, SnapshotMonitors, AllowRemoteRPC, DelayConMgrTimeout,fDenyTSConnections, DeleteTempDirsOnExit, fSingleSessionPerUser, PerSessionTempDir,
TSUserEnabled, fCredentialLessLogonSupported, fCredentialLessLogonSupportedTSS, fCredentialLessLogonSupportedKMRDP

$Table = $Selection.Tables.add(
    $Selection.Range,11,2,
    [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
    [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent
) 

$Table.Cell(1,1).range.Text = "NotificationTimeOut"
$Table.Cell(2,1).range.Text = "SnapshotMonitors"
$Table.Cell(3,1).range.Text = "AllowRemoteRPC"
$Table.Cell(4,1).range.Text = "DelayConMgrTimeout"
$Table.Cell(5,1).range.Text = "fDenyTSConnections"
$Table.Cell(6,1).range.Text = "DeleteTempDirsOnExit"
$Table.Cell(7,1).range.Text = "fSingleSessionPerUser"
$Table.Cell(8,1).range.Text = "PerSessionTempDir"
$Table.Cell(9,1).range.Text = "TSUserEnabled"
$Table.Cell(10,1).range.Text = "fCredentialLessLogonSupported"
$Table.Cell(11,1).range.Text = "fCredentialLessLogonSupportedTSS"
$Table.Cell(12,1).range.Text = "fCredentialLessLogonSupportedKMRDP"

$Table.Cell(1,2).range.Text = "$($TS.NotificationTimeOut)"
$Table.Cell(2,2).range.Text = "$($TS.SnapshotMonitors)"
$Table.Cell(3,2).range.Text = "$($TS.AllowRemoteRPC)"
$Table.Cell(4,2).range.Text = "$($TS.DelayConMgrTimeout)"
$Table.Cell(5,2).range.Text = "$($TS.fDenyTSConnections)"
$Table.Cell(6,2).range.Text = "$($TS.DeleteTempDirsOnExit)"
$Table.Cell(7,2).range.Text = "$($TS.fSingleSessionPerUser)"
$Table.Cell(8,2).range.Text = "$($TS.PerSessionTempDir)"
$Table.Cell(9,2).range.Text = "$($TS.TSUserEnabled)"
$Table.Cell(10,2).range.Text = "$($TS.fCredentialLessLogonSupported)"
$Table.Cell(11,2).range.Text = "$($TS.fCredentialLessLogonSupportedTSS)"
$Table.Cell(12,2).range.Text = "$($TS.fCredentialLessLogonSupportedKMRDP)"
$Word.Selection.Start= $Document.Content.End
$Selection.TypeParagraph()

$TC = Get-ItemProperty "HKCU:\Software\Microsoft\Terminal Server Client\"

if ($TC -eq $null)
{
    $selection.TypeText("Server Client no configurado")
}
else
{
    $TC = Get-ItemProperty "HKCU:\Software\Microsoft\Terminal Server Client\" | 
    select alternate shell, authentication level, autoReconnection Enabled , bitmapCacheSize, bitmapPersistCacheLocation, compression, connect to console ,
    disable cursor setting , EnableCredSspSupport, keyboardhook, redirectclipboard, redirectdrives, redirectcomports, redirectprinters, redirectsmartcards, shell working directory
}
$Selection.TypeParagraph()

$Selection.TypeParagraph()

#------------------------------------------------------
$selection.TypeText("NIVEL DE CIFRADO TERMINAL SERVICE")#------------------------<
#------------------------------------------------------
$cifrado = (Get-ItemProperty "HKLM:\System\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp\").MinEncryptionLevel
$Table = $Selection.Tables.add(
    $Selection.Range,2,1,
    [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
    [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent
) 
$Table.Cell(1,1).range.Text = "Nivel de Cifrado"

switch($cifrado){

    1 {$Table.Cell(2,1).range.Text = "Low"; break}
    2 {$Table.Cell(2,1).range.Text = "Client Compatible"; break}
    3 {$Table.Cell(2,1).range.Text = "High"; break}
    4 {$Table.Cell(2,1).range.Text = "FIPS-Compliant"; break}
    default {$Table.Cell(2,1).range.Text = "Desconocido"; break}
}
$Word.Selection.Start= $Document.Content.End
$Selection.TypeParagraph()

#------------------------------------------------------
$selection.TypeText("DURACION DE LA SESION TERMINAL SERVICES")#------------------------<
#------------------------------------------------------

$Time = Get-ItemProperty "HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services" | Format-List MaxDisconnectionTime, MaxIdleTime

    $Table = $Selection.Tables.add(
    $Selection.Range,2,2,
    [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
    [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent
) 
    $Table.Cell(1,1).range.Text = "MaxDisconnectionTime"
    $Table.Cell(1,2).range.Text = " MaxIdleTime"
    $Table.Cell(2,1).range.Text = "$($Time.MaxDisconnectionTime)"
    $Table.Cell(2,2).range.Text = "$($Time.MaxIdleTime)"
$Word.Selection.Start= $Document.Content.End
$Selection.TypeParagraph()

$selection.Style = "Título 1"
$selection.typetext("AUDITORIA")
$selection.TypeParaGraph()
$selection.Style = "Sin espaciado"

#------------------------------------------------------
$selection.TypeText("TAMAÑO Y PERMISOS DE LOG DE EVENTOS")#------------------------<
#------------------------------------------------------
$selection.TypeParaGraph()
$selection.typetext("APPLICATION")
$Table = $Selection.Tables.add(
    $Selection.Range,2,2,
    [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
    [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent
) 
$Table.Cell(1,1).range.Text = "Tamaño"
$Table.Cell(1,2).range.Text = "Permisos"
$Table.Cell(2,1).range.Text = "$((Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Services\Eventlog\Application").MaxSize)"
$Table.Cell(2,2).range.Text = "$((Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Services\Eventlog\Application").RestrictGuestAccess)"
$Word.Selection.Start= $Document.Content.End

$selection.TypeParaGraph()

$selection.typetext("SYSTEM")
$Table = $Selection.Tables.add(
    $Selection.Range,2,2,
    [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
    [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent
) 
$Table.Cell(1,1).range.Text = "Tamaño"
$Table.Cell(1,2).range.Text = "Permisos"
$Table.Cell(2,1).range.Text = "$((Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Services\Eventlog\System").MaxSize)"
$Table.Cell(2,2).range.Text = "$((Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Services\Eventlog\System").RestrictGuestAccess)"
$Word.Selection.Start= $Document.Content.End

$selection.TypeParaGraph()

$selection.typetext("SECURITY")
$Table = $Selection.Tables.add(
    $Selection.Range,2,2,
    [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
    [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent
) 
$Table.Cell(1,1).range.Text = "Tamaño"
$Table.Cell(1,2).range.Text = "Permisos"
$Table.Cell(2,1).range.Text = "$((Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Services\Eventlog\Security").MaxSize)"
$Table.Cell(2,2).range.Text = "$((Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Services\Eventlog\Security").RestrictGuestAccess)"
$Word.Selection.Start= $Document.Content.End
$selection.TypeParaGraph()

#------------------------------------------------------
$selection.TypeText("TAMAÑO Y PERMISOS DE LOG DE EVENTOS")#------------------------<
#------------------------------------------------------
$selection.TypeParaGraph()

$usuario = ((Get-Item C:\WINDOWS\system32\winevt\Logs\Application.evtx).getaccesscontrol("Access")).Access | select IdentityReference
$acceso = ((Get-Item C:\WINDOWS\system32\winevt\Logs\Application.evtx).getaccesscontrol("Access")).Access | select FileSystemRights
$Table = $Selection.Tables.add(
    $Selection.Range,($usuario.count + 1),2,
    [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
    [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent
)
$Table.Cell(1,1).range.Text = "ENTITY"
$Table.Cell(1,2).range.Text = "NIVEL DE ACCESO"

$j=2
for($i=0;$i -lt ($usuario.count);$i++)
{
    $Table.Cell($j,1).range.Text = "$($usuario[$i])"
    $Table.Cell($j,2).range.Text = "$($acceso[$i])"
    $j++
}
$Word.Selection.Start= $Document.Content.End
$selection.TypeParaGraph()

#------------------------------------------------------
$selection.TypeText("ACLS DE CARPETAS SENSIBLES")#------------------------<
#------------------------------------------------------
$selection.TypeParaGraph()
#------------------------------------------------------
$selection.TypeText("SystemRoot")#------------------------<
#------------------------------------------------------
$usuario = (Get-Acl "$env:SystemRoot").Access | select {$_.IdentityReference}
$acceso = (Get-Acl "$env:SystemRoot").Access | select {$_.FileSystemRights}
$Table = $Selection.Tables.add(
    $Selection.Range,($usuario.count + 1),2,
    [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
    [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent
) 
$Table.Cell(1,1).range.Text = "ENTITY"
$Table.Cell(1,2).range.Text = "NIVEL DE ACCESO"

$j=2
for($i=0;$i -lt ($usuario.count);$i++)
{
    $Table.Cell($j,1).range.Text = "$($usuario[$i])"
    $Table.Cell($j,2).range.Text = "$($acceso[$i])"
    $j++
}
$Word.Selection.Start= $Document.Content.End
$selection.TypeParaGraph()

#------------------------------------------------------
$selection.TypeText("system32")#------------------------<
#------------------------------------------------------
$usuario = (Get-Acl "$env:SystemRoot\system32").Access | select {$_.IdentityReference}
$acceso = (Get-Acl "$env:SystemRoot\system32").Access | select {$_.FileSystemRights}
$Table = $Selection.Tables.add(
    $Selection.Range,($usuario.count + 1),2,
    [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
    [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent
) 
$Table.Cell(1,1).range.Text = "ENTITY"
$Table.Cell(1,2).range.Text = "NIVEL DE ACCESO"

$j=2
for($i=0;$i -lt ($usuario.count);$i++)
{
    $Table.Cell($j,1).range.Text = "$($usuario[$i])"
    $Table.Cell($j,2).range.Text = "$($acceso[$i])"
    $j++
}
$Word.Selection.Start= $Document.Content.End
$selection.TypeParaGraph()

#------------------------------------------------------
$selection.TypeText("drivers")#------------------------<
#------------------------------------------------------
$usuario = (Get-Acl "$env:SystemRoot\system32\drivers").Access | select {$_.IdentityReference}
$acceso = (Get-Acl "$env:SystemRoot\system32\drivers").Access | select {$_.FileSystemRights}
$Table = $Selection.Tables.add(
    $Selection.Range,($usuario.count + 1),2,
    [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
    [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent
) 
$Table.Cell(1,1).range.Text = "ENTITY"
$Table.Cell(1,2).range.Text = "NIVEL DE ACCESO"

$j=2
for($i=0;$i -lt ($usuario.count);$i++)
{
    $Table.Cell($j,1).range.Text = "$($usuario[$i])"
    $Table.Cell($j,2).range.Text = "$($acceso[$i])"
    $j++
}
$Word.Selection.Start= $Document.Content.End
$selection.TypeParaGraph()

#------------------------------------------------------
$selection.TypeText("config")#------------------------<
#------------------------------------------------------
$usuario = (Get-Acl "$env:SystemRoot\System32\config").Access | select {$_.IdentityReference}
$acceso = (Get-Acl "$env:SystemRoot\System32\config").Access | select {$_.FileSystemRights}
$Table = $Selection.Tables.add(
    $Selection.Range,($usuario.count + 1),2,
    [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
    [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent
) 
$Table.Cell(1,1).range.Text = "ENTITY"
$Table.Cell(1,2).range.Text = "NIVEL DE ACCESO"

$j=2
for($i=0;$i -lt ($usuario.count);$i++)
{
    $Table.Cell($j,1).range.Text = "$($usuario[$i])"
    $Table.Cell($j,2).range.Text = "$($acceso[$i])"
    $j++
}
$Word.Selection.Start= $Document.Content.End
$selection.TypeParaGraph()

#------------------------------------------------------
$selection.TypeText("spool")#------------------------<
#------------------------------------------------------
$usuario = (Get-Acl "$env:SystemRoot\System32\spool").Access | select IdentityReference
$acceso = (Get-Acl "$env:SystemRoot\System32\spool").Access | select FileSystemRights
$Table = $Selection.Tables.add(
    $Selection.Range,($usuario.count + 1),2,
    [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
    [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent
) 
$Table.Cell(1,1).range.Text = "ENTITY"
$Table.Cell(1,2).range.Text = "NIVEL DE ACCESO"

$j=2
for($i=0;$i -lt ($usuario.count);$i++)
{
    $Table.Cell($j,1).range.Text = "$($usuario[$i])"
    $Table.Cell($j,2).range.Text = "$($acceso[$i])"
    $j++
}
$Word.Selection.Start= $Document.Content.End
$selection.TypeParaGraph()

#------------------------------------------------------
$selection.TypeText("ACL´S DE REGISTRO")#------------------------<
#------------------------------------------------------
$selection.TypeParaGraph()

#------------------------------------------------------
$selection.TypeText("Perflib")#------------------------<
#------------------------------------------------------
$usuario = (Get-Acl "HKLM:\Software\Microsoft\Windows NT\CurrentVersion\Perflib").Access | select IdentityReference
$acceso = (Get-Acl "HKLM:\Software\Microsoft\Windows NT\CurrentVersion\Perflib").Access | select AccessControlType
$Table = $Selection.Tables.add(
    $Selection.Range,($usuario.count + 1),2,
    [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
    [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent
) 
$Table.Cell(1,1).range.Text = "ENTITY"
$Table.Cell(1,2).range.Text = "NIVEL DE ACCESO"

$j=2
for($i=0;$i -lt ($usuario.count);$i++)
{
    $Table.Cell($j,1).range.Text = "$($usuario[$i])"
    $Table.Cell($j,2).range.Text = "$($acceso[$i])"
    $j++
}
$Word.Selection.Start= $Document.Content.End
$selection.TypeParaGraph()

#------------------------------------------------------
$selection.TypeText("Winlogon")#------------------------<
#------------------------------------------------------
$usuario = (Get-Acl "HKLM:\Software\Microsoft\Windows NT\CurrentVersion\Winlogon").Access | select IdentityReference
$acceso = (Get-Acl "HKLM:\Software\Microsoft\Windows NT\CurrentVersion\Winlogon").Access | select AccessControlType
$Table = $Selection.Tables.add(
    $Selection.Range,($usuario.count + 1),2,
    [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
    [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent
) 
$Table.Cell(1,1).range.Text = "ENTITY"
$Table.Cell(1,2).range.Text = "NIVEL DE ACCESO"

$j=2
for($i=0;$i -lt ($usuario.count);$i++)
{
    $Table.Cell($j,1).range.Text = "$($usuario[$i])"
    $Table.Cell($j,2).range.Text = "$($acceso[$i])"
    $j++
}
$Word.Selection.Start= $Document.Content.End
$selection.TypeParaGraph()

#------------------------------------------------------
$selection.TypeText("Lsa")#------------------------<
#------------------------------------------------------
$usuario = (Get-Acl "HKLM:\SYSTEM\CurrentControlSet\Control\Lsa").Access | select IdentityReference
$acceso = (Get-Acl "HKLM:\SYSTEM\CurrentControlSet\Control\Lsa").Access | select AccessControlType
$Table = $Selection.Tables.add(
    $Selection.Range,($usuario.count + 1),2,
    [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
    [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent
) 
$Table.Cell(1,1).range.Text = "ENTITY"
$Table.Cell(1,2).range.Text = "NIVEL DE ACCESO"

$j=2
for($i=0;$i -lt ($usuario.count);$i++)
{
    $Table.Cell($j,1).range.Text = "$($usuario[$i])"
    $Table.Cell($j,2).range.Text = "$($acceso[$i])"
    $j++
}
$Word.Selection.Start= $Document.Content.End
$selection.TypeParaGraph()

#------------------------------------------------------
$selection.TypeText("SecurePipeServers")#------------------------<
#------------------------------------------------------
$usuario = (Get-Acl "HKLM:\SYSTEM\CurrentControlSet\Control\SecurePipeServers").Access | select IdentityReference
$acceso = (Get-Acl "HKLM:\SYSTEM\CurrentControlSet\Control\SecurePipeServers").Access | select AccessControlType
$Table = $Selection.Tables.add(
    $Selection.Range,($usuario.count + 1),2,
    [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
    [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent
) 
$Table.Cell(1,1).range.Text = "ENTITY"
$Table.Cell(1,2).range.Text = "NIVEL DE ACCESO"

$j=2
for($i=0;$i -lt ($usuario.count);$i++)
{
    $Table.Cell($j,1).range.Text = "$($usuario[$i])"
    $Table.Cell($j,2).range.Text = "$($acceso[$i])"
    $j++
}
$Word.Selection.Start= $Document.Content.End
$selection.TypeParaGraph()

#------------------------------------------------------
$selection.TypeText("KnownDLLs")#------------------------<
#------------------------------------------------------
$usuario = (Get-Acl "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\KnownDLLs").Access | select IdentityReference
$acceso = (Get-Acl "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\KnownDLLs").Access | select AccessControlType
$Table = $Selection.Tables.add(
    $Selection.Range,($usuario.count + 1),2,
    [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
    [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent
) 
$Table.Cell(1,1).range.Text = "ENTITY"
$Table.Cell(1,2).range.Text = "NIVEL DE ACCESO"

$j=2
for($i=0;$i -lt ($usuario.count);$i++)
{
    $Table.Cell($j,1).range.Text = "$($usuario[$i])"
    $Table.Cell($j,2).range.Text = "$($acceso[$i])"
    $j++
}
$Word.Selection.Start= $Document.Content.End
$selection.TypeParaGraph()

#------------------------------------------------------
$selection.TypeText("AllowedPaths")#------------------------<
#------------------------------------------------------
$usuario = (Get-Acl "HKLM:\SYSTEM\CurrentControlSet\Control\SecurePipeServers\winreg\AllowedPaths").Access | select IdentityReference
$acceso = (Get-Acl "HKLM:\SYSTEM\CurrentControlSet\Control\SecurePipeServers\winreg\AllowedPaths").Access | select AccessControlType
$Table = $Selection.Tables.add(
    $Selection.Range,($usuario.count + 1),2,
    [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
    [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent
) 
$Table.Cell(1,1).range.Text = "ENTITY"
$Table.Cell(1,2).range.Text = "NIVEL DE ACCESO"

$j=2
for($i=0;$i -lt ($usuario.count);$i++)
{
    $Table.Cell($j,1).range.Text = "$($usuario[$i])"
    $Table.Cell($j,2).range.Text = "$($acceso[$i])"
    $j++
}
$Word.Selection.Start= $Document.Content.End
$selection.TypeParaGraph()

#------------------------------------------------------
$selection.TypeText("Shares")#------------------------<
#------------------------------------------------------
$usuario = (Get-Acl "HKLM:\SYSTEM\CurrentControlSet\Services\LanManServer\Shares").Access | select IdentityReference
$acceso = (Get-Acl "HKLM:\SYSTEM\CurrentControlSet\Services\LanManServer\Shares").Access | select AccessControlType
 
$Table = $Selection.Tables.add(
    $Selection.Range,($usuario.count + 1),2,
    [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
    [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent
) 
$Table.Cell(1,1).range.Text = "ENTITY"
$Table.Cell(1,2).range.Text = "NIVEL DE ACCESO"

$j=2
for($i=0;$i -lt ($usuario.count);$i++)
{
    $Table.Cell($j,1).range.Text = "$($usuario[$i])"
    $Table.Cell($j,2).range.Text = "$($acceso[$i])"
    $j++
}
$Word.Selection.Start= $Document.Content.End
$selection.TypeParaGraph()

#------------------------------------------------------
$selection.TypeText("SNMP KEYS")#------------------------<
#------------------------------------------------------
$selection.TypeParaGraph()

#------------------------------------------------------
$selection.TypeText("ValidCommunities")#------------------------<
#------------------------------------------------------
$selection.TypeParaGraph()
$existe = Test-Path -Path "HKLM:\SYSTEM\CurrentControlSet\Services\SNMP\Parameters\ValidCommunities"
if($existe -eq $false){
    $selection.TypeText("La llave ValidCommunities no existe")
    $selection.TypeParaGraph()
}
else
{
    $usuario = (Get-Acl "HKLM:\SYSTEM\CurrentControlSet\Services\SNMP\Parameters\ValidCommunities").Access | select IdentityReference
    $acceso = (Get-Acl "HKLM:\SYSTEM\CurrentControlSet\Services\SNMP\Parameters\ValidCommunities").Access | select AccessControlType
    $Table = $Selection.Tables.add(
    $Selection.Range,($usuario.count + 1),2,
    [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
    [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent
    ) 
    $Table.Cell(1,1).range.Text = "ENTITY"
    $Table.Cell(1,2).range.Text = "NIVEL DE ACCESO"

    $j=2
    for($i=0;$i -lt ($usuario.count);$i++)
    {
        $Table.Cell($j,1).range.Text = "$($usuario[$i])"
        $Table.Cell($j,2).range.Text = "$($acceso[$i])"
        $j++
    }
    $Word.Selection.Start= $Document.Content.End
    $selection.TypeParaGraph()
}

$selection.TypeParaGraph()
#------------------------------------------------------
$selection.TypeText("PermittedManagers")#------------------------<
#------------------------------------------------------
$selection.TypeParaGraph()
$existe = Test-Path -Path "HKLM:\SYSTEM\CurrentControlSet\Services\SNMP\Parameters\PermittedManagers"
if($existe -eq $false){
    $selection.TypeText("La llave PermittedManagers no existe")
    $selection.TypeParaGraph()
}
else
{
    $usuario = (Get-Acl "HKLM:\SYSTEM\CurrentControlSet\Services\SNMP\Parameters\PermittedManagers").Access | select IdentityReference
    $acceso = (Get-Acl "HKLM:\SYSTEM\CurrentControlSet\Services\SNMP\Parameters\PermittedManagers").Access | select AccessControlType
    $Table = $Selection.Tables.add(
    $Selection.Range,($usuario.count + 1),2,
    [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
    [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent
    ) 
    $Table.Cell(1,1).range.Text = "ENTITY"
    $Table.Cell(1,2).range.Text = "NIVEL DE ACCESO"

    $j=2
    for($i=0;$i -lt ($usuario.count);$i++)
    {
        $Table.Cell($j,1).range.Text = "$($usuario[$i])"
        $Table.Cell($j,2).range.Text = "$($acceso[$i])"
        $j++
    }
    $Word.Selection.Start= $Document.Content.End
    $selection.TypeParaGraph()
}

$selection.TypeParaGraph()
#------------------------------------------------------
$selection.TypeText("ValidCommunities")#------------------------<
#------------------------------------------------------
$selection.TypeParaGraph()
$existe = Test-Path -Path "HKLM:\SOFTWARE\Policies\SNMP\Parameters\ValidCommunities"
if($existe -eq $false){
    $selection.TypeText("La llave ValidCommunities no existe")
    $selection.TypeParaGraph()
}
else
{
    $usuario = (Get-Acl "HKLM:\SOFTWARE\Policies\SNMP\Parameters\ValidCommunities").Access | select IdentityReference
    $acceso = (Get-Acl "HKLM:\SOFTWARE\Policies\SNMP\Parameters\ValidCommunities").Access | select AccessControlType
     $Table = $Selection.Tables.add(
    $Selection.Range,($usuario.count + 1),2,
    [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
    [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent
    ) 
    $Table.Cell(1,1).range.Text = "ENTITY"
    $Table.Cell(1,2).range.Text = "NIVEL DE ACCESO"

    $j=2
    for($i=0;$i -lt ($usuario.count);$i++)
    {
        $Table.Cell($j,1).range.Text = "$($usuario[$i])"
        $Table.Cell($j,2).range.Text = "$($acceso[$i])"
        $j++
    }
    $Word.Selection.Start= $Document.Content.End
    $selection.TypeParaGraph()
}

$selection.TypeParaGraph()
#------------------------------------------------------
$selection.TypeText("PermittedManagers")#------------------------<
#------------------------------------------------------
$selection.TypeParaGraph()
$existe = Test-Path -Path "HKLM:\SOFTWARE\Policies\SNMP\Parameters\PermittedManagers"
if($existe -eq $false){
    $selection.TypeText("La llave PermittedManagers no existe")
    $selection.TypeParaGraph()
}
else
{
    $usuario = (Get-Acl "HKLM:\SOFTWARE\Policies\SNMP\Parameters\PermittedManagers").Access | select IdentityReference
    $acceso = (Get-Acl "HKLM:\SOFTWARE\Policies\SNMP\Parameters\PermittedManagers").Access | select AccessControlType
    $Table = $Selection.Tables.add(
    $Selection.Range,($usuario.count + 1),2,
    [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
    [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent
    ) 
    $Table.Cell(1,1).range.Text = "ENTITY"
    $Table.Cell(1,2).range.Text = "NIVEL DE ACCESO"

    $j=2
    for($i=0;$i -lt ($usuario.count);$i++)
    {
        $Table.Cell($j,1).range.Text = "$($usuario[$i])"
        $Table.Cell($j,2).range.Text = "$($acceso[$i])"
        $j++
    }
    $Word.Selection.Start= $Document.Content.End
    $selection.TypeParaGraph()
}

$selection.TypeParaGraph()
#------------------------------------------------------
$selection.TypeText("KEYS CURRENTVERSION")#------------------------<
#------------------------------------------------------
$selection.TypeParaGraph()

#------------------------------------------------------
$selection.TypeText("Run")#------------------------<
#------------------------------------------------------
$existe = Test-Path -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
if($existe -eq $false){
    $selection.TypeText("La llave Run no existe")
    $selection.TypeParaGraph()
}
else
{
    $usuario = (Get-Acl "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run").Access | select IdentityReference
    $acceso = (Get-Acl "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run").Access | select AccessControlType
    $Table = $Selection.Tables.add(
    $Selection.Range,($usuario.count + 1),2,
    [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
    [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent
    ) 
    $Table.Cell(1,1).range.Text = "ENTITY"
    $Table.Cell(1,2).range.Text = "NIVEL DE ACCESO"

    $j=2
    for($i=0;$i -lt ($usuario.count);$i++)
    {
        $Table.Cell($j,1).range.Text = "$($usuario[$i])"
        $Table.Cell($j,2).range.Text = "$($acceso[$i])"
        $j++
    }
    $Word.Selection.Start= $Document.Content.End
    $selection.TypeParaGraph()
 }

$selection.TypeParaGraph()
#------------------------------------------------------
$selection.TypeText("RunOnce")#------------------------<
#------------------------------------------------------
$existe = Test-Path -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce"
if($existe -eq $false){
    $selection.TypeText("La llave RunOnce no existe")
    $selection.TypeParaGraph()
}
else
{
    $usuario = (Get-Acl "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce").Access | select IdentityReference
    $acceso = (Get-Acl "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce").Access | select AccessControlType
    $Table = $Selection.Tables.add(
    $Selection.Range,($usuario.count + 1),2,
    [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
    [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent
    ) 
    $Table.Cell(1,1).range.Text = "ENTITY"
    $Table.Cell(1,2).range.Text = "NIVEL DE ACCESO"

    $j=2
    for($i=0;$i -lt ($usuario.count);$i++)
    {
        $Table.Cell($j,1).range.Text = "$($usuario[$i])"
        $Table.Cell($j,2).range.Text = "$($acceso[$i])"
        $j++
    }
    $Word.Selection.Start= $Document.Content.End
    $selection.TypeParaGraph()
}

$selection.TypeParaGraph()
#------------------------------------------------------
$selection.TypeText("CurrentVersion")#------------------------<
#------------------------------------------------------
$existe = Test-Path -Path "HKLM:\SOFTWARE\MICROSOFT\Windows NT\CurrentVersion"
if($existe -eq $false){
    $selection.TypeText("La Ruta CurrentVersion no existe")
    $selection.TypeParaGraph()
}
else
{
    $usuario = (Get-Acl "HKLM:\SOFTWARE\MICROSOFT\Windows NT\CurrentVersion").Access | select IdentityReference
    $acceso = (Get-Acl "HKLM:\SOFTWARE\MICROSOFT\Windows NT\CurrentVersion").Access | select AccessControlType
    $Table = $Selection.Tables.add(
    $Selection.Range,($usuario.count + 1),2,
    [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
    [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent
    ) 
    $Table.Cell(1,1).range.Text = "ENTITY"
    $Table.Cell(1,2).range.Text = "NIVEL DE ACCESO"

    $j=2
    for($i=0;$i -lt ($usuario.count);$i++)
    {
        $Table.Cell($j,1).range.Text = "$($usuario[$i])"
        $Table.Cell($j,2).range.Text = "$($acceso[$i])"
        $j++
    }
    $Word.Selection.Start= $Document.Content.End
    $selection.TypeParaGraph()
}

$selection.TypeParaGraph()
#------------------------------------------------------
$selection.TypeText("AeDebug")#------------------------<
#------------------------------------------------------
$existe = Test-Path -Path "HKLM:\SOFTWARE\MICROSOFT\Windows NT\CurrentVersion\AeDebug"
if($existe -eq $false){
    $selection.TypeText("La llave AeDebug no existe")
    $selection.TypeParaGraph()
}
else
{
    $usuario = (Get-Acl "HKLM:\SOFTWARE\MICROSOFT\Windows NT\CurrentVersion\AeDebug").Access | select IdentityReference
    $acceso = (Get-Acl "HKLM:\SOFTWARE\MICROSOFT\Windows NT\CurrentVersion\AeDebug").Access | select AccessControlType
    $Table = $Selection.Tables.add(
    $Selection.Range,($usuario.count + 1),2,
    [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
    [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent
    ) 
    $Table.Cell(1,1).range.Text = "ENTITY"
    $Table.Cell(1,2).range.Text = "NIVEL DE ACCESO"

    $j=2
    for($i=0;$i -lt ($usuario.count);$i++)
    {
        $Table.Cell($j,1).range.Text = "$($usuario[$i])"
        $Table.Cell($j,2).range.Text = "$($acceso[$i])"
        $j++
    }
    $Word.Selection.Start= $Document.Content.End
    $selection.TypeParaGraph()
}

$selection.TypeParaGraph()
#------------------------------------------------------
$selection.TypeText("Fonts")#------------------------<
#------------------------------------------------------
$existe = Test-Path -Path "HKLM:\SOFTWARE\MICROSOFT\Windows NT\CurrentVersion\Fonts"
if($existe -eq $false){
    $selection.TypeText("La llave Fonts no existe")
    $selection.TypeParaGraph()
}
else
{
    $usuario = (Get-Acl "HKLM:\SOFTWARE\MICROSOFT\Windows NT\CurrentVersion\Fonts").Access | select IdentityReference
    $acceso = (Get-Acl "HKLM:\SOFTWARE\MICROSOFT\Windows NT\CurrentVersion\Fonts").Access | select AccessControlType
    $Table = $Selection.Tables.add(
    $Selection.Range,($usuario.count + 1),2,
    [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
    [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent
    ) 
    $Table.Cell(1,1).range.Text = "ENTITY"
    $Table.Cell(1,2).range.Text = "NIVEL DE ACCESO"

    $j=2
    for($i=0;$i -lt ($usuario.count);$i++)
    {
        $Table.Cell($j,1).range.Text = "$($usuario[$i])"
        $Table.Cell($j,2).range.Text = "$($acceso[$i])"
        $j++
    }
    $Word.Selection.Start= $Document.Content.End
    $selection.TypeParaGraph()
}

$selection.TypeParaGraph()
#------------------------------------------------------
$selection.TypeText("FontSubstitutes")#------------------------<
#------------------------------------------------------
$existe = Test-Path -Path "HKLM:\SOFTWARE\MICROSOFT\Windows NT\CurrentVersion\FontSubstitutes"
if($existe -eq $false){
    $selection.TypeText("La llave FontSubstitutes no existe")
    $selection.TypeParaGraph()
}
else
{
    $usuario = (Get-Acl "HKLM:\SOFTWARE\MICROSOFT\Windows NT\CurrentVersion\FontSubstitutes").Access | select IdentityReference
    $acceso = (Get-Acl "HKLM:\SOFTWARE\MICROSOFT\Windows NT\CurrentVersion\FontSubstitutes").Access | select AccessControlType
    $Table = $Selection.Tables.add(
    $Selection.Range,($usuario.count + 1),2,
    [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
    [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent
    ) 
    $Table.Cell(1,1).range.Text = "ENTITY"
    $Table.Cell(1,2).range.Text = "NIVEL DE ACCESO"

    $j=2
    for($i=0;$i -lt ($usuario.count);$i++)
    {
        $Table.Cell($j,1).range.Text = "$($usuario[$i])"
        $Table.Cell($j,2).range.Text = "$($acceso[$i])"
        $j++
    }
    $Word.Selection.Start= $Document.Content.End
    $selection.TypeParaGraph()
}

$selection.TypeParaGraph()
#------------------------------------------------------
$selection.TypeText("Font Drivers")#------------------------<
#------------------------------------------------------
$existe = Test-Path -Path "HKLM:\SOFTWARE\MICROSOFT\Windows NT\CurrentVersion\Font Drivers"
if($existe -eq $false){
    $selection.TypeText("La llave Font Drivers no existe")
    $selection.TypeParaGraph()
}
else
{
    $usuario = (Get-Acl "HKLM:\SOFTWARE\MICROSOFT\Windows NT\CurrentVersion\Font Drivers").Access | select IdentityReference
    $acceso = (Get-Acl "HKLM:\SOFTWARE\MICROSOFT\Windows NT\CurrentVersion\Font Drivers").Access | select AccessControlType
    $Table = $Selection.Tables.add(
    $Selection.Range,($usuario.count + 1),2,
    [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
    [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent
    ) 
    $Table.Cell(1,1).range.Text = "ENTITY"
    $Table.Cell(1,2).range.Text = "NIVEL DE ACCESO"

    $j=2
    for($i=0;$i -lt ($usuario.count);$i++)
    {
        $Table.Cell($j,1).range.Text = "$($usuario[$i])"
        $Table.Cell($j,2).range.Text = "$($acceso[$i])"
        $j++
    }
    $Word.Selection.Start= $Document.Content.End
    $selection.TypeParaGraph()
}

$selection.TypeParaGraph()
#------------------------------------------------------
$selection.TypeText("FontMapper")#------------------------<
#------------------------------------------------------
$existe = Test-Path -Path "HKLM:\SOFTWARE\MICROSOFT\Windows NT\CurrentVersion\FontMapper"
if($existe -eq $false){
    $selection.TypeText("La llave FontMapper no existe")
    $selection.TypeParaGraph()
}
else
{
    $usuario = (Get-Acl "HKLM:\SOFTWARE\MICROSOFT\Windows NT\CurrentVersion\FontMapper").Access | select IdentityReference
    $acceso = (Get-Acl "HKLM:\SOFTWARE\MICROSOFT\Windows NT\CurrentVersion\FontMapper").Access | select AccessControlType
    $Table = $Selection.Tables.add(
    $Selection.Range,($usuario.count + 1),2,
    [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
    [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent
    ) 
    $Table.Cell(1,1).range.Text = "ENTITY"
    $Table.Cell(1,2).range.Text = "NIVEL DE ACCESO"

    $j=2
    for($i=0;$i -lt ($usuario.count);$i++)
    {
        $Table.Cell($j,1).range.Text = "$($usuario[$i])"
        $Table.Cell($j,2).range.Text = "$($acceso[$i])"
        $j++
    }
    $Word.Selection.Start= $Document.Content.End
    $selection.TypeParaGraph()
}

$selection.TypeParaGraph()
#------------------------------------------------------
$selection.TypeText("GRE_Initialize")#------------------------<
#------------------------------------------------------
$existe = Test-Path -Path "HKLM:\SOFTWARE\MICROSOFT\Windows NT\CurrentVersion\GRE_Initialize"
if($existe -eq $false){
    $selection.TypeText("La llave GRE_Initialize no existe")
    $selection.TypeParaGraph()
}
else
{
    $usuario = (Get-Acl "HKLM:\SOFTWARE\MICROSOFT\Windows NT\CurrentVersion\GRE_Initialize").Access | select IdentityReference
    $acceso = (Get-Acl "HKLM:\SOFTWARE\MICROSOFT\Windows NT\CurrentVersion\GRE_Initialize").Access | select AccessControlType
    $Table = $Selection.Tables.add(
    $Selection.Range,($usuario.count + 1),2,
    [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
    [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent
    ) 
    $Table.Cell(1,1).range.Text = "ENTITY"
    $Table.Cell(1,2).range.Text = "NIVEL DE ACCESO"

    $j=2
    for($i=0;$i -lt ($usuario.count);$i++)
    {
        $Table.Cell($j,1).range.Text = "$($usuario[$i])"
        $Table.Cell($j,2).range.Text = "$($acceso[$i])"
        $j++
    }
    $Word.Selection.Start= $Document.Content.End
    $selection.TypeParaGraph()
}

$selection.TypeParaGraph()
#------------------------------------------------------
$selection.TypeText("MCI Extensions")#------------------------<
#------------------------------------------------------
$existe = Test-Path -Path "HKLM:\SOFTWARE\MICROSOFT\Windows NT\CurrentVersion\MCI Extensions"
if($existe -eq $false){
    $selection.TypeText("La llave MCI Extensions no existe")
    $selection.TypeParaGraph()
}
else
{
    $usuario = (Get-Acl "HKLM:\SOFTWARE\MICROSOFT\Windows NT\CurrentVersion\MCI Extensions").Access | select IdentityReference
    $acceso = (Get-Acl "HKLM:\SOFTWARE\MICROSOFT\Windows NT\CurrentVersion\MCI Extensions").Access | select AccessControlType
    $Table = $Selection.Tables.add(
    $Selection.Range,($usuario.count + 1),2,
    [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
    [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent
    ) 
    $Table.Cell(1,1).range.Text = "ENTITY"
    $Table.Cell(1,2).range.Text = "NIVEL DE ACCESO"

    $j=2
    for($i=0;$i -lt ($usuario.count);$i++)
    {
        $Table.Cell($j,1).range.Text = "$($usuario[$i])"
        $Table.Cell($j,2).range.Text = "$($acceso[$i])"
        $j++
    }
    $Word.Selection.Start= $Document.Content.End
    $selection.TypeParaGraph()
}

$selection.TypeParaGraph()
#------------------------------------------------------
$selection.TypeText("Ports")#------------------------<
#------------------------------------------------------
$existe = Test-Path -Path "HKLM:\SOFTWARE\MICROSOFT\Windows NT\CurrentVersion\Ports"
if($existe -eq $false){
    $selection.TypeText("La llave Ports no existe")
    $selection.TypeParaGraph()
}
else
{
    $usuario = (Get-Acl "HKLM:\SOFTWARE\MICROSOFT\Windows NT\CurrentVersion\Ports").Access | select IdentityReference
    $acceso = (Get-Acl "HKLM:\SOFTWARE\MICROSOFT\Windows NT\CurrentVersion\Ports").Access | select AccessControlType
    $Table = $Selection.Tables.add(
    $Selection.Range,($usuario.count + 1),2,
    [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
    [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent
    ) 
    $Table.Cell(1,1).range.Text = "ENTITY"
    $Table.Cell(1,2).range.Text = "NIVEL DE ACCESO"

    $j=2
    for($i=0;$i -lt ($usuario.count);$i++)
    {
        $Table.Cell($j,1).range.Text = "$($usuario[$i])"
        $Table.Cell($j,2).range.Text = "$($acceso[$i])"
        $j++
    }
    $Word.Selection.Start= $Document.Content.End
    $selection.TypeParaGraph()
}

$selection.TypeParaGraph()
#------------------------------------------------------
$selection.TypeText("ProfileList")#------------------------<
#------------------------------------------------------
$existe = Test-Path -Path "HKLM:\SOFTWARE\MICROSOFT\Windows NT\CurrentVersion\ProfileList"
if($existe -eq $false){
    $selection.TypeText("La llave ProfileList no existe")
    $selection.TypeParaGraph()
}
else
{
    $usuario = (Get-Acl "HKLM:\SOFTWARE\MICROSOFT\Windows NT\CurrentVersion\ProfileList").Access | select IdentityReference
    $acceso = (Get-Acl "HKLM:\SOFTWARE\MICROSOFT\Windows NT\CurrentVersion\ProfileList").Access | select AccessControlType
    $Table = $Selection.Tables.add(
    $Selection.Range,($usuario.count + 1),2,
    [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
    [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent
    ) 
    $Table.Cell(1,1).range.Text = "ENTITY"
    $Table.Cell(1,2).range.Text = "NIVEL DE ACCESO"

    $j=2
    for($i=0;$i -lt ($usuario.count);$i++)
    {
        $Table.Cell($j,1).range.Text = "$($usuario[$i])"
        $Table.Cell($j,2).range.Text = "$($acceso[$i])"
        $j++
    }
    $Word.Selection.Start= $Document.Content.End
    $selection.TypeParaGraph()
}

$selection.TypeParaGraph()
#------------------------------------------------------
$selection.TypeText("Compatibility32")#------------------------<
#------------------------------------------------------
$existe = Test-Path -Path "HKLM:\SOFTWARE\MICROSOFT\Windows NT\CurrentVersion\Compatibility32"
if($existe -eq $false){
    $selection.TypeText("La llave Compatibility32 no existe")
    $selection.TypeParaGraph()
}
else
{
    $usuario = (Get-Acl "HKLM:\SOFTWARE\MICROSOFT\Windows NT\CurrentVersion\Compatibility32").Access | select IdentityReference
    $acceso = (Get-Acl "HKLM:\SOFTWARE\MICROSOFT\Windows NT\CurrentVersion\Compatibility32").Access | select AccessControlType
    $Table = $Selection.Tables.add(
    $Selection.Range,($usuario.count + 1),2,
    [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
    [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent
    ) 
    $Table.Cell(1,1).range.Text = "ENTITY"
    $Table.Cell(1,2).range.Text = "NIVEL DE ACCESO"

    $j=2
    for($i=0;$i -lt ($usuario.count);$i++)
    {
        $Table.Cell($j,1).range.Text = "$($usuario[$i])"
        $Table.Cell($j,2).range.Text = "$($acceso[$i])"
        $j++
    }
    $Word.Selection.Start= $Document.Content.End
    $selection.TypeParaGraph()
}

$selection.TypeParaGraph()
#------------------------------------------------------
$selection.TypeText("Drivers32")#--------------------
#------------------------------------------------------
$existe = Test-Path -Path "HKLM:\SOFTWARE\MICROSOFT\Windows NT\CurrentVersion\Drivers32"
if($existe -eq $false){
    $selection.TypeText("La llave Drivers32 no existe")
    $selection.TypeParaGraph()
}
else
{
    $usuario = (Get-Acl "HKLM:\SOFTWARE\MICROSOFT\Windows NT\CurrentVersion\Drivers32").Access | select IdentityReference
    $acceso = (Get-Acl "HKLM:\SOFTWARE\MICROSOFT\Windows NT\CurrentVersion\Drivers32").Access | select AccessControlType
    $Table = $Selection.Tables.add(
    $Selection.Range,($usuario.count + 1),2,
    [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
    [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent
    ) 
    $Table.Cell(1,1).range.Text = "ENTITY"
    $Table.Cell(1,2).range.Text = "NIVEL DE ACCESO"

    $j=2
    for($i=0;$i -lt ($usuario.count);$i++)
    {
        $Table.Cell($j,1).range.Text = "$($usuario[$i])"
        $Table.Cell($j,2).range.Text = "$($acceso[$i])"
        $j++
    }
    $Word.Selection.Start= $Document.Content.End
    $selection.TypeParaGraph()
}

$selection.TypeParaGraph()
#------------------------------------------------------
$selection.TypeText("MCI32")#--------------------
#------------------------------------------------------
$existe = Test-Path -Path "HKLM:\SOFTWARE\MICROSOFT\Windows NT\CurrentVersion\MCI32"
if($existe -eq $false){
    $selection.TypeText("La llave MCI32 no existe")
    $selection.TypeParaGraph()
}
else
{
    $usuario = (Get-Acl "HKLM:\SOFTWARE\MICROSOFT\Windows NT\CurrentVersion\MCI32").Access | select IdentityReference
    $acceso = (Get-Acl "HKLM:\SOFTWARE\MICROSOFT\Windows NT\CurrentVersion\MCI32").Access | select AccessControlType
    $Table = $Selection.Tables.add(
    $Selection.Range,($usuario.count + 1),2,
    [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
    [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent
    ) 
    $Table.Cell(1,1).range.Text = "ENTITY"
    $Table.Cell(1,2).range.Text = "NIVEL DE ACCESO"

    $j=2
    for($i=0;$i -lt ($usuario.count);$i++)
    {
        $Table.Cell($j,1).range.Text = "$($usuario[$i])"
        $Table.Cell($j,2).range.Text = "$($acceso[$i])"
        $j++
    }
    $Word.Selection.Start= $Document.Content.End
    $selection.TypeParaGraph()
}
$selection.TypeParaGraph()

#----------------------------------------------------------
$selection.TypeText("ACLS DE ARCHIVOS EJECUTALES SENSIBLES")#--------------------
#----------------------------------------------------------
$selection.TypeParaGraph()
#------------------------------------------------------
$selection.TypeText("arp.exe")#--------------------
#------------------------------------------------------
$selection.TypeParaGraph()
$existe = Test-Path -Path "C:\Windows\system32\arp.exe"
(Get-Acl "C:\Windows\system32\arp.exe").Access | select IdentityReference
(Get-Acl "C:\Windows\system32\arp.exe").Access | select FileSystemRights

$selection.TypeParaGraph()
#------------------------------------------------------
$selection.TypeText("at.exe")#--------------------
#------------------------------------------------------
$selection.TypeParaGraph()
$existe = Test-Path -Path "C:\Windows\system32\at.exe"
(Get-Acl "C:\Windows\system32\at.exe").Access | select IdentityReference
(Get-Acl "C:\Windows\system32\at.exe").Access | select FileSystemRights

$selection.TypeParaGraph()
#------------------------------------------------------
$selection.TypeText("attrib.exe")#--------------------
#------------------------------------------------------
$selection.TypeParaGraph()
$existe = Test-Path -Path "C:\Windows\system32\attrib.exe"
(Get-Acl "C:\Windows\system32\attrib.exe").Access | select IdentityReference
(Get-Acl "C:\Windows\system32\attrib.exe").Access | select FileSystemRights

$selection.TypeParaGraph()
#------------------------------------------------------
$selection.TypeText("cacls.exe")#--------------------
#------------------------------------------------------
$selection.TypeParaGraph()
$existe = Test-Path -Path "C:\Windows\system32\cacls.exe"
(Get-Acl "C:\Windows\system32\cacls.exe").Access | select IdentityReference
(Get-Acl "C:\Windows\system32\cacls.exe").Access | select FileSystemRights

$selection.TypeParaGraph()
#------------------------------------------------------
$selection.TypeText("cmd.exe")#--------------------
#------------------------------------------------------
$selection.TypeParaGraph()
$existe = Test-Path -Path "C:\Windows\system32\cmd.exe"
(Get-Acl "C:\Windows\system32\cmd.exe").Access | select IdentityReference
(Get-Acl "C:\Windows\system32\cmd.exe").Access | select FileSystemRights

$selection.TypeParaGraph()
#------------------------------------------------------
$selection.TypeText("dcpromo.exe")#--------------------
#------------------------------------------------------
$selection.TypeParaGraph()
$existe = Test-Path -Path "C:\Windows\system32\dcpromo.exe"
(Get-Acl "C:\Windows\system32\dcpromo.exe").Access | select IdentityReference
(Get-Acl "C:\Windows\system32\dcpromo.exe").Access | select FileSystemRights

$selection.TypeParaGraph()
#------------------------------------------------------
$selection.TypeText("eventcreate.exe")#--------------------
#------------------------------------------------------
$selection.TypeParaGraph()
$existe = Test-Path -Path "C:\Windows\system32\eventcreate.exe"
(Get-Acl "C:\Windows\system32\eventcreate.exe").Access | select IdentityReference
(Get-Acl "C:\Windows\system32\eventcreate.exe").Access | select FileSystemRights

$selection.TypeParaGraph()
#------------------------------------------------------
$selection.TypeText("finger.exe")#--------------------
#------------------------------------------------------
$selection.TypeParaGraph()
$existe = Test-Path -Path "C:\Windows\system32\finger.exe"
(Get-Acl "C:\Windows\system32\finger.exe").Access | select IdentityReference
(Get-Acl "C:\Windows\system32\finger.exe").Access | select FileSystemRights

$selection.TypeParaGraph()
#------------------------------------------------------
$selection.TypeText("ftp.exe")#--------------------
#------------------------------------------------------
$selection.TypeParaGraph()
$existe = Test-Path -Path "C:\Windows\system32\ftp.exe"
(Get-Acl "C:\Windows\system32\ftp.exe").Access | select IdentityReference
(Get-Acl "C:\Windows\system32\ftp.exe").Access | select FileSystemRights

$selection.TypeParaGraph()
#------------------------------------------------------
$selection.TypeText("gpupdate.exe")#--------------------
#------------------------------------------------------
$selection.TypeParaGraph()
$existe = Test-Path -Path "C:\Windows\system32\gpupdate.exe"
(Get-Acl "C:\Windows\system32\gpupdate.exe").Access | select IdentityReference
(Get-Acl "C:\Windows\system32\gpupdate.exe").Access | select FileSystemRights

$selection.TypeParaGraph()
#------------------------------------------------------
$selection.TypeText("icacls.exe")#--------------------
#------------------------------------------------------
$selection.TypeParaGraph()
$existe = Test-Path -Path "C:\Windows\system32\icacls.exe"
(Get-Acl "C:\Windows\system32\icacls.exe").Access | select IdentityReference
(Get-Acl "C:\Windows\system32\icacls.exe").Access | select FileSystemRights

$selection.TypeParaGraph()
#------------------------------------------------------
$selection.TypeText("ipconfig.exe")#--------------------
#------------------------------------------------------
$selection.TypeParaGraph()
$existe = Test-Path -Path "C:\Windows\system32\ipconfig.exe"
(Get-Acl "C:\Windows\system32\ipconfig.exe").Access | select IdentityReference
(Get-Acl "C:\Windows\system32\ipconfig.exe").Access | select FileSystemRights

$selection.TypeParaGraph()
#------------------------------------------------------
$selection.TypeText("nbtstat.exe")#--------------------
#------------------------------------------------------
$selection.TypeParaGraph()
$existe = Test-Path -Path "C:\Windows\system32\nbtstat.exe"
(Get-Acl "C:\Windows\system32\nbtstat.exe").Access | select IdentityReference
(Get-Acl "C:\Windows\system32\nbtstat.exe").Access | select FileSystemRights

$selection.TypeParaGraph()
#------------------------------------------------------
$selection.TypeText("net.exe")#--------------------
#------------------------------------------------------
$selection.TypeParaGraph()
$existe = Test-Path -Path "C:\Windows\system32\net.exe"
(Get-Acl "C:\Windows\system32\net.exe").Access | select IdentityReference
(Get-Acl "C:\Windows\system32\net.exe").Access | select FileSystemRights

$selection.TypeParaGraph()
#------------------------------------------------------
$selection.TypeText("net1.exe")#--------------------
#------------------------------------------------------
$selection.TypeParaGraph()
$existe = Test-Path -Path "C:\Windows\system32\net1.exe"
(Get-Acl "C:\Windows\system32\net1.exe").Access | select IdentityReference
(Get-Acl "C:\Windows\system32\net1.exe").Access | select FileSystemRights

$selection.TypeParaGraph()
#------------------------------------------------------
$selection.TypeText("netsh.exe")#--------------------
#------------------------------------------------------
$selection.TypeParaGraph()
$existe = Test-Path -Path "C:\Windows\system32\netsh.exe"
(Get-Acl "C:\Windows\system32\netsh.exe").Access | select IdentityReference
(Get-Acl "C:\Windows\system32\netsh.exe").Access | select FileSystemRights

$selection.TypeParaGraph()
#------------------------------------------------------
$selection.TypeText("netstat.exe")#--------------------
#------------------------------------------------------
$selection.TypeParaGraph()
$existe = Test-Path -Path "C:\Windows\system32\netstat.exe"
(Get-Acl "C:\Windows\system32\netstat.exe").Access | select IdentityReference
(Get-Acl "C:\Windows\system32\netstat.exe").Access | select FileSystemRights

$selection.TypeParaGraph()
#------------------------------------------------------
$selection.TypeText("nslookup.exe")#--------------------
#------------------------------------------------------
$selection.TypeParaGraph()
$existe = Test-Path -Path "C:\Windows\system32\nslookup.exe"
(Get-Acl "C:\Windows\system32\nslookup.exe").Access | select IdentityReference
(Get-Acl "C:\Windows\system32\nslookup.exe").Access | select FileSystemRights

$selection.TypeParaGraph()
#------------------------------------------------------
$selection.TypeText("ping.exe")#--------------------
#------------------------------------------------------
$selection.TypeParaGraph()
$existe = Test-Path -Path "C:\Windows\system32\ping.exe"
(Get-Acl "C:\Windows\system32\ping.exe").Access | select IdentityReference
(Get-Acl "C:\Windows\system32\ping.exe").Access | select FileSystemRights

$selection.TypeParaGraph()
#------------------------------------------------------
$selection.TypeText("reg.exe")#--------------------
#------------------------------------------------------
$selection.TypeParaGraph()
$existe = Test-Path -Path "C:\Windows\system32\reg.exe"
(Get-Acl "C:\Windows\system32\reg.exe").Access | select IdentityReference
(Get-Acl "C:\Windows\system32\reg.exe").Access | select FileSystemRights

$selection.TypeParaGraph()
#------------------------------------------------------
$selection.TypeText("regedt32.exe")#--------------------
#------------------------------------------------------
$selection.TypeParaGraph()
$existe = Test-Path -Path "C:\Windows\system32\regedt32.exe"
(Get-Acl "C:\Windows\system32\regedt32.exe").Access | select IdentityReference
(Get-Acl "C:\Windows\system32\regedt32.exe").Access | select FileSystemRights

$selection.TypeParaGraph()
#------------------------------------------------------
$selection.TypeText("regini.exe")#--------------------
#------------------------------------------------------
$selection.TypeParaGraph()
$existe = Test-Path -Path "C:\Windows\system32\regini.exe"
(Get-Acl "C:\Windows\system32\regini.exe").Access | select IdentityReference
(Get-Acl "C:\Windows\system32\regini.exe").Access | select FileSystemRights

$selection.TypeParaGraph()
#------------------------------------------------------
$selection.TypeText("regsvr32.exe")#--------------------
#------------------------------------------------------
$selection.TypeParaGraph()
$existe = Test-Path -Path "C:\Windows\system32\regsvr32.exe"
(Get-Acl "C:\Windows\system32\regsvr32.exe").Access | select IdentityReference
(Get-Acl "C:\Windows\system32\regsvr32.exe").Access | select FileSystemRights

$selection.TypeParaGraph()
#------------------------------------------------------
$selection.TypeText("route.exe")#--------------------
#------------------------------------------------------
$selection.TypeParaGraph()
$existe = Test-Path -Path "C:\Windows\system32\route.exe"
(Get-Acl "C:\Windows\system32\route.exe").Access | select IdentityReference
(Get-Acl "C:\Windows\system32\route.exe").Acces | select FileSystemRights

$selection.TypeParaGraph()
#------------------------------------------------------
$selection.TypeText("runonce.exe")#--------------------
#------------------------------------------------------
$selection.TypeParaGraph()
$existe = Test-Path -Path "C:\Windows\system32\runonce.exe"
(Get-Acl "C:\Windows\system32\runonce.exe").Access | select IdentityReference
(Get-Acl "C:\Windows\system32\runonce.exe").Access | select FileSystemRights

$selection.TypeParaGraph()
#------------------------------------------------------
$selection.TypeText("sc.exe")#--------------------
#------------------------------------------------------
$selection.TypeParaGraph()
$existe = Test-Path -Path "C:\Windows\system32\sc.exe"
(Get-Acl "C:\Windows\system32\sc.exe").Access | select IdentityReference
(Get-Acl "C:\Windows\system32\sc.exe").Access | select FileSystemRights

$selection.TypeParaGraph()
#------------------------------------------------------
$selection.TypeText("secedit.exe")#--------------------
#------------------------------------------------------
$selection.TypeParaGraph()
$existe = Test-Path -Path "C:\Windows\system32\secedit.exe"
(Get-Acl "C:\Windows\system32\secedit.exe").Access | select IdentityReference
(Get-Acl "C:\Windows\system32\secedit.exe").Access | select FileSystemRights

$selection.TypeParaGraph()
#------------------------------------------------------
$selection.TypeText("subst.exe")#--------------------
#------------------------------------------------------
$selection.TypeParaGraph()
$existe = Test-Path -Path "C:\Windows\system32\subst.exe"
(Get-Acl "C:\Windows\system32\subst.exe").Access | select IdentityReference
(Get-Acl "C:\Windows\system32\subst.exe").Access | select FileSystemRights

$selection.TypeParaGraph()
#------------------------------------------------------
$selection.TypeText("systeminfo.exe")#--------------------
#------------------------------------------------------
$selection.TypeParaGraph()
$existe = Test-Path -Path "C:\Windows\system32\systeminfo.exe"
(Get-Acl "C:\Windows\system32\systeminfo.exe").Access | select IdentityReference
(Get-Acl "C:\Windows\system32\systeminfo.exe").Access | select FileSystemRights

$selection.TypeParaGraph()
#------------------------------------------------------
$selection.TypeText("syskey.exe")#--------------------
#------------------------------------------------------
$selection.TypeParaGraph()
$existe = Test-Path -Path "C:\Windows\system32\syskey.exe"
(Get-Acl "C:\Windows\system32\syskey.exe").Access | select IdentityReference
(Get-Acl "C:\Windows\system32\syskey.exe").Access | select FileSystemRights

$selection.TypeParaGraph()
#------------------------------------------------------
$selection.TypeText("telnet.exe")#--------------------
#------------------------------------------------------
$selection.TypeParaGraph()
$existe = Test-Path -Path "C:\Windows\system32\telnet.exe"
(Get-Acl "C:\Windows\system32\telnet.exe").Access | select IdentityReference
(Get-Acl "C:\Windows\system32\telnet.exe").Access | select FileSystemRights

$selection.TypeParaGraph()
#------------------------------------------------------
$selection.TypeText("tftp.exe")#--------------------
#------------------------------------------------------
$selection.TypeParaGraph()
$existe = Test-Path -Path "C:\Windows\system32\tftp.exe"
(Get-Acl "C:\Windows\system32\tftp.exe").Access | select IdentityReference
(Get-Acl "C:\Windows\system32\tftp.exe").Access | select FileSystemRights

$selection.TypeParaGraph()
#------------------------------------------------------
$selection.TypeText("tlntsvr.exe")#--------------------
#------------------------------------------------------
$selection.TypeParaGraph()
$existe = Test-Path -Path "C:\Windows\system32\tlntsvr.exe"
(Get-Acl "C:\Windows\system32\tlntsvr.exe").Access | select IdentityReference
(Get-Acl "C:\Windows\system32\tlntsvr.exe").Access | select FileSystemRights


$selection.TypeParaGraph()
#------------------------------------------------------
$selection.TypeText("tracert.exe")#--------------------
#------------------------------------------------------
$selection.TypeParaGraph()
$existe = Test-Path -Path "C:\Windows\system32\tracert.exe"
(Get-Acl "C:\Windows\system32\tracert.exe").Access | select IdentityReference
(Get-Acl "C:\Windows\system32\tracert.exe").Access | select FileSystemRights

$selection.TypeParaGraph()
#------------------------------------------------------
$selection.TypeText("xcopy.exe")#--------------------
#------------------------------------------------------
$selection.TypeParaGraph()
$existe = Test-Path -Path "C:\Windows\system32\xcopy.exe"
(Get-Acl "C:\Windows\system32\xcopy.exe").Access | select IdentityReference
(Get-Acl "C:\Windows\system32\xcopy.exe").Access | select FileSystemRights

#--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------                            
                        }
                    }
                    else
                    {
#-------------------------------------------------------VALIDACION SIN REPORTE---------------------------------------------------------------------------------------------------
                        $comando = 'Invoke-Command -scriptBlock {' + $scriptReportetxt + '}'
                        $comando.ToString()
                        Invoke-Expression $comando
#--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                    }
                }
                ElseIf($Entorno -eq "remote")
                {
                    if($Reporte)
                    {
                        if( (Test-Path -Path "HKLM:\SOFTWARE\Microsoft\Office") -eq $false ){Write-Host "No existe Microsoft Office, Imposible generar el reporte" -ForegroundColor Red -BackgroundColor black}   
                        else
                        {
                            Invoke-Command -Session $session -ScriptBlock {Get-Host}
                            Write-Host "Reporte"
                        }
                    }
                    else
                    {
#--------------------------------------------------------REMOTO SIN REPORTE-------------------------------------------------------------------------------------------------------                    
                        Invoke-Command -Session $session -ScriptBlock $scriptReportetxt                       
#---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------                        
                        
                        
                    }
                }
                
		    }#Cierre PROCESS
            
	    END {
                if($Entorno -eq "remote")
                {
                    $sessions = Get-PSSession
                    if( ($sessions).Count -eq $null){$e = $sessions.id; Remove-PSSession -Id $e}
                    else
                    { 
                        for($i=0; $i -lt ($sessions).Count; $i++)
                        {
                            $e = ($sessions[$i]).id
                            Remove-PSSession -Id $e
                        }
                    }
                }  
		    }#Cierre END
            
}#Cierre FUNCION