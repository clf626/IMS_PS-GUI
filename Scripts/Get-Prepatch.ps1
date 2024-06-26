Function Get-PrePatch {
[cmdletbinding()]
param($Computer)    
Begin {
    If (!$PSBoundParameters['computer']) {
        Write-Verbose "No computer name given, using local computername"
        [string]$computer = $Env:Computername
        }
    #Create container for Report
    Write-Verbose "Creating report collection"
    $report = @()
    $erroCount = 0    
    }
Process {
    ForEach ($c in $Computer) {
        Write-Verbose "Computer: $($c)"
        If (Test-Connection -ComputerName $c -Count 1 -Quiet) {
            
          #Check OS
            try {
            $OS = (gwmi Win32_OperatingSystem -computer $c).caption
            }
            Catch{
            #$OS = "$($Error[0])"
            $OS = "Error getting OS version"
            $erroCount += 1
            }

          #Check PING
            Try{
            $PING = "Success"}
            Catch{
            $PING = "Error Getting PING"
            $erroCount += 1
            }
            
          #Check Uptime/LastBootUpTime
            Try{
            $Getping = gwmi Win32_OperatingSystem -computer $c
                   $LBTime = $Getping.ConvertToDateTime($Getping.Lastbootuptime) 
                   [TimeSpan]$uptime = New-TimeSpan $LBTime $(get-date)
                   $UT = "$($uptime.days)day $($uptime.hours)hr $($uptime.minutes)min $($uptime.seconds)sec" 
            }
            Catch{
            $UT = "Error getting Uptime"
            $LBTime = "Error getting LastBootUpTime"
            $erroCount += 1
            }

          #Check RDP
            Try{
            $checkRDP = (New-Object System.Net.Sockets.TCPClient -ArgumentList $c ,3389 | select Connected).connected
            $rdp = $checkRDP
            }Catch{
            $rdp = "Error getting RDP status"
            $erroCount += 1
            }

          #Check Netlogon
            Try{
            $netlogon = (Get-Service -ComputerName $c -DisplayName "netlogon" | select Status).Status
            $NETlogon = $netlogon
            }Catch{
            $NETlogon = "Error getting Netlogon status"
            $erroCount += 1
            }
            
          #Check Script Folder
            try {
            $CF = Test-Path -path "\\$c\c$\users\administrator\scripts\"
            }
            Catch{
            #$CF = "$($Error[0])"
            $CF = "Error getting Script folder existence"
            $erroCount += 1
            }
            
          #Check CHEF Version
            try {
            $CHEFversion = (gwmi -ComputerName $c -Class Win32_Product | select name, version | where-object {$_.name  -Like "Chef*"})."version"
            }
            Catch{
            #$CHEFversion = "$($Error[0])"
            $CHEFversion = "Error getting CHEF version"
            $erroCount += 1
            }

          #Check CDrive Space
            try {
            $disk = gwmi Win32_LogicalDisk -ComputerName $c -Filter "DeviceID='C:'" | Select-Object Size,FreeSpace
                   $csize = [math]::Round($disk.Size / 1GB,1)
                   $free = [math]::Round($disk.FreeSpace / 1GB,1)
            }
            Catch{
            #$free = "$($Error[0])"
            $free = "Error getting C drive free space"
            $erroCount += 1
            }
            
            
          #Check SEP Version
            try {
            #$SEPversion = (gwmi -Class win32_product -ComputerName $c -Filter "name LIKE 'Symantec Endpoint Protection'")."version"
            $SEPversion = Invoke-Command -ComputerName $c -ScriptBlock{
                $ver = (Get-ItemProperty 'HKLM:\SOFTWARE\Symantec\Symantec Endpoint Protection\CurrentVersion' | select "PRODUCTVERSION")."PRODUCTVERSION"
                if($ver -ne $null){return "$ver"}
                Else{return "SEP not installed"}
            }
            }
            Catch{
            #$SEPversion = "$($Error[0])"
            $sepversion = "Error gettting SEP version"
            $erroCount += 1
            }

          #Check SEP Definition
            try {
            #$SEPdef = Invoke-Command -cn $c -scriptblock {Get-ChildItem -Path  "C:\ProgramData\Symantec\Symantec Endpoint Protection\CurrentVersion\Data\Definitions\SDSDefs" | where-object {$_.Name -notmatch "[a-z]"}}
              $SEPdef = Invoke-Command -cn $c -scriptblock {
                $def = (get-ItemProperty 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Symantec\Symantec Endpoint Protection\CurrentVersion\public-opstate'  | select LatestVirusDefsDate).LatestVirusDefsDate
                $rev = (get-ItemProperty 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Symantec\Symantec Endpoint Protection\CurrentVersion\public-opstate' | select LatestVirusDefsRevision).LatestVirusDefsRevision
                if($def -ne $null){return "$def r$rev"}
                Else{return "SEP not installed"}
              }
            }Catch{
            $SEPdef = "Error gettting SEP version"
            $erroCount += 1
            }

          #Check SEP Status
            try {
            $chkSEPService = (Get-Service -ComputerName $c -DisplayName "Symantec Endpoint Protection" | select Status).Status
            if($chkSEPService -ne $null){$SEPService = $chkSEPService}
                Else{$SEPService = "SEP not installed"}
            }
            Catch{
            $SEPService = "Error gettting SEP status"
            #$SEPService = "$($Error[0])"
            $erroCount += 1
            }
            
         #Get WSUS Target Group
           try{
             $key = 'SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate'
             $valuename = 'TargetGroup'
             $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $c)
             $regkey = $reg.opensubkey($key)
             $WSUS = $regkey.getvalue($valuename)
            }Catch{
            #$WSUS = "$($Error[0])"
            $WSUS = "Error gettting WSUS Target Group"
            $erroCount += 1
            }

         #Get PSVersion
            Try{
            $PSver = Invoke-Command -ComputerName $c -ScriptBlock {$psversiontable.psversion.major}
            }Catch{
            $PSver = "Error gettting WSUS Target Group"
            $erroCount += 1
            }

         #Get DotnetVersion
            try{
             $key = 'SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\full'
             $valuename = 'Release'
             $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $c)
             $regkey = $reg.opensubkey($key)
             $Release = $regkey.getvalue($valuename)
                #$Release = Invoke-Command -ComputerName $c -ScriptBlock {
                #(Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full' | select "Release")."Release"}
               
	            If ($Release -ge 528040)
	            	{$DotNet = "4.8.0"}
	            	elseif ($Release -ge 461808)
	            		{$DotNet = "4.7.2"}
	            	elseif ($Release -ge 461308)
	            		{$DotNet = "4.7.1"}
	            	elseif ($Release -ge 460798) 
	            		{$DotNet = "4.7.0"}
	            	elseif ($Release -ge 394802) 
	            		{$DotNet = "4.6.2"}
	            	elseif ($Release -ge 394254) 
	            		{$DotNet = "4.6.1"}
	            	elseif ($Release -ge 393295) 
	            		{$DotNet = "4.6.0"}
	            	elseif ($Release -ge 379893) 
	            		{$DotNet = "4.5.2"}
	            	elseif ($Release -ge 378675) 
	            		{$DotNet = "4.5.1"}
	            	elseif ($Release -ge 378389) 
	            		{$DotNet = "4.5.0"}
	            else
	            	{$DotNet = "No 4.5 or later version detected."}
            }Catch{
            $DotNet = "Error gettting WSUS Target Group"
            $erroCount += 1
            }

                            Write-Verbose "Creating report"
                            $temp = "" | Select Computer, OS, RDP, NetLogon, PING, LastBootUpTime, Uptime,CHEF_folder,CHEF_version,C_drive,SEP_version,SEP_def,SEP_service,WSUS_group,PSVersion,DotNet,Error
                            $temp.Computer = $c
                            $temp.OS = $OS
                            $temp.RDP = $rdp
                            $temp.NetLogon = $NETlogon
                            $temp.PING = $PING
                            $Temp.LastBootUpTime = $LBtime
                            $temp.Uptime = $UT
                            $temp.CHEF_folder = $CF
                            $temp.CHEF_version = $CHEFversion
                            $temp.C_drive = $free
                            $temp.SEP_version = $SEPversion
                            $temp.SEP_def = $SEPdef
                            $temp.SEP_service = $SEPService
                            $temp.WSUS_group = $WSUS
                            $temp.PSVersion = $PSver
                            $temp.DotNet = $DotNet
                            $temp.Error = $erroCount
                            $report += $temp                  
                
           
            }
        Else {
            #Nothing to install at this time
            Write-Warning "$($c): Offline"
            
            #Create Temp collection for report
           $temp = "" | Select Computer, OS, RDP, NetLogon, PING, LastBootUpTime, Uptime,CHEF_folder,CHEF_version,C_drive,SEP_version,SEP_def,SEP_service,WSUS_group,PSVersion,DotNet,Error
                            $temp.Computer = $c
                            $temp.OS = "Offline"
                            $temp.PING = "Offline"
                            $Temp.LastBootUpTime = "Offline"
                            $temp.Uptime = "Offline"
                            $temp.RDP = "Offline"
                            $temp.NetLogon = "Offline"
                            $temp.CHEF_folder = "Offline"
                            $temp.CHEF_version = "Offline"
                            $temp.C_drive = "Offline"
                            $temp.SEP_version = "Offline"
                            $temp.SEP_def = "Offline"
                            $temp.SEP_service = "Offline"
                            $temp.WSUS_group = "Offline"
                            $temp.PSVersion = "Offline"
                            $temp.DotNet = "Offline"
                            $temp.Error = "Offline"
                            $report += $temp           
            }
        } 
    }
End {
    Write-Output $report
    }    
}