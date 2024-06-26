Function Get-PendingUpdates {
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
    }
Process {
    ForEach ($c in $Computer) {
        Write-Verbose "Computer: $($c)"
        If (Test-Connection -ComputerName $c -Count 1 -Quiet) {
            Try {
            
            $OS = (gwmi Win32_OperatingSystem -computer $c).caption

            $Getping = gwmi Win32_OperatingSystem -computer $c
                   $LBTime = $Getping.ConvertToDateTime($Getping.Lastbootuptime) 
                   [TimeSpan]$uptime = New-TimeSpan $LBTime $(get-date)
                   if ($Getping) {
                   $ping = "PING:Success : LastBootUpTime:$LBtime : Uptime: $($uptime.days)Day $($uptime.hours)Hr $($uptime.minutes)Min $($uptime.seconds)Sec" 
                   } else {
                   $ping = "PING:Success but PSexec not present at the server."}
            
            $checkRDP = (New-Object System.Net.Sockets.TCPClient -ArgumentList $c ,3389 | select Connected).connected
            $netlogon = (Get-Service -ComputerName $c -DisplayName "netlogon" | select Status).Status
            $rdpNETlogon = "RDP:$checkRDP : Netlogon:$netlogon"
            
            $CF = Test-Path -path "\\$c\c$\users\administrator\scripts\"
            
            $CHEFversion = (gwmi -ComputerName $c -Class Win32_Product | select name, version | where-object {$_.name  -Like "Chef*"})."version"
                   if ($CHEFversion){
                   $CV = $CHEFversion }
                   Else{
                   $CV = "No software installed!"}
            
            $disk = gwmi Win32_LogicalDisk -ComputerName $c -Filter "DeviceID='C:'" | Select-Object Size,FreeSpace
                   $csize = [math]::Round($disk.Size / 1GB,1)
                   $cfree = [math]::Round($disk.FreeSpace / 1GB,1)
                   if ($disk){
                   $free = "$cfree"}
                   Else{
                   $free = "No PSexec installed!"}
            
            $SEPversion = (gwmi -Class win32_product -ComputerName $c -Filter "name LIKE 'Symantec Endpoint Protection'")."version"
            $SEPService = (Get-Service -ComputerName $c -DisplayName "Symantec Endpoint Protection" | select Status).Status
                   if ($SEPversion){
                   $SEP = $SEPversion}
                   Else{
                   $SEP = "No software installed!"}

             $key = 'SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate'
             $valuename = 'TargetGroup'
             $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $c)
             $regkey = $reg.opensubkey($key)
             $WSUS = $regkey.getvalue($valuename)

                            Write-Verbose "Creating report"
                            $temp = "" | Select Computer, OS, RDP, PING_result,CHEF_folder,CHEF_version,C_drive,SEP_version,SEP_service,WSUS_group
                            $temp.Computer = $c
                            $temp.OS = $OS
                            $temp.RDP = $rdpNETlogon
                            $temp.PING_result = $ping
                            $temp.CHEF_folder = $CF
                            $temp.CHEF_version = $CV
                            $temp.C_drive = $free
                            $temp.SEP_version = $SEP
                            $temp.SEP_service = $SEPService
                            $temp.WSUS_group = $WSUS
                            $report += $temp                  
                }
            Catch {
                Write-Warning "$($Error[0])"
                Write-Verbose "Creating report"
                #Create Temp collection for report
                 $temp = "" |  Select Computer, OS, RDP, PING_result,CHEF_folder,CHEF_version,C_drive,SEP_version,SEP_service,WSUS_group
                            $temp.Computer = $c
                            $temp.OS = "Error"
                            $temp.RDP = "Error"
                            $temp.PING_result = "Error"
                            $temp.CHEF_folder = "Error"
                            $temp.CHEF_version = "Error"
                            $temp.C_drive = "Error"
                            $temp.SEP_version = "Error"
                            $temp.SEP_service = "Error"
                            $temp.WSUS_group = "Error"
                            $report += $temp  
                }
            }
        Else {
            #Nothing to install at this time
            Write-Warning "$($c): Offline"
            
            #Create Temp collection for report
           $temp = "" | Select Computer, OS, RDP, PING_result,CHEF_folder,CHEF_version,C_drive,SEP_version,SEP_service,WSUS_group
                            $temp.Computer = $c
                            $temp.OS = "Offline"
                            $temp.PING_result = "Offline"
                            $temp.RDP = "Offline"
                            $temp.CHEF_folder = "Offline"
                            $temp.CHEF_version = "Offline"
                            $temp.C_drive = "Offline"
                            $temp.SEP_version = "Offline"
                            $temp.SEP_service = "Offline"
                            $temp.WSUS_group = "Offline"
                            $report += $temp           
            }
        } 
    }
End {
    Write-Output $report
    }    
}