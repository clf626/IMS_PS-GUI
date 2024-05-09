Function Get-PostPatch {
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
            
          #Check SEP Status
            try {
            $SEPService = (Get-Service -ComputerName $c -DisplayName "Symantec Endpoint Protection" | select Status).Status
            }
            Catch{
            $SEPService = "Error gettting SEP status"
            #$SEPService = "$($Error[0])"
            $erroCount += 1
            }

                            Write-Verbose "Creating report"
                            $temp = "" | Select Computer, RDP, NetLogon, PING, LastBootUpTime, Uptime, SEP_service,Error
                            $temp.Computer = $c
                            $temp.RDP = $rdp
                            $temp.NetLogon = $NETlogon
                            $temp.PING = $PING
                            $Temp.LastBootUpTime = $LBtime
                            $temp.Uptime = $UT
                            $temp.SEP_service = $SEPService
                            $temp.Error = $erroCount
                            $report += $temp                  
                
           
            }
        Else {
            #Nothing to install at this time
            Write-Warning "$($c): Offline"
            
            #Create Temp collection for report
                            $temp = "" | Select Computer, RDP, NetLogon, PING, LastBootUpTime, Uptime, SEP_service,Error
                            $temp.Computer = $c
                            $temp.RDP = "Offline"
                            $temp.NetLogon = "Offline"
                            $temp.PING = "Offline"
                            $Temp.LastBootUpTime = "Offline"
                            $temp.Uptime = "Offline"
                            $temp.SEP_service = "Offline"
                            $temp.Error = "Offline"
                            $report += $temp            
            }
        } 
    }
End {
    Write-Output $report
    }    
}