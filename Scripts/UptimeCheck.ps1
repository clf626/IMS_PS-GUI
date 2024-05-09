 $names = Get-Content "C:\Users\cfontan5\Documents\test_servers.txt" 
      @( 
           foreach ($name in $names) 
          { 
            try{
              if ( Test-Connection -ComputerName $name -Count 1 -ErrorAction SilentlyContinue ){ 
                 
                   $wmi = gwmi Win32_OperatingSystem -computer $name 
                   $LBTime = $wmi.ConvertToDateTime($wmi.Lastbootuptime) 
                   [TimeSpan]$uptime = New-TimeSpan $LBTime $(get-date) 
                   Write-output "$name Uptime is  $($uptime.days) Days $($uptime.hours) Hours $($uptime.minutes) Minutes $($uptime.seconds) Seconds" 
                   write-output 
               }   
               else{ 
                    Write-output "$name is not pinging" 
               }
             }   
             catch{
                Write-Output "$($name, $Error[0])"
             }
              
           }
               
       ) | Out-file -FilePath "C:\temp\UptimePS\results.txt"

       Get-service -DisplayName Ops* -Verbose
Get-Service -DisplayName Sym* -Verbose
Get-Service -name Netlogon -Verbose
Get-WmiObject win32_operatingsystem | select csname, @{LABEL='LastBootUpTime';EXPRESSION={$_.ConverttoDateTime($_.lastbootuptime)}}
If (New-Object System.Net.Sockets.TCPClient -ArgumentList $env:COMPUTERNAME,3389) { Write-Host 'RDP Connection is open' }
If ($? -eq $false) { Write-Host 'Something went wrong - unable to connect' } 


e2 ung name ng scripts sa HPSA -- "SCR_WSQC_PS_ALL_Post-Patch.ps1"