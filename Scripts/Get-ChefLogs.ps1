Function Get-ChefLogs {
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
            
          #Get Chef logs
            try {
            $cheflogs = Get-content "\\$c\c$\Users\Administrator\patch.log" -Tail 10
            }
            Catch{
            $error =  "$($Error[0])"
            }
                            Write-Verbose "Creating report"
                            $temp = "" | Select Computer,Logs,Error
                            $temp.Computer = $c
                            $temp.Logs = $cheflogs
                            $temp.Error = $erroCount
                            $report += $temp                  
                
           
            }
        Else {
            #Nothing to install at this time
            Write-Warning "$($c): Offline"
            
            #Create Temp collection for report
           $temp = "" | Select Computer,Logs,Error
                            $temp.Computer = $c
                            $temp.Logs = "Offline"
                            $temp.Error = "Offline"
                            $report += $temp           
            }
        } 
    }
End {
    Write-Output $report
    $report | Export-Csv -path .\PatchLogsExcel_$((Get-Date).ToString('MM-dd-yyyy_hh-mm-ss')).csv
    }    
}