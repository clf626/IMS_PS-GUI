Function Get-PatchLog {
[cmdletbinding()]
Param (
    [parameter()]
    [string[]]$Computername,
    [parameter()]
    [int32]$Last
    )
Begin {
    }
Process{    
    ForEach ($computer in $computername) {
        $aftrDate = (get-date).AddDays(-5).ToString("MM/dd/yyyy") + " 21:00:00"
        $b4Date = (get-date).AddDays(-4).ToString("MM/dd/yyyy")
        $logpath = Get-EventLog -Logname "Application" -ComputerName $computer -Source "Optum_Patching" -after $aftrDate -Before $b4Date | where {$_.EventID -eq 880}
        $updatelog = "\\{0}\{1}" -f $computer,($logpath -replace ":","$")
        If ($PSBoundParameters['Last']) {    
            [array]$log = get-content $updatelog |
                Select -Last $last
                $log
            }
        Else {
            [array]$log = get-content $updatelog
            $log
            }
        #ForEach ($line in $log) {
        #    $split = $line -split "`t"
        #    $hash = @{
        #        Computer = $computer
        #        Date = $split[1]
        #        EntryType = $split[2]
        #        Source = $split[3]
        #        Message = $split[5]
        #        }
        #    $object = New-Object PSObject -Property $hash
        #    $object.PSTypeNames.Insert(0,'UpdateLog')
        #    $object
        #    }
        }
    }
End {
    }
}