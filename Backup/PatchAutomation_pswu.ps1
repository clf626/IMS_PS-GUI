<#
.SYNOPSIS
This script will automatically install all avaialable windows updates on a device and will automatically reboot if needed, after reboot, windows updates will continue to run until no more updates are available.
.PARAMETER URL
User the Computer parameter to specify the Computer to remotely install windows updates on.
#>

[CmdletBinding()]

param (

[parameter(Mandatory=$true,Position=1)]
[string[]]$computer


)

ForEach ($c in $computer){

    
    
    #install pswindows updates module on remote machine
    $nugetinstall = invoke-command -ComputerName $c -ScriptBlock {Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force}
    invoke-command -ComputerName $c -ScriptBlock {install-module pswindowsupdate -force}

    invoke-command -ComputerName $c -ScriptBlock {Import-Module PSWindowsUpdate -force}

    Do{
        #Reset Timeouts
        $connectiontimeout = 0
        $updatetimeout = 0
        
        #starts up a remote powershell session to the computer
        do{
            $session = New-PSSession -ComputerName $c
            "reconnecting remotely to $c"
            sleep -seconds 10
            $connectiontimeout++
        } until ($session.state -match "Opened" -or $connectiontimeout -ge 10)

        #retrieves a list of available updates

        "Checking for new updates available on $c"

        $updates = invoke-command -session $session -scriptblock {Get-wulist -verbose}

        #counts how many updates are available

        $updatenumber = ($updates.kb).count

        #if there are available updates proceed with installing the updates and then reboot the remote machine

        if ($updates -ne $null){

            #remote command to install windows updates, creates a scheduled task on remote computer

            invoke-command -ComputerName $c -ScriptBlock { Invoke-WUjob -ComputerName localhost -Script "ipmo PSWindowsUpdate; Install-WindowsUpdate -AcceptAll | Out-File C:\PSWindowsUpdate.log" -Confirm:$false -RunNow}

            #Show update status until the amount of installed updates equals the same as the amount of updates available

            sleep -Seconds 30

            do {$updatestatus = Get-Content \\$c\c$\PSWindowsUpdate.log

                "Currently processing the following update:"

                Get-Content \\$c\c$\PSWindowsUpdate.log | select-object -last 1

                sleep -Seconds 10

                $ErrorActionPreference = ‘SilentlyContinue’

                $installednumber = ([regex]::Matches($updatestatus, "Installed" )).count

                $Failednumber = ([regex]::Matches($updatestatus, "Failed" )).count

                $ErrorActionPreference = ‘Continue’

                $updatetimeout++


            }until ( ($installednumber + $Failednumber) -eq $updatenumber -or $updatetimeout -ge 720)

            #restarts the remote computer and waits till it starts up again

            "restarting remote computer"

             #removes schedule task from computer

            invoke-command -computername $c -ScriptBlock {Unregister-ScheduledTask -TaskName PSWindowsUpdate -Confirm:$false}

             # rename update log
             $date = Get-Date -Format "MM-dd-yyyy_hh-mm-ss"
             Rename-Item \\$c\c$\PSWindowsUpdate.log -NewName "WindowsUpdate-$date.log"

            Restart-Computer -Wait -ComputerName $c -Force

        }
   

    }until($updates -eq $null)

    

    "Windows is now up to date on $c"

}