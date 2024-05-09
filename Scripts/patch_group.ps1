$ServerListFile = "C:\Users\cfontan5\Documents\PoshPAIG-master\Servers.txt"
$ServerList = Get-Content $ServerListFile

ForEach($ServerName in $ServerList) {
Echo $ServerName

    if (Test-Connection -ComputerName $ServerName -Quiet)
        {
        #Test for logon service running, if running then do the commands.
        #Test for the following, if passes found then do the commands.
        #"The specified network name is no longer available."
        #"There is not enough space on the disk."
        #"The network path was not found."
        #"Logon failure: unknown user name or bad password."
        #"The trust relationship between the primary domain and the trusted domain failed."

        #:: Run remote patch_now
        Invoke-Command -ComputerName $ServerName -ScriptBlock {C:\Users\Administrator\scripts\patch_now.ps1 -n -q}
        }
}
