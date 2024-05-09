$userInput = Read-Host 'Enter Server Name:'
if ( $userInput -ne $null ) 
{
 while ( (get-service -Name "TSM*proxy*sched*" -ComputerName $userInput).Status -eq "Running" ) {
              "TSM still Running"
              Start-Sleep -Seconds 10
          }
          "TSM is stopped"
}
else
{
 echo "No server input"
}