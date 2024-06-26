Function Create-UpdateVBS {
Param ($computername)
    #Create Here-String of vbscode to create file on remote system
    $vbsstring = @"
ON ERROR RESUME NEXT
CONST ForAppending = 8
CONST ForWriting = 2
CONST ForReading = 1
strlocalhost = "."
Set oShell = CreateObject("WScript.Shell") 
set ofso = createobject("scripting.filesystemobject")
Set updateSession = CreateObject("Microsoft.Update.Session")
Set updateSearcher = updateSession.CreateupdateSearcher()
Set updatesToInstall = CreateObject("Microsoft.Update.UpdateColl")
Set searchResult = updateSearcher.Search("IsInstalled=0 and Type='Software'")
Set objWMI = GetObject("winmgmts:\\" & strlocalhost & "\root\CIMV2")
set colitems = objWMI.ExecQuery("SELECT Name FROM Win32_ComputerSystem")
	For Each objcol in colitems
		strcomputer = objcol.Name
	Next
set objtextfile = ofso.createtextfile("C:\" & strcomputer & "_patchlog.csv", True)
objtextfile.writeline "Computer" & vbTab & "Title" & vbTab & "KB" & vbTab & "IsDownloaded" & vbTab & "Status"
If searchresult.updates.count = 0 Then
	Wscript.echo "No updates to install."
	objtextfile.writeline strcomputer & vbTab & "NA" & vbTab & "NA" & vbTab & "NA" & vbTab & "NA"
	Wscript.Quit
Else
For I = 0 To searchResult.Updates.Count-1
    set update = searchResult.Updates.Item(I)
	    If update.IsDownloaded = true Then
	        updatesToInstall.Add(update)	
	    End If
Next
End If
err.clear
Wscript.Echo "Installing Updates"
Set installer = updateSession.CreateUpdateInstaller()
installer.Updates = updatesToInstall
Set installationResult = installer.Install()
	If err.number <> 0 Then
		objtextfile.writeline strcomputer & "," & update.Title & "," & err.number
	Else		
		For I = 0 to updatesToInstall.Count - 1
		objtextfile.writeline strcomputer & vbTab & updatesToInstall.Item(i).Title & vbTab & "NA" & vbTab & "NA" & vbTab & installationResult.GetUpdateResult(i).ResultCode 
		Next
	End If
Wscript.Quit
"@

    Write-Verbose "Creating vbscript file on $computername"
    $vbsstring | Out-File "\\$computername\c$\update.vbs" -Force
}


Function Format-InstallPatchLog {
    [cmdletbinding()]
    param ($computername)
    
    #Create empty collection
    $installreport = @()
    #Check for logfile
    If (Test-Path "\\$computername\c$\$($computername)_patchlog.csv") {
        #Retrieve the logfile from remote server
        $CSVreport = Import-Csv "\\$computername\c$\$($computername)_patchlog.csv" -Delimiter "`t"
        #Iterate through all items in patchlog
        ForEach ($log in $CSVreport) {
            $temp = "" | Select Computer,Title,KB,IsDownloaded,Status
            $temp.Computer = $log.Computer
            $temp.Title = $log.title.split('\(')[0]
            If ($temp.title -eq 'NA') {
                $temp.KB = "NA"
            } Else {
                $temp.KB = ($log.title.split('\(')[1]).split('\)')[0]
            }            
            $temp.IsDownloaded = "NA"
            Switch ($log.Status) {
                1 {$temp.Status = "No Reboot required"}
                2 {$temp.Status = "Reboot Required"}
                4 {$temp.Status = "Failed to Install Patch"}
                "NA" {$temp.Status = "NA"}
                Default {$temp.Status = "Unable to determine Result Code"}            
                }
            $installreport += $temp
            }
        }
    Else {
        $temp = "" | Select Computer, Title, KB,IsDownloaded,Status
        $temp.Computer = $computername
        $temp.Title = "OFFLINE"
        $temp.KB = "OFFLINE"
        $temp.IsDownloaded = "OFFLINE"
        $temp.Status = "OFFLINE"  
        $installreport += $temp      
        }
    Write-Output $installreport
}

Function Install-Patches 
{
    [cmdletbinding()]
    Param($Computername)
    
    If (Test-Path psexec.exe) {

        Write-Verbose "Creating update.vbs file on remote server."
        Create-UpdateVBS -computer $Computername
        Write-Verbose "Patching computer: $($Computername)"
        .\psexec.exe -accepteula -s -i \\$Computername cscript.exe C:\update.vbs        
        If ($LASTEXITCODE -eq 0) {
            #$host.ui.WriteLine("Successful run of install script!")
            Write-Verbose "Formatting log file and adding to report"
            Format-InstallPatchLog -computer $Computername
            }            
        Else {
            #$host.ui.WriteLine("Unsuccessful run of install script!")
            $report = "" | Select Computer,Title,KB,IsDownloaded,Status
            $report.Computer = $Computername
            $report.Title = "ERROR"
            $report.KB = "ERROR"
            $report.IsDownloaded = "ERROR"
            $report.Status = "ERROR" 
            Write-Output $report
            }
        }      
    Else {
        Write-Verbose "PSExec not in same directory as script!"            
        }
    #$host.ui.WriteLine("Exiting function")
}