<#
.SYNOPSIS
Name: PatchAutomation_Collab.ps1
The purpose of this script is to give users a convenience in terms of patching process.

.DESCRIPTION
The script is intended give patch engineers a tool with user friendly GUI that can support the things needed in line with the weekly patching activity
such as Pre and Post patching activity, running chef patch cookbook, checking services status and application existence, auditing and installing
patch, server reboots, running ad hoc scripts, etc.

.PARAMETER InitialDirectory
The initial directory which this example script will use. .\PatchAutomation

.NOTES
=================================
Date: 2023-December
What's New: 
Win2012 ESU Year 1
---------------------------------
Updated Version: 4.2
Date: 2023-August
What's New: 
Trend Micro DSA Health Checks
Trend Micro EDR Deployment
Check CHEF Environment Variables
Add CHEF Environment Variables
---------------------------------
Updated Version: 4.1
Date: 2023-04
What's New:
Release Date: 2019-08-11
Latest Patch version: 4.0
Latest Release Date: 2022-07
Author: Charles Lyndon Fontanilla
Revisions/Enhancements: 
Vincent Venzon
Raymond San Miguel
Jayson Galope
Edzel Barbosa
Aran Espiritu
=================================
#>



#region Synchronized Collections
$uiHash = [hashtable]::Synchronized(@{})
$runspaceHash = [hashtable]::Synchronized(@{})
$jobs = [system.collections.arraylist]::Synchronized((New-Object System.Collections.ArrayList))
$jobCleanup = [hashtable]::Synchronized(@{})
$Global:updateAudit = [system.collections.arraylist]::Synchronized((New-Object System.Collections.ArrayList))
$Global:prep = [system.collections.arraylist]::Synchronized((New-Object System.Collections.ArrayList))
$Global:postPatch = [system.collections.arraylist]::Synchronized((New-Object System.Collections.ArrayList))
$Global:invoke = [system.collections.arraylist]::Synchronized((New-Object System.Collections.ArrayList))
$Global:invkClstr = [system.collections.arraylist]::Synchronized((New-Object System.Collections.ArrayList))
$Global:WUinvoke = [system.collections.arraylist]::Synchronized((New-Object System.Collections.ArrayList))
$Global:adhocScript = [system.collections.arraylist]::Synchronized((New-Object System.Collections.ArrayList))
$Global:adhocBATScript = [system.collections.arraylist]::Synchronized((New-Object System.Collections.ArrayList))
$Global:ESUreport = [system.collections.arraylist]::Synchronized((New-Object System.Collections.ArrayList))
$Global:installAudit = [system.collections.arraylist]::Synchronized((New-Object System.Collections.ArrayList))
$Global:servicesAudit = [system.collections.arraylist]::Synchronized((New-Object System.Collections.ArrayList))
$Global:installedUpdates = [system.collections.arraylist]::Synchronized((New-Object System.Collections.ArrayList))
[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
#endregion

#region Startup Checks and configurations
#Determine if running from ISE

#Write-Verbose "Checking to see if running from console"
#If ($Host.name -eq "Windows PowerShell ISE Host") {
#    Write-Warning "Unable to run this from the PowerShell ISE due to issues with PSexec!`nPlease run from console."
#    Break
#} 

#Validate user is an Administrator
Write-Verbose "Checking Administrator credentials"
If (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(`
    [Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "You are not running this as an Administrator!`nRe-running script and will prompt for administrator credentials."
    Start-Process -Verb "Runas" -File PowerShell.exe -Argument "-STA -noprofile -file $($myinvocation.mycommand.definition)"
    Break
}

#Ensure that we are running the GUI from the correct location
Set-Location $(Split-Path $MyInvocation.MyCommand.Path)
$Global:Path = $(Split-Path $MyInvocation.MyCommand.Path)
Write-Debug "Current location: $Path"

#Check for PSExec
Write-Verbose "Checking for psexec.exe"
If (-Not (Test-Path psexec.exe)) {
    Write-Warning ("Psexec.exe missing from {0}!`n Please place file in the path so UI can work properly" -f (Split-Path $MyInvocation.MyCommand.Path))
    Break
}

#Determine if this instance of PowerShell can run WPF 
Write-Verbose "Checking the apartment state"
If ($host.Runspace.ApartmentState -ne "STA") {
    Write-Warning "This script must be run in PowerShell started using -STA switch!`nScript will attempt to openPowerShell in STA and run re-run script."
    Start-Process -File PowerShell.exe -Argument "-STA -noprofile -WindowStyle hidden -file $($myinvocation.mycommand.definition)"
    Break
}

#Load Required Assemblies
Add-Type –assemblyName PresentationFramework
Add-Type –assemblyName PresentationCore
Add-Type –assemblyName WindowsBase
Add-Type –assemblyName Microsoft.VisualBasic
Add-Type –assemblyName System.Windows.Forms

#Computer Cache collection
$Script:ComputerCache = New-Object System.Collections.ArrayList  

#DotSource Help script
. ".\HelpFiles\HelpOverview.ps1"

#DotSource About script
. ".\HelpFiles\About.ps1"
#endregion


Function Get-SelectedItemCount{
$listcount = $uiHash.Listview.SelectedItems.count
$uiHash.StatusTextBox.Text = "Total selected server(s): $listcount"
}

Function Set-PoshPAIGOption {
    [CmdletBinding()]
    Param ()
    # - Updated to use Environment to get Desktop location
    # - Check for valid report path on load
    # - Export-CliXML updated to use $Path instead of $pwd
    # - Simplified the setting/testing of options, either load/set defaults and then run validation


    # If the Options.xml file exists, then use it, if not then set default option values
    # Also, if the imported options are Null, then rebuild
    $Optionshash = $Null
    If (Test-Path (Join-Path $Path 'options.xml')) {
        Write-Debug "Options.xml file found"
        $Optionshash = Import-Clixml -Path (Join-Path $Path 'options.xml')
    } 
    
    If ($Optionshash -eq $null) {
        Write-Debug "Options.xml file not present. Setting default values"
        $optionshash = @{
            MaxJobs = 5
            MaxRebootJobs = 5
            ReportPath = [Environment]::GetFolderPath("Desktop")
        }
    }

    # Validate the MaxJobs Option
    If ($Optionshash['MaxJobs'])
    {
        If ([int]$Optionshash['MaxJobs'] -gt 1) {
            $Global:maxConcurrentJobs = $Optionshash['MaxJobs']
        } Else {
            $Optionshash['MaxJobs'] = $Global:maxConcurrentJobs = 5
        }
    } Else {
        $Optionshash['MaxJobs'] = $Global:maxConcurrentJobs = 5
    }

    # Validate the MaxRebootJobs Option
    If ($Optionshash['MaxRebootJobs'])
    {
        If ([int]$Optionshash['MaxRebootJobs'] -gt 1) {
            $Global:maxRebootJobs = $Optionshash['MaxRebootJobs']
        } Else {
            $Optionshash['MaxRebootJobs'] = $Global:maxRebootJobs = 5
        }
    } Else {
        $Optionshash['MaxRebootJobs'] = $Global:maxRebootJobs = 5
    }    
        
    # Validate the ReportPath Option
    If ($Optionshash['ReportPath']) {
        If (Test-Path $Optionshash['ReportPath']) {
            Write-Debug "Stored ReportPath option found and is valid"
            $Global:reportpath = $Optionshash['ReportPath']
        } Else {
            Write-Debug "Stored ReportPath option is invalid. Reverting to default"
            $Optionshash['ReportPath'] = $Global:reportpath = [Environment]::GetFolderPath("Desktop")
        }
    
    } Else {
        Write-Debug "ReportPath option not found in imported file. Reverting to default"
        $Optionshash['ReportPath'] = $Global:reportpath = [Environment]::GetFolderPath("Desktop")
    }

    # Export all options, regardless of whether they are the same as what is already in the file
    Write-Debug "Exporting options.xml"
    $optionshash | Export-Clixml -Path (Join-Path $Path 'options.xml') -Force
}

#Function for Debug output
Function Global:Show-DebugState {
    Write-Debug ("Number of Items: {0}" -f $uiHash.Listview.ItemsSource.count)
    Write-Debug ("First Item: {0}" -f $uiHash.Listview.ItemsSource[0].Computer)
    Write-Debug ("Last Item: {0}" -f $uiHash.Listview.ItemsSource[$($uiHash.Listview.ItemsSource.count) -1].Computer)
    Write-Debug ("Max Progress Bar: {0}" -f $uiHash.ProgressBar.Maximum)
}

#Reboot Warning Message
Function Show-Warning ([string]$msgType) {
    $sc = $uiHash.Listview.selecteditems.count
    $title = "Reboot Server Warning"
        if($msgType -eq "ServerReboot"){$message = "You are about to reboot $sc server(s) which can affect the environment! `nAre you sure you want to do this?"}
        elseif($msgType -eq "uninstallTBMR"){$message = "You are about to Uninstall TBMR software to $sc server(s)! `nAre you sure you want to proceed?"}
        elseif($msgType -eq "2k8ESULicense"){$message = "You are about to run 2k8 ESU License activation to $sc server(s)! `nAre you sure you want to proceed?"}
    $button = [System.Windows.Forms.MessageBoxButtons]::YesNo
    $icon = [Windows.Forms.MessageBoxIcon]::Warning
    $res = [windows.forms.messagebox]::Show($message,$title,$button,$icon)
    if($res -eq "Yes"){
        if($msgType -eq "ServerReboot"){System-Reboot}
        elseif($msgType -eq "uninstallTBMR"){Run-removeTBMR}
        #elseif($msgType -eq "2k8ESULicense"){Run-2k8ESULicense}
    }
}

#Format and display errors
Function Get-Error {
    Process {
        ForEach ($err in $error) {
            Switch ($err) {
                {$err -is [System.Management.Automation.ErrorRecord]} {
                        $hash = @{
                        Category = $err.categoryinfo.Category
                        Activity = $err.categoryinfo.Activity
                        Reason = $err.categoryinfo.Reason
                        Type = $err.GetType().ToString()
                        Exception = ($err.exception -split ": ")[1]
                        QualifiedError = $err.FullyQualifiedErrorId
                        CharacterNumber = $err.InvocationInfo.OffsetInLine
                        LineNumber = $err.InvocationInfo.ScriptLineNumber
                        Line = $err.InvocationInfo.Line
                        TargetObject = $err.TargetObject
                        }
                    }               
                Default {
                    $hash = @{
                        Category = $err.errorrecord.categoryinfo.category
                        Activity = $err.errorrecord.categoryinfo.Activity
                        Reason = $err.errorrecord.categoryinfo.Reason
                        Type = $err.GetType().ToString()
                        Exception = ($err.errorrecord.exception -split ": ")[1]
                        QualifiedError = $err.errorrecord.FullyQualifiedErrorId
                        CharacterNumber = $err.errorrecord.InvocationInfo.OffsetInLine
                        LineNumber = $err.errorrecord.InvocationInfo.ScriptLineNumber
                        Line = $err.errorrecord.InvocationInfo.Line                    
                        TargetObject = $err.errorrecord.TargetObject
                    }               
                }                        
            }
        $object = New-Object PSObject -Property $hash
        $object.PSTypeNames.Insert(0,'ErrorInformation')
        $object
        }
    }
}

#Add new server to GUI
Function Add-Server {
    $computers = [Microsoft.VisualBasic.Interaction]::InputBox("Enter a server name or names. Separate servers with a comma (,) or semi-colon (;).", "Add Server/s")
    If (-Not [System.String]::IsNullOrEmpty($computers)) {
        [string[]]$computername = $computers -split ",|;"
        ForEach ($computer in $computername) { 
            If (-NOT [System.String]::IsNullOrEmpty($computer) -AND -NOT $ComputerCache.Contains($Computer.Trim()) -AND -NOT $Exempt -contains $computer) {
                [void]$ComputerCache.Add($Computer.Trim())
                $clientObservable.Add((
                    New-Object PSObject -Property @{
                        Computer = ($computer).Trim()
                        Audited = 0 -as [int]
                        Installed = 0 -as [int]
                        InstallErrors = 0 -as [int]
                        Services = 0 -as [int]
                        Status = $Null
                        Notes = $Null
                    }
                ))     
                Show-DebugState
            }
        }
    } 
}

#Remove server from GUI
Function Remove-Server {
    $Servers = @($uiHash.Listview.SelectedItems)
    ForEach ($server in $servers) {
        $clientObservable.Remove($server)
        $ComputerCache.Remove($Server.Computer)
    }
    $uiHash.ProgressBar.Maximum = $uiHash.Listview.ItemsSource.count
    Show-DebugState  
}

#Report Generation function
Function Start-Report {
    Write-Debug ("Data: {0}" -f $uiHash.ReportComboBox.SelectedItem.Text)
    Switch ($uiHash.ReportComboBox.SelectedItem.Text) {
        "Prep CSV Report" {
            If ($prep.count -gt 0) { 
                $uiHash.StatusTextBox.Foreground = "Black"
                $savedreport = Join-Path $reportpath "PrepReport-$((Get-Date -Uformat %b%d%Y-%I%M%S).ToString()).csv"
                $prep | Sort-Object Servername | FT
                $prep | Export-Csv $savedreport -NoTypeInformation
                $uiHash.StatusTextBox.Text = "Report saved to $savedreport"
                } Else {
                $uiHash.StatusTextBox.Foreground = "Red"
                $uiHash.StatusTextBox.Text = "No report to create!"         
            }         
        }
        "PostPatch CSV Report" {
            If ($postPatch.count -gt 0) { 
                $uiHash.StatusTextBox.Foreground = "Black"
                $savedreport = Join-Path $reportpath "PostPatchReport-$((Get-Date -Uformat %b%d%Y-%I%M%S).ToString()).csv"
                $postPatch | Sort-Object Servername | FT
                $postPatch | Export-Csv $savedreport -NoTypeInformation
                $uiHash.StatusTextBox.Text = "Report saved to $savedreport"
                } Else {
                $uiHash.StatusTextBox.Foreground = "Red"
                $uiHash.StatusTextBox.Text = "No report to create!"         
            }         
        }
        "Audit CSV Report" {
            If ($updateAudit.count -gt 0) { 
                $uiHash.StatusTextBox.Foreground = "Black"
                $savedreport = Join-Path $reportpath "AuditReport-$((Get-Date -Uformat %b%d%Y-%I%M%S).ToString()).csv"
                $updateAudit | Export-Csv $savedreport -NoTypeInformation
                $uiHash.StatusTextBox.Text = "Report saved to $savedreport"
                } Else {
                $uiHash.StatusTextBox.Foreground = "Red"
                $uiHash.StatusTextBox.Text = "No report to create!"         
            }         
        }
        "Prep UI Report" {
            If ($prep.count -gt 0) {
                $prep | Out-GridView -Title 'Prep Report'
            } Else {
                $uiHash.StatusTextBox.Foreground = "Red"
                $uiHash.StatusTextBox.Text = "No report to create!"         
            }
        }
        "PostPatch UI Report" {
            If ($postPatch.count -gt 0) {
                $postPatch | Out-GridView -Title 'Post-Patch Report'
            } Else {
                $uiHash.StatusTextBox.Foreground = "Red"
                $uiHash.StatusTextBox.Text = "No report to create!"         
            }
        }
        "Audit UI Report" {
            If ($updateAudit.count -gt 0) {
                $updateAudit | Out-GridView -Title 'Audit Report'
            } Else {
                $uiHash.StatusTextBox.Foreground = "Red"
                $uiHash.StatusTextBox.Text = "No report to create!"         
            }
        }
        "Install CSV Report" {
            If ($installAudit.count -gt 0) { 
                $uiHash.StatusTextBox.Foreground = "Black"
                $savedreport = Join-Path $reportpath "InstallReport-$((Get-Date -Uformat %b%d%Y-%I%M%S).ToString()).csv"
                $installAudit | Export-Csv $savedreport -NoTypeInformation
                $uiHash.StatusTextBox.Text = "Report saved to $savedreport"
                } Else {
                $uiHash.StatusTextBox.Foreground = "Red"
                $uiHash.StatusTextBox.Text = "No report to create!"         
            }        
        }
        "Install UI Report" {
            If ($installAudit.count -gt 0) {
                $installAudit | Out-GridView -Title 'Install Report'
            } Else {
                $uiHash.StatusTextBox.Foreground = "Red"
                $uiHash.StatusTextBox.Text = "No report to create!"         
            }        
        }
        "Installed Updates CSV Report" {
            If ($installedUpdates.count -gt 0) { 
                $uiHash.StatusTextBox.Foreground = "Black"
                $savedreport = Join-Path $reportpath "InstalledUpdatesReport-$((Get-Date -Uformat %b%d%Y-%I%M%S).ToString()).csv"
                $installedUpdates | Export-Csv $savedreport -NoTypeInformation
                $uiHash.StatusTextBox.Text = "Report saved to $savedreport"
                } Else {
                $uiHash.StatusTextBox.Foreground = "Red"
                $uiHash.StatusTextBox.Text = "No report to create!"         
            }        
        }
        "Installed Updates UI Report" {
            If ($installedUpdates.count -gt 0) {
                $installedUpdates | Out-GridView -Title 'Installed Updates Report'
            } Else {
                $uiHash.StatusTextBox.Foreground = "Red"
                $uiHash.StatusTextBox.Text = "No report to create!"         
            }        
        }
        "Host File List" {
            If ($uiHash.Listview.Items.count -gt 0) { 
                $uiHash.StatusTextBox.Foreground = "Black"
                $savedreport = Join-Path $reportpath "hosts-$((Get-Date -Uformat %b%d%Y-%I%M%S).ToString()).txt"
                $uiHash.Listview.items | Select -Expand Computer | Out-File $savedreport
                $uiHash.StatusTextBox.Text = "Report saved to $savedreport"
                } Else {
                $uiHash.StatusTextBox.Foreground = "Red"
                $uiHash.StatusTextBox.Text = "No report to create!"         
            }        
        }
        "Computer List Report" {
            If ($uiHash.Listview.Items.count -gt 0) {
                $uiHash.StatusTextBox.Foreground = "Black"
                $savedreport = Join-Path $Global:ReportPath "serverlist-$((Get-Date -Uformat %b%d%Y-%I%M%S).ToString()).csv"
                $uiHash.Listview.Items | Export-Csv -NoTypeInformation $savedreport
                $uiHash.StatusTextBox.Text = "Report saved to $savedreport"        
            } Else {
                $uiHash.StatusTextBox.Foreground = "Red"
                $uiHash.StatusTextBox.Text = "No report to create!"         
            }         
        }
        "Error UI Report" {Get-Error | Out-GridView -Title 'Error Report'}
        "Services UI Report" {
            If (@($servicesAudit).count -gt 0) {
                $servicesAudit | Select @{L='Computername';E={$_.__Server}},Name,DisplayName,State,StartMode,ExitCode,Status | Out-GridView -Title 'Services Report'
            } Else {
                $uiHash.StatusTextBox.Foreground = "Red"
                $uiHash.StatusTextBox.Text = "No report to create!"             
            }
        }
        "Services CSV Report" {
            If (@($servicesAudit).count -gt 0) { 
                $uiHash.StatusTextBox.Foreground = "Black"
                $savedreport = Join-Path $reportpath "ServicesReport-$((Get-Date -Uformat %b%d%Y-%I%M%S).ToString()).csv"
                $servicesAudit | Select @{L='Computername';E={$_.__Server}},Name,DisplayName,State,StartMode,ExitCode,Status | Export-Csv $savedreport -NoTypeInformation
                $uiHash.StatusTextBox.Text = "Report saved to $savedreport"
                } Else {
                $uiHash.StatusTextBox.Foreground = "Red"
                $uiHash.StatusTextBox.Text = "No report to create!"         
            }         
        }
        "ESU License UI Report" {
            If ($Esureport.count -gt 0) {
                $Esureport | Out-GridView -Title 'ESU License Report'
            } Else {
                $uiHash.StatusTextBox.Foreground = "Red"
                $uiHash.StatusTextBox.Text = "No report to create!"         
            }
        }
        
        "ESU License CSV Report" {
            If ($Esureport.count -gt 0) { 
                $uiHash.StatusTextBox.Foreground = "Black"
                $savedreport = Join-Path $reportpath "ESUreport-$((Get-Date -Uformat %b%d%Y-%I%M%S).ToString()).csv"
                $Esureport | Sort-Object Servername | FT
                $Esureport | Export-Csv $savedreport -NoTypeInformation
                $uiHash.StatusTextBox.Text = "Report saved to $savedreport"
                } Else {
                $uiHash.StatusTextBox.Foreground = "Red"
                $uiHash.StatusTextBox.Text = "No report to create!"         
            }         
        }
    }
}

#Run Pre-patching activity
Function RunPrePatch{
    Write-Debug ("ComboBox {0}" -f $uiHash.RunOptionComboBox.Text)
    $selectedItems = $uiHash.Listview.SelectedItems
    If ($selectedItems.Count -gt 0) {
        $uiHash.ProgressBar.Maximum = $selectedItems.count
        $uiHash.Listview.ItemsSource | ForEach {$_.Status = $Null; $_.notes = $Null} 
            $uiHash.RunButton.IsEnabled = $False
            $uiHash.StartImage.Source = "$pwd\Images\Start_locked.jpg"
            $uiHash.CancelButton.IsEnabled = $True
            $uiHash.CancelImage.Source = "$pwd\Images\Stop.jpg"            
            $uiHash.StatusTextBox.Foreground = "Black"
            $uiHash.StatusTextBox.Text = "Generating pre-patch activity...Please Wait"            
            $Global:updatelayout = [Windows.Input.InputEventHandler]{ $uiHash.ProgressBar.UpdateLayout() }
            $uiHash.StartTime = (Get-Date)
            
            [Float]$uiHash.ProgressBar.Value = 0
            $scriptBlock = {
                Param (
                    $Path,
                    $Computer,
                    $prep,
                    $uiHash
                )
                $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                    $uiHash.Listview.Items.EditItem($Computer)
                    $computer.Status = "Running Pre-Patch report"
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh() 
                })  
                Set-Location $path
                . .\Scripts\Get-PrePatch.ps1                
                $prepRes = @(Get-PrePatch -Computer $computer.computer)

                $prep.AddRange($prepRes) | Out-Null

                $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{
                    $uiHash.Listview.Items.EditItem($Computer)
                    If ($prepRes[0].Error -eq 0) {
                        $Computer.Status = "Completed"
                    }ElseIf ($prepRes[0].Error -eq "Offline") {
                        $Computer.Status = "Offline"
                    } 
                    Else {
                        $errcnt = $prepRes[0].Error
                        $Computer.Status = "Failed"
                        $computer.Notes = "There were $errcnt error(s) encountered!"
                    }
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh()
                })
                $uiHash.ProgressBar.Dispatcher.Invoke("Normal",[action]{
                    $uiHash.ProgressBar.value++ 
                }) 
                    
                $uiHash.Window.Dispatcher.Invoke("Normal",[action]{
                    #Check to see if find job
                    If ($uiHash.ProgressBar.value -eq $uiHash.ProgressBar.Maximum) {    
                        $End = New-Timespan $uihash.StartTime (Get-Date)                                             
                        $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)                                     
                        $uiHash.RunButton.IsEnabled = $True
                        $uiHash.StartImage.Source = "$pwd\Images\Start.jpg"
                        $uiHash.CancelButton.IsEnabled = $False
                        $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.jpg"                    
                    }
                })  
                
            }

            Write-Verbose ("Creating runspace pool and session states")
            $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
            $runspaceHash.runspacepool = [runspacefactory]::CreateRunspacePool(1, $maxConcurrentJobs, $sessionstate, $Host)
            $runspaceHash.runspacepool.Open()    

            ForEach ($Computer in $selectedItems) {
                $uiHash.Listview.Items.EditItem($Computer)
                $computer.Status = "Pending Pre-patch Activity"
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh() 
                #Create the powershell instance and supply the scriptblock with the other parameters 
                $powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($Path).AddArgument($computer).AddArgument($prep).AddArgument($uiHash)
           
                #Add the runspace into the powershell instance
                $powershell.RunspacePool = $runspaceHash.runspacepool
           
                #Create a temporary collection for each runspace
                $temp = "" | Select-Object PowerShell,Runspace,Computer
                $Temp.Computer = $Computer.computer
                $temp.PowerShell = $powershell
           
                #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
                $temp.Runspace = $powershell.BeginInvoke()
                Write-Verbose ("Adding {0} collection" -f $temp.Computer)
                $jobs.Add($temp) | Out-Null  
                }
               } 
             Else {
                $uiHash.StatusTextBox.Foreground = "Red"
                $uiHash.StatusTextBox.Text = "No server/s selected!"
        }  
}

#Run Post-patching activity
Function RunPostPatch{
    Write-Debug ("ComboBox {0}" -f $uiHash.RunOptionComboBox.Text)
    $selectedItems = $uiHash.Listview.SelectedItems
    If ($selectedItems.Count -gt 0) {
        $uiHash.ProgressBar.Maximum = $selectedItems.count
        $uiHash.Listview.ItemsSource | ForEach {$_.Status = $Null; $_.notes = $Null} 
            $uiHash.RunButton.IsEnabled = $False
            $uiHash.StartImage.Source = "$pwd\Images\Start_locked.jpg"
            $uiHash.CancelButton.IsEnabled = $True
            $uiHash.CancelImage.Source = "$pwd\Images\Stop.jpg"            
            $uiHash.StatusTextBox.Foreground = "Black"
            $uiHash.StatusTextBox.Text = "Generating post-patch activity...Please Wait"            
            $Global:updatelayout = [Windows.Input.InputEventHandler]{ $uiHash.ProgressBar.UpdateLayout() }
            $uiHash.StartTime = (Get-Date)
            
            [Float]$uiHash.ProgressBar.Value = 0
            $scriptBlock = {
                Param (
                    $Path,
                    $Computer,
                    $postPatch,
                    $uiHash
                )
                $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                    $uiHash.Listview.Items.EditItem($Computer)
                    $computer.Status = "Running Post-Patch report"
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh() 
                })  
                Set-Location $path
                . .\Scripts\Get-PostPatch.ps1                
                $postPatchRes = @(Get-PostPatch -Computer $computer.computer)

                $postPatch.AddRange($postPatchRes) | Out-Null

                $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{
                    $uiHash.Listview.Items.EditItem($Computer)
                    If ($postPatchRes[0].Error -eq 0) { 
                        $Computer.Status = "Completed"
                    }ElseIf ($postPatchRes[0].Error -eq "Offline") {
                        $Computer.Status = "Offline"
                    } 
                    Else {
                        $errcnt = $postPatchRes[0].Error
                        $Computer.Status = "Failed"
                        $computer.Notes = "There were $errcnt error(s) encountered!"
                    }
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh()
                })
                $uiHash.ProgressBar.Dispatcher.Invoke("Normal",[action]{
                    $uiHash.ProgressBar.value++ 
                }) 
                    
                $uiHash.Window.Dispatcher.Invoke("Normal",[action]{
                    #Check to see if find job
                    If ($uiHash.ProgressBar.value -eq $uiHash.ProgressBar.Maximum) {    
                        $End = New-Timespan $uihash.StartTime (Get-Date)                                             
                        $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)                                     
                        $uiHash.RunButton.IsEnabled = $True
                        $uiHash.StartImage.Source = "$pwd\Images\Start.jpg"
                        $uiHash.CancelButton.IsEnabled = $False
                        $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.jpg"                    
                    }
                })  
                
            }

            Write-Verbose ("Creating runspace pool and session states")
            $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
            $runspaceHash.runspacepool = [runspacefactory]::CreateRunspacePool(1, $maxConcurrentJobs, $sessionstate, $Host)
            $runspaceHash.runspacepool.Open()    

            ForEach ($Computer in $selectedItems) {
                $uiHash.Listview.Items.EditItem($Computer)
                $computer.Status = "Pending Post-patch Activity"
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh() 
                #Create the powershell instance and supply the scriptblock with the other parameters 
                $powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($Path).AddArgument($computer).AddArgument($postPatch).AddArgument($uiHash)
           
                #Add the runspace into the powershell instance
                $powershell.RunspacePool = $runspaceHash.runspacepool
           
                #Create a temporary collection for each runspace
                $temp = "" | Select-Object PowerShell,Runspace,Computer
                $Temp.Computer = $Computer.computer
                $temp.PowerShell = $powershell
           
                #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
                $temp.Runspace = $powershell.BeginInvoke()
                Write-Verbose ("Adding {0} collection" -f $temp.Computer)
                $jobs.Add($temp) | Out-Null  
                }
               } 
             Else {
                $uiHash.StatusTextBox.Foreground = "Red"
                $uiHash.StatusTextBox.Text = "No server/s selected!"
        }  
}

#Message prompt for CHEF run actions
Function Show-ChefRunWarning ([string]$chefstring) {
    $sc = $uiHash.Listview.selecteditems.count
    $title = "Running $chefstring Warning"
    $message = "You are about to run $chefstring to $sc server(s)! `nAre you sure you want to proceed?"
    $button = [System.Windows.Forms.MessageBoxButtons]::YesNo
    $icon = [Windows.Forms.MessageBoxIcon]::Warning
    $msgbox = [windows.forms.messagebox]::Show($message,$title,$button,$icon)
    if($chefstring -eq "Windows Update"){
        if ($msgbox -eq "Yes") {WU-run}
    }else{
        if ($msgbox -eq "Yes") {Chef-run}
    }
}

#Run Windows Update
Function WU-run {
 Write-Debug ("ComboBox {0}" -f $uiHash.RunOptionComboBox.Text)
    $selectedItems = $uiHash.Listview.SelectedItems
    If ($selectedItems.Count -gt 0) {
        $uiHash.ProgressBar.Maximum = $selectedItems.count
        $uiHash.Listview.ItemsSource | ForEach {$_.Status = $Null; $_.notes = $Null} 
#region WU
            $uiHash.RunButton.IsEnabled = $False
            $uiHash.StartImage.Source = "$pwd\Images\Start_locked.jpg"
            $uiHash.CancelButton.IsEnabled = $True
            $uiHash.CancelImage.Source = "$pwd\Images\Stop.jpg"            
            $uiHash.StatusTextBox.Foreground = "Black"
            $uiHash.StatusTextBox.Text = "Windows Update is being run..."            
            $uiHash.StartTime = (Get-Date)
            [Float]$uiHash.ProgressBar.Value = 0
            $scriptBlock = {
                Param (
                    $Computer,
                    $uiHash,
                    $Path,
                    $WUinvoke
                    
                )               
                $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                    $uiHash.Listview.Items.EditItem($Computer)
                    $computer.Status = "Initiating"
                    $computer.Notes = "Configuring pre-requisites..."
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh() 
                })                
                Set-Location $Path
                $Connection = (Test-Connection -ComputerName $Computer.computer -Count 1 -Quiet)
                $server = $Computer.computer
                winrm set winrm/config/client ‘@{TrustedHosts="'$server'"}’
                $Destination = "\\" + $server + "\C$\Windows\System32\WindowsPowerShell\v1.0\Modules\PSWindowsUpdate"
                $testPath = try{Invoke-Command -ComputerName $server -ScriptBlock{
                            test-path -path "c:\Windows\System32\WindowsPowerShell\v1.0\Modules\PSWindowsUpdate\*"}-ErrorVariable e -ErrorAction SilentlyContinue }
                            Catch{return "Failed"}
                #$testPath = Test-Path -Path $Destination
                if($testPath -eq $False){
                    try{
                      $sourcePath = $Path + "\Module\pswindowsupdate"
                      robocopy "$sourcePath" $Destination /S /Z /R:1 /NFL /NDL /NJH /NJS /nc /ns /np
                      Invoke-Command -ComputerName $server -ScriptBlock{
                        Import-Module -Name C:\Windows\System32\WindowsPowerShell\v1.0\Modules\PSWindowsUpdate\PSWindowsUpdate.psm1
                        Import-Module -Name C:\Windows\System32\WindowsPowerShell\v1.0\Modules\PSWindowsUpdate\PSWindowsUpdate.psd1} -ErrorAction SilentlyContinue
                      $isTrue = $True
                      $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                      $uiHash.Listview.Items.EditItem($Computer)
                      $computer.Status = "Completed"
                      $computer.Notes = "Pre-requisite configuration Completed"
                      $uiHash.Listview.Items.CommitEdit()
                      $uiHash.Listview.Items.Refresh()})
                    }Catch{
                      $getStatus = "Failed"
                      $getNotes = $error[0]
                      $ua = 0
                      $isTrue = $false
                    }
                 }elseif(-not($e[0] -eq $null)){
                       $getStatus = "Failed"
                       $getNotes = $e[0] #"Invoke is not working, might be a WinRM issue! Please patch manually!"
                       $ua = 0
                       $isTrue = $false
                 }else{$uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                       $uiHash.Listview.Items.EditItem($Computer)
                       $computer.Status = "Completed"
                       $computer.Notes = "Pre-requisite is already configured"
                       $uiHash.Listview.Items.CommitEdit()
                       $uiHash.Listview.Items.Refresh()})
                       $isTrue = $True}
              if($isTrue){
              Switch($WUinvoke){
              "WU"    {Start-Sleep -s 2
                       $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                       $uiHash.Listview.Items.EditItem($Computer)
                       $computer.Status = "Running"
                       $computer.Notes = "Checking available updates..."
                       $uiHash.Listview.Items.CommitEdit()
                       $uiHash.Listview.Items.Refresh()})
                       $uac,$err = invoke-command -ComputerName $server -scriptblock {
                                   try {Get-wulist}
                                   Catch {$err = $($error[0])
                                            return "Failed", $err}
                              }
                              Write-host $err
                       if($uac -eq "Failed"){$getStatus = "Failed"; $getNotes = "Unable to get Available Updates, please check manually!"; $ua = 0}
                       else{
                       $ua = ($uac.kb).count
                       Do{
                          $connectiontimeout = 0
                          $updatetimeout = 0
                          $session = New-PSSession -ComputerName $server
                          if($session.state -match "Opened"){
                                $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                                $uiHash.Listview.Items.EditItem($Computer)
                                $computer.Status = "Running"
                                $computer.Notes = "Checking available updates..."
                                $uiHash.Listview.Items.CommitEdit()
                                $uiHash.Listview.Items.Refresh()})
                              $updates = invoke-command -session $session -scriptblock {Get-wulist}
                              $updatenumber = ($updates.kb).count
                                  if ($updates -ne $null){
                                      $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                                      $uiHash.Listview.Items.EditItem($Computer)
                                      $computer.Audited = $ua
                                      $computer.Notes = "Patch installation in progress..."
                                      $uiHash.Listview.Items.CommitEdit()
                                      $uiHash.Listview.Items.Refresh()})
                                      try{
                                          invoke-command -ComputerName $server -ScriptBlock {Invoke-WUjob -ComputerName localhost -Script "ipmo PSWindowsUpdate; Install-WindowsUpdate -AcceptAll | Out-File C:\temp\PSWU.log" -Confirm:$false -RunNow}
                                          do {$updatestatus = Get-Content \\$server\c$\temp\PSWU.log
                                              #Get-Content \\$server\c$\temp\PSWU.log | select-object -last 1
                                              sleep -Seconds 5
                                              $ErrorActionPreference = ‘SilentlyContinue’
                                              $installednumber = ([regex]::Matches($updatestatus, "Installed" )).count
                                              $Failednumber = ([regex]::Matches($updatestatus, "Failed" )).count
                                              $ErrorActionPreference = ‘Continue’
                                                $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                                                $uiHash.Listview.Items.EditItem($Computer)
                                                $computer.Audited = $updatenumber
                                                $computer.Installed = $installednumber
                                                $computer.InstallErrors = $Failednumber
                                                $uiHash.Listview.Items.CommitEdit()
                                                $uiHash.Listview.Items.Refresh()})
                                              $updatetimeout++
                                          }until (($installednumber + $Failednumber) -eq $updatenumber -or $updatetimeout -ge 1440)
                                          if(-not($updatetimeout -ge 1440)){
                                          $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                                          $uiHash.Listview.Items.EditItem($Computer)
                                          $computer.Status = "Rebooting"
                                          $computer.Notes = "Reboot in progress..."
                                          $uiHash.Listview.Items.CommitEdit()
                                          $uiHash.Listview.Items.Refresh()})
                                          Restart-Computer -Wait -ComputerName $server -Force
                                          do{#$session = New-PSSession -ComputerName $server
                                            sleep -seconds 5
                                            $connectiontimeout++
                                          }until ($session.state -match "Opened" -or $connectiontimeout -ge 60)
                                          }$uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                                            $uiHash.Listview.Items.EditItem($Computer)
                                            $computer.Status = "Rebooting"
                                            $computer.Notes = "Server comes back from reboot"
                                            $uiHash.Listview.Items.CommitEdit()
                                            $uiHash.Listview.Items.Refresh()})
                                      }Catch{$getStatus = "Failed";$getNotes = $($error[0])}
                                  }
                          }elseif($connectiontimeout -ge 60){$Connection -eq $false; $updates -eq $null; $Computer.Notes = "Server is did not come back after restart!"}
                       }until($updates -eq $null -or $Failednumber -gt 0 -or $updatetimeout -ge 1440)
                    if(-not($updatetimeout -ge 1440)){
                        if(-not($ua -eq 0)){
                          if($Failednumber -gt 0){$getStatus = "Failed";$getNotes = "Total Update(s) =  $ua : Installed =  $installednumber : Failed =  $Failednumber"
                          }else{$getStatus = "Completed";$getNotes = "Patch Completed"}
                        }Else{$getStatus = "Completed"; $getNotes = "No Updates Available"}
                    }else{$getStatus = "Failed";$getNotes="Patch Installation already reached 2 hrs threshold, please check for possible issue(s)."}
                 invoke-command -computername $server -ScriptBlock {Unregister-ScheduledTask -TaskName PSWindowsUpdate -Confirm:$false} -ErrorAction Continue
                 $date = Get-Date -Format "MM-dd-yyyy_hh-mm-ss"
                 if(Test-Path \\$server\c$\temp\PSWU.log){Rename-Item \\$server\c$\temp\PSWU.log -NewName "WindowsUpdate-$date.log"}}
              }
              "NR"   {  Start-Sleep -s 2
                        $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                        $uiHash.Listview.Items.EditItem($Computer)
                        $computer.Status = "Running"
                        $computer.Notes = "Checking available updates..."
                        $uiHash.Listview.Items.CommitEdit()
                        $uiHash.Listview.Items.Refresh()})
                      #$session = New-PSSession -ComputerName $server
                      $updates = invoke-command -ComputerName $server -scriptblock {
                                    try {Get-wulist}
                                    Catch {return "Failed"}
                                 }
                      $updatenumber = ($updates.kb).count
                      if ($updates -ne $null){
                        if($updates -ne "Failed"){
                            $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                            $uiHash.Listview.Items.EditItem($Computer)
                            $computer.Audited = $updatenumber
                            $computer.Notes = "Patch installation in progress..."
                            $uiHash.Listview.Items.CommitEdit()
                            $uiHash.Listview.Items.Refresh()})
                            try{
                            invoke-command -ComputerName $server -ScriptBlock{Invoke-WUjob -ComputerName localhost -Script "ipmo PSWindowsUpdate; Install-WindowsUpdate -AcceptAll -IgnoreReboot | Out-File C:\temp\PSWU.log" -Confirm:$false -RunNow}
                              do {
                                  Start-Sleep -s 5
							      $updatestatus = Get-Content \\$server\c$\temp\PSWU.log
                                  #Get-Content \\$server\c$\temp\PSWU.log | select-object -last 1
                                  $ErrorActionPreference = ‘SilentlyContinue’
                                  $installednumber = ([regex]::Matches($updatestatus, "Installed" )).count
                                  $Failednumber = ([regex]::Matches($updatestatus, "Failed" )).count
                                  $ErrorActionPreference = ‘Continue’
                                    $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                                    $uiHash.Listview.Items.EditItem($Computer)
                                    $computer.Audited = $updatenumber
                                    $computer.Installed = $installednumber
                                    $computer.InstallErrors = $Failednumber
                                    $uiHash.Listview.Items.CommitEdit()
                                    $uiHash.Listview.Items.Refresh()})
                                  $updatetimeout++
							    }Until(($installednumber + $Failednumber) -eq $updatenumber -or $updatetimeout -ge 1440)
                            }Catch{$getStatus = "Failed";$getNotes = $($error[0])}
                            if(-not($updatetimeout -ge 1440)){
                              if($Failednumber -eq 0){$getStatus = "Completed";$getNotes = "Patch installation Complete! Server needs to be rebooted!";$ua = $updatenumber}
                              else{$getStatus = "Failed";$getNotes = "There is/are failed patch(es) installation, Please check manually!";$ua = $updatenumber}
						    }else{$getStatus = "Failed";$getNotes = "Already reached Patch installation 2hrs threshold, Please check manually!";$ua = $updatenumber}
                        }Eslse{$getStatus = "Failed";$getNotes = "Unable to get Available Updates, please check manually!";$ua = 0}
                      }Else{$getStatus = "Completed";$getNotes = "No Updates Availble";$ua = 0}
                     invoke-command -computername $server -ScriptBlock {Unregister-ScheduledTask -TaskName PSWindowsUpdate -Confirm:$false} -ErrorAction Continue
                     $date = Get-Date -Format "MM-dd-yyyy_hh-mm-ss"
                     if(Test-Path \\$server\c$\temp\PSWU.log){Rename-Item \\$server\c$\temp\PSWU.log -NewName "WindowsUpdate-$date.log"}
                      
                     }
              "SR"   {$getStatus = "Failed"; $getNotes = "This feature is under development."}
              Default{$null}
              }
              }
                       
                $uiHash.ListView.Dispatcher.Invoke("Background",[action]{                    
                    $uiHash.Listview.Items.EditItem($Computer)
                    $Computer.Notes = $null
                    If ($Connection) {
                        $Computer.Status = $getStatus
                        $Computer.Notes = $getNotes
                        if(-not($ua -eq 0)){$computer.InstallErrors = $Failednumber; $computer.Audited = $ua; $computer.Installed = $installednumber}
                        else{$computer.InstallErrors = 0;$computer.Audited = 0;$computer.Installed = 0}
                    } ElseIf (-Not $Connection) {
                        $Computer.Status = "Offline"
                    } Else {
                        $Computer.Status = "Unknown"
                    } 
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh()
                })
                $uiHash.ProgressBar.Dispatcher.Invoke("Normal",[action]{   
                    $uiHash.ProgressBar.value++  
                })
                $uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
                    #Check to see if find job
                    If ($uiHash.ProgressBar.value -eq $uiHash.ProgressBar.Maximum) { 
                        $End = New-Timespan $uihash.StartTime (Get-Date)                                             
                        $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)
                        $uiHash.RunButton.IsEnabled = $True
                        $uiHash.StartImage.Source = "$pwd\Images\Start.jpg"
                        $uiHash.CancelButton.IsEnabled = $False
                        $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.jpg"                    
                    }
                })  
            }
            Write-Verbose ("Creating runspace pool and session states")
            $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
            $runspaceHash.runspacepool = [runspacefactory]::CreateRunspacePool(1, $maxConcurrentJobs, $sessionstate, $Host)
            $runspaceHash.runspacepool.Open()   

            ForEach ($Computer in $selectedItems) {
                $uiHash.Listview.Items.EditItem($Computer)
                $computer.Status = "Pending Windows Update instance..."
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh() 
                #Create the powershell instance and supply the scriptblock with the other parameters 
                $powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($computer).AddArgument($uiHash).AddArgument($Path).AddArgument($WUinvoke)
           
                #Add the runspace into the powershell instance
                $powershell.RunspacePool = $runspaceHash.runspacepool
           
                #Create a temporary collection for each runspace
                $temp = "" | Select-Object PowerShell,Runspace,Computer
                $Temp.Computer = $Computer.computer
                $temp.PowerShell = $powershell
           
                #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
                $temp.Runspace = $powershell.BeginInvoke()
                Write-Verbose ("Adding {0} collection" -f $temp.Computer)
                $jobs.Add($temp) | Out-Null                
            }
            #endregion
            }Else {
                $uiHash.StatusTextBox.Foreground = "Red"
                $uiHash.StatusTextBox.Text = "No server/s selected!"
            }
}

#Run CHEF patch_now; with attributes
Function Chef-run {
 Write-Debug ("ComboBox {0}" -f $uiHash.RunOptionComboBox.Text)
    $selectedItems = $uiHash.Listview.SelectedItems
    If ($selectedItems.Count -gt 0) {
        $uiHash.ProgressBar.Maximum = $selectedItems.count
        $uiHash.Listview.ItemsSource | ForEach {$_.Status = $Null; $_.notes = $Null} 
#region Chef-invoke
            $uiHash.RunButton.IsEnabled = $False
            $uiHash.StartImage.Source = "$pwd\Images\Start_locked.jpg"
            $uiHash.CancelButton.IsEnabled = $True
            $uiHash.CancelImage.Source = "$pwd\Images\Stop.jpg"            
            $uiHash.StatusTextBox.Foreground = "Black"
            $uiHash.StatusTextBox.Text = "CHEF invoke is being run..."            
            $uiHash.StartTime = (Get-Date)
            [Float]$uiHash.ProgressBar.Value = 0
            $scriptBlock = {
                Param (
                    $Computer,
                    $uiHash,
                    $Path,
                    $invoke
                )               
                $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                    $uiHash.Listview.Items.EditItem($Computer)
                    $computer.Status = "Invoking, Please wait..."
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh() 
                })                
                Set-Location $Path
                $Connection = (Test-Connection -ComputerName $Computer.computer -Count 1 -Quiet)
                        $server = $Computer.computer  
                                            #$ReaderMUIname = (Get-WmiObject -Class win32_product -ComputerName $server -Filter "name LIKE 'Adobe%Reader%'")."name"
                                            #$ReaderMUIversion = (Get-WmiObject -Class win32_product -ComputerName $server -Filter "name LIKE 'Adobe%Reader%'")."version"
                                            #if ($ReaderMUIversion){
                                            #$resout = "$ReaderMUIname Version: $ReaderMUIversion"}
                                            #Else{
                                            #$resout = "No software installed!"}
                          Switch ($invoke) {
                            "Chefinvoke" {$chefinvoke = Invoke-Command -ComputerName $server -ScriptBlock {
                                                        c:\Users\Administrator\scripts\patch_now.ps1 -q
                                                        start-sleep -s 5
                                                        get-content -path "c:\users\administrator\patch_trigger.txt"}
                                              if ($?){
                                              $getStatus = "Completed"
                                              $getNotes = $chefinvoke
                                              }Else{
                                              $getStatus = "$invoke Failed $($chefinvoke.exitcode)"}
                                          
                            }
                            "ChefinvokeNR" {$chefinvoke = Invoke-Command -ComputerName $server -ScriptBlock {
                                                        c:\Users\Administrator\scripts\patch_now.ps1 -n -q
                                                        start-sleep -s 5
                                                        get-content -path "c:\users\administrator\patch_trigger.txt"}
                                              if ($?){
                                              $getStatus = "Completed"
                                              $getNotes = $chefinvoke
                                              }Else{
                                              $getStatus = "$invoke Failed $($chefinvoke.exitcode)"}
                            }
                            "ChefinvokeAbort" {$chefinvoke = Invoke-Command -ComputerName $server -ScriptBlock {
                                                        c:\Users\Administrator\scripts\patch_now.ps1 -a -q 
                                                        start-sleep -s 5
                                                        get-content -path "c:\users\administrator\patch_trigger.txt"}
                                              if ($?){
                                              $getStatus = "Completed"
                                              $getNotes = $chefinvoke
                                              }Else{
                                              $getStatus = "$invoke Failed $($chefinvoke.exitcode)"}
                             }
                            Default {$Null}
                            }
                     
                $uiHash.ListView.Dispatcher.Invoke("Background",[action]{                    
                    $uiHash.Listview.Items.EditItem($Computer)
                    If ($Connection) {
                        $Computer.Status = $getStatus
                        $Computer.Notes = $getNotes
                    } ElseIf (-Not $Connection) {
                        $Computer.Status = "Offline"
                    } Else {
                        $Computer.Status = "Unknown"
                    } 
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh()
                })
                $uiHash.ProgressBar.Dispatcher.Invoke("Normal",[action]{   
                    $uiHash.ProgressBar.value++  
                })
                $uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
                    #Check to see if find job
                    If ($uiHash.ProgressBar.value -eq $uiHash.ProgressBar.Maximum) { 
                        $End = New-Timespan $uihash.StartTime (Get-Date)                                             
                        $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)
                        $uiHash.RunButton.IsEnabled = $True
                        $uiHash.StartImage.Source = "$pwd\Images\Start.jpg"
                        $uiHash.CancelButton.IsEnabled = $False
                        $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.jpg"                    
                    }
                })  
            }
            Write-Verbose ("Creating runspace pool and session states")
            $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
            $runspaceHash.runspacepool = [runspacefactory]::CreateRunspacePool(1, $maxConcurrentJobs, $sessionstate, $Host)
            $runspaceHash.runspacepool.Open()   

            ForEach ($Computer in $selectedItems) {
                $uiHash.Listview.Items.EditItem($Computer)
                $computer.Status = "Pending CHEF Invoke instance..."
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh() 
                #Create the powershell instance and supply the scriptblock with the other parameters 
                $powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($computer).AddArgument($uiHash).AddArgument($Path).AddArgument($invoke)
           
                #Add the runspace into the powershell instance
                $powershell.RunspacePool = $runspaceHash.runspacepool
           
                #Create a temporary collection for each runspace
                $temp = "" | Select-Object PowerShell,Runspace,Computer
                $Temp.Computer = $Computer.computer
                $temp.PowerShell = $powershell
           
                #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
                $temp.Runspace = $powershell.BeginInvoke()
                Write-Verbose ("Adding {0} collection" -f $temp.Computer)
                $jobs.Add($temp) | Out-Null                
            }
            #endregion
            }Else {
                $uiHash.StatusTextBox.Foreground = "Red"
                $uiHash.StatusTextBox.Text = "No server/s selected!"
            }
}

#Inputbox for Adhoc run
Function ShowInputbox ([string]$adhocType) {
#region

$global:x

# Create an object to hold the pop-up dialog box.
$objForm = New-Object System.Windows.Forms.Form
$objForm.Text = "Ad hoc script"
$objForm.Size = New-Object System.Drawing.Size(480, 320)
$objForm.StartPosition = "CenterScreen"

# Create a label for the pop-up menu then attach it to the Textbox List.
$objLabel = New-Object System.Windows.Forms.Label
$objLabel.Location = New-Object System.Drawing.Size(10, 20)
$objLabel.Size = New-Object System.Drawing.Size(280, 20)
$objLabel.Text = "Input script below:"
$objForm.Controls.Add($objLabel)

#Create textbox with multiline.
$objTextBox1 = New-Object System.Windows.Forms.TextBox 
$objTextBox1.Multiline = $True;
$objTextBox1.Location = New-Object System.Drawing.Size(10, 40) 
$objTextBox1.Size = New-Object System.Drawing.Size(450,200)
$objTextBox1.Scrollbars = 3
$objForm.Controls.Add($objTextBox1)

# Create an OK button.
$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Size(300, 250)
$OKButton.Size = New-Object System.Drawing.Size(75, 23)
$OKButton.Text = "OK"
#endregion

# If the OK button is left-clicked, select the highlighted menu item then close the pop-up menu.
$OKButton.Add_Click({ 
    if($adhocType -eq "PS"){
    $Global:adhocScript = $null
    $Global:adhocScript = $objTextBox1.text
    PSScript-Run}
    elseif($adhocType -eq "BAT"){
    $Global:adhocBATScript = $null
    $Global:adhocBATScript = $objTextBox1.text
    BATScript-Run}
$objForm.Close() })
$objForm.Controls.Add($OKButton)

$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Size(385, 250)
$CancelButton.Size = New-Object System.Drawing.Size(75, 23)
$CancelButton.Text = "Cancel"

# If the Cancel button is left-clicked, simply close the pop-up menu.
$CancelButton.Add_Click({ $objForm.Close() })
$objForm.Controls.Add($CancelButton)

# Shift the operating system focus to the pop-up menu then activate it — wait for the user response.
$objForm.Topmost = $True

$objForm.Add_Shown({ $objForm.Activate() })
[void] $objForm.ShowDialog()


}

Function PSScript-Run {

 Write-Debug ("ComboBox {0}" -f $uiHash.RunOptionComboBox.Text)
    $selectedItems = $uiHash.Listview.SelectedItems
    If ($selectedItems.Count -gt 0) {
        $uiHash.ProgressBar.Maximum = $selectedItems.count
        $uiHash.Listview.ItemsSource | ForEach {$_.Status = $Null; $_.notes = $Null} 
#region PSScript
            $uiHash.RunButton.IsEnabled = $False
            $uiHash.StartImage.Source = "$pwd\Images\Start_locked.jpg"
            $uiHash.CancelButton.IsEnabled = $True
            $uiHash.CancelImage.Source = "$pwd\Images\Stop.jpg"            
            $uiHash.StatusTextBox.Foreground = "Black"
            $uiHash.StatusTextBox.Text = "PS Script is being run..."            
            $uiHash.StartTime = (Get-Date)
            
            [Float]$uiHash.ProgressBar.Value = 0
            $scriptBlock = {
                Param (
                    $Computer,
                    $uiHash,
                    $Path,
                    $adhocScript
                )         
                $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                    $uiHash.Listview.Items.EditItem($Computer)
                    $computer.Status = "PS Script is being run..."
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh() 
                })                
                Set-Location $Path
                $Connection = (Test-Connection -ComputerName $Computer.computer -Count 1 -Quiet)
                        $server = $Computer.computer 
                        $temp = 'Try{' + "$adhocScript" + '}Catch{return $_.Exception.Message}'
                        $scrptBlck = [scriptblock]::Create($temp)
                        $remotesession = new-pssession -cn $server
                        $adhocres = invoke-command -ScriptBlock $scrptBlck -Session $remotesession | Out-String
                        $remotelastexitcode = invoke-command -ScriptBlock { return $?} -Session $remotesession
                        if ($remotelastexitcode){$status = "Completed"}
                        else{$status = "Failed"}
                        Clear-Variable -Name "adhocScript"
                $uiHash.ListView.Dispatcher.Invoke("Background",[action]{                    
                    $uiHash.Listview.Items.EditItem($Computer)
                    If ($Connection) {
                        $Computer.Status = $status
                        $Computer.Notes = $adhocres
                    } ElseIf (-Not $Connection) {
                        $Computer.Status = "Offline"
                    } Else {
                        $Computer.Status = "Unknown"
                    } 
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh()
                })
                $uiHash.ProgressBar.Dispatcher.Invoke("Normal",[action]{   
                    $uiHash.ProgressBar.value++  
                })
                $uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
                    #Check to see if find job
                    If ($uiHash.ProgressBar.value -eq $uiHash.ProgressBar.Maximum) { 
                        $End = New-Timespan $uihash.StartTime (Get-Date)                                             
                        $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)
                        $uiHash.RunButton.IsEnabled = $True
                        $uiHash.StartImage.Source = "$pwd\Images\Start.jpg"
                        $uiHash.CancelButton.IsEnabled = $False
                        $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.jpg"                    
                    }
                })  
            }
            Write-Verbose ("Creating runspace pool and session states")
            $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
            $runspaceHash.runspacepool = [runspacefactory]::CreateRunspacePool(1, $maxConcurrentJobs, $sessionstate, $Host)
            $runspaceHash.runspacepool.Open()   

            ForEach ($Computer in $selectedItems) {
                $uiHash.Listview.Items.EditItem($Computer)
                $computer.Status = "Pending PS script instance..."
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh() 
                #Create the powershell instance and supply the scriptblock with the other parameters 
                $powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($computer).AddArgument($uiHash).AddArgument($Path).AddArgument($adhocScript)
           
                #Add the runspace into the powershell instance
                $powershell.RunspacePool = $runspaceHash.runspacepool
           
                #Create a temporary collection for each runspace
                $temp = "" | Select-Object PowerShell,Runspace,Computer
                $Temp.Computer = $Computer.computer
                $temp.PowerShell = $powershell
           
                #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
                $temp.Runspace = $powershell.BeginInvoke()
                Write-Verbose ("Adding {0} collection" -f $temp.Computer)
                $jobs.Add($temp) | Out-Null                
            }
            #endregion
            } 
             Else {
                $uiHash.StatusTextBox.Foreground = "Red"
                $uiHash.StatusTextBox.Text = "No server/s selected!"
        }
}

Function BATScript-Run {

 Write-Debug ("ComboBox {0}" -f $uiHash.RunOptionComboBox.Text)
    $selectedItems = $uiHash.Listview.SelectedItems
    If ($selectedItems.Count -gt 0) {
        $uiHash.ProgressBar.Maximum = $selectedItems.count
        $uiHash.Listview.ItemsSource | ForEach {$_.Status = $Null; $_.notes = $Null} 
#region BATScript
            $uiHash.RunButton.IsEnabled = $False
            $uiHash.StartImage.Source = "$pwd\Images\Start_locked.jpg"
            $uiHash.CancelButton.IsEnabled = $True
            $uiHash.CancelImage.Source = "$pwd\Images\Stop.jpg"            
            $uiHash.StatusTextBox.Foreground = "Black"
            $uiHash.StatusTextBox.Text = "BAT Script is being run..."            
            $uiHash.StartTime = (Get-Date)
            
            [Float]$uiHash.ProgressBar.Value = 0
            $scriptBlock = {
                Param (
                    $Computer,
                    $uiHash,
                    $Path,
                    $adhocBATScript
                )         
                $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                    $uiHash.Listview.Items.EditItem($Computer)
                    $computer.Status = "BAT Script is being run..."
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh() 
                })                
                Set-Location $Path
                $Connection = (Test-Connection -ComputerName $Computer.computer -Count 1 -Quiet)
                        $server = $Computer.computer 
                        $TestScriptBlock = {
                                try{
                                $targetPath = 'C:\Temp\tempContainer.bat'
                                $testPath = test-path -path $targetPath
                                    if($testPath){Remove-Item –path $targetPath –recurse}
                                    Add-Content -Path $targetPath -Value $Using:adhocBATScript
                                    #$cmdOutput = cmd /c $targetPath '2>&1'
                                    #$cmdOutput = Start-Process -FilePath cmd.exe -ArgumentList '/c', $targetPath -Wait 
                                    $cmdOutput = Invoke-Expression -Command:"cmd.exe /c 'C:\Temp\tempContainer.bat' -verb RunAs -wait"
                                    Remove-Item –path $targetPath –recurse
                                    return $cmdOutput
                                }Catch{
                                return $_.Exception.Message
                                }
                        }
                        
                        $remotesession = new-pssession -computername $server
                        $adhocres = invoke-command -ScriptBlock $TestScriptBlock -Session $remotesession | Out-String
                        $remotelastexitcode = invoke-command -ScriptBlock { return $?} -Session $remotesession
                        if ($remotelastexitcode){ $status = "Completed"}
                        else{$status = "Failed"}
                        
                        Clear-Variable -Name "adhocBATScript"
                $uiHash.ListView.Dispatcher.Invoke("Background",[action]{                    
                    $uiHash.Listview.Items.EditItem($Computer)
                    If ($Connection) {
                        $Computer.Status = $status
                        $Computer.Notes = $adhocres
                    } ElseIf (-Not $Connection) {
                        $Computer.Status = "Offline"
                    } Else {
                        $Computer.Status = "Unknown"
                    } 
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh()
                })
                $uiHash.ProgressBar.Dispatcher.Invoke("Normal",[action]{   
                    $uiHash.ProgressBar.value++  
                })
                $uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
                    #Check to see if find job
                    If ($uiHash.ProgressBar.value -eq $uiHash.ProgressBar.Maximum) { 
                        $End = New-Timespan $uihash.StartTime (Get-Date)                                             
                        $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)
                        $uiHash.RunButton.IsEnabled = $True
                        $uiHash.StartImage.Source = "$pwd\Images\Start.jpg"
                        $uiHash.CancelButton.IsEnabled = $False
                        $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.jpg"                    
                    }
                })  
            }
            Write-Verbose ("Creating runspace pool and session states")
            $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
            $runspaceHash.runspacepool = [runspacefactory]::CreateRunspacePool(1, $maxConcurrentJobs, $sessionstate, $Host)
            $runspaceHash.runspacepool.Open()   

            ForEach ($Computer in $selectedItems) {
                $uiHash.Listview.Items.EditItem($Computer)
                $computer.Status = "Pending BAT script instance..."
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh() 
                #Create the powershell instance and supply the scriptblock with the other parameters 
                $powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($computer).AddArgument($uiHash).AddArgument($Path).AddArgument($adhocBATScript)
           
                #Add the runspace into the powershell instance
                $powershell.RunspacePool = $runspaceHash.runspacepool
           
                #Create a temporary collection for each runspace
                $temp = "" | Select-Object PowerShell,Runspace,Computer
                $Temp.Computer = $Computer.computer
                $temp.PowerShell = $powershell
           
                #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
                $temp.Runspace = $powershell.BeginInvoke()
                Write-Verbose ("Adding {0} collection" -f $temp.Computer)
                $jobs.Add($temp) | Out-Null                
            }
            #endregion
            } 
             Else {
                $uiHash.StatusTextBox.Foreground = "Red"
                $uiHash.StatusTextBox.Text = "No server/s selected!"
        }
}


Function Chk-ClstrNameNode{
 Write-Debug ("ComboBox {0}" -f $uiHash.RunOptionComboBox.Text)
        $selectedItems = $uiHash.Listview.SelectedItems
        If ($selectedItems.Count -gt 0) {
        $uiHash.ProgressBar.Maximum = $selectedItems.count
        $uiHash.Listview.ItemsSource | ForEach {$_.Status = $Null; $_.notes = $Null} 
        #region ClstrNameNode
            $uiHash.RunButton.IsEnabled = $False
            $uiHash.StartImage.Source = "$pwd\Images\Start_locked.jpg"
            $uiHash.CancelButton.IsEnabled = $True
            $uiHash.CancelImage.Source = "$pwd\Images\Stop.jpg"            
            $uiHash.StatusTextBox.Foreground = "Black"
            $uiHash.StatusTextBox.Text = "Checking Cluster Name/Node..."            
            $uiHash.StartTime = (Get-Date)
            
            [Float]$uiHash.ProgressBar.Value = 0
            $scriptBlock = {
                Param (
                    $Computer,
                    $uiHash,
                    $Path,
                    $invkClstr
                )               
                
                $Connection = (Test-Connection -ComputerName $Computer.computer -Count 1 -Quiet)
                $server = $Computer.computer
            Switch ($invkClstr)
            {
            "chk"   { Try{
                          $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                                  $uiHash.Listview.Items.EditItem($Computer)
                                  $computer.Status = "Checking Cluster Name/Node..."
                                  $uiHash.Listview.Items.CommitEdit()
                                  $uiHash.Listview.Items.Refresh()})
                          $resout, $status = Invoke-Command -ComputerName $server -ScriptBlock{
                              $chkFO = (Get-WindowsFeature Failover*).installed
                              $clstrSrvice = (Get-Service -Name ClusSvc).Status
                              $ownerGrp = (Get-ClusterResource | Where-Object {$_.ResourceType -eq "SQL Server Availability Group"}).name
                              if($chkFO){
                               if($clstrSrvice -eq "Running"){
                                  $ServiceInstance = (Get-ClusterResource | select Name, ResourceType |
                                  Where-Object {$_.ResourceType -eq "SQL Server Availability Group"}).name
                                  if($ServiceInstance -eq $null){return "SQL Server Availability Group resource type not found! Please check manually", "Failed"}
                                  else{$OwnerNode = (Get-ClusterGroup | Select Name, OwnerNode |
                                  Where-Object {$_.Name -eq $ServiceInstance}).OwnerNode
                                  $Node = ($OwnerNode | Select Name).name
                                  $thisNode = Get-Content env:COMPUTERNAME
                                  if($Node -eq $thisNode){$nt = "Active"}else{$nt = "Passive"}}
                                  $ClstrName = (Get-Cluster|Select Name).name
                                  $clstrNodes = (Get-ClusterNode | select Name).name
                                  $resout = "$ClstrName - $clstrNodes - $ownerGrp - ($nt)"
                                  $status = "Completed"
                                  return $resout, $status
                                }Else{return "Cluster Service NOT running!","Failed"}
                              }Else{ return "No Failover Cluster Installed!","Failed"}
                          }
                      }Catch{$resout = $($error[0]); $status = "Failed"}
                    }
            "move"  {   
                       $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                                 $uiHash.Listview.Items.EditItem($Computer)
                                 $computer.Status = "Moving resource(s), Please wait..."
                                 $uiHash.Listview.Items.CommitEdit()
                                 $uiHash.Listview.Items.Refresh()}) 
                       $chkFO, $nt = Invoke-Command -ComputerName $server -ScriptBlock{
                            $chkFO = (Get-WindowsFeature Failover*).installed
                            $ServiceInstance = (Get-ClusterResource | select Name, ResourceType |
                                  Where-Object {$_.ResourceType -eq "SQL Server Availability Group"}).name
                                  if($ServiceInstance -eq $null){$nt = "SQL Server Availability Group resource type not found! Please check manually!"}
                                  else{$OwnerNode = (Get-ClusterGroup | Select Name, OwnerNode |
                                  Where-Object {$_.Name -eq $ServiceInstance}).OwnerNode
                                  $Node = ($OwnerNode | Select Name).name
                                  $thisNode = Get-Content env:COMPUTERNAME
                                  if($Node -eq $thisNode){$nt = "Active"}else{ $nt = "Passive"}}
                            return $chkFO, $nt
                            }
                       if($chkFO){
                        if($nt -eq 'Active'){
                           Try{
                               $resout = Invoke-Command -ComputerName $server -ScriptBlock{Get-ClusterNode $env:COMPUTERNAME | Get-ClusterGroup | Move-ClusterGroup | Out-String}
                               $status = "Completed"
                           }Catch{$resout = $($error[0]); $status = "Failed"}
                        }Elseif($nt -eq 'Passive'){ $resout =  "Server is NOT an ACTIVE node!"; $status = "Failed"}
                        Else{$resout =  "SQL Server Availability Group resource type not found! Please check manually!"; $status = "Failed"}
                       }Else{ $resout =  "No Failover Cluster Installed!"; $status = "Failed"}
                    }
            Default {$Null}
            }
               $uiHash.ListView.Dispatcher.Invoke("Background",[action]{                    
                    $uiHash.Listview.Items.EditItem($Computer)
                    If ($Connection) {
                        $Computer.status = $status
                        $Computer.notes = $resout
                    } ElseIf (-Not $Connection) {
                        $Computer.Status = "Offline"
                    } Else {
                        $Computer.Status = "Unknown"
                    } 
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh()
                })
                $uiHash.ProgressBar.Dispatcher.Invoke("Normal",[action]{   
                    $uiHash.ProgressBar.value++  
                })
                $uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
                    #Check to see if find job
                    If ($uiHash.ProgressBar.value -eq $uiHash.ProgressBar.Maximum) { 
                        $End = New-Timespan $uihash.StartTime (Get-Date)                                             
                        $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)
                        $uiHash.RunButton.IsEnabled = $True
                        $uiHash.StartImage.Source = "$pwd\Images\Start.jpg"
                        $uiHash.CancelButton.IsEnabled = $False
                        $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.jpg"                    
                    }
                })  
                
            }

            Write-Verbose ("Creating runspace pool and session states")
            $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
            $runspaceHash.runspacepool = [runspacefactory]::CreateRunspacePool(1, $maxConcurrentJobs, $sessionstate, $Host)
            $runspaceHash.runspacepool.Open()   

            ForEach ($Computer in $selectedItems) {
                $uiHash.Listview.Items.EditItem($Computer)
                $computer.Status = "Pending Cluster Name/Node checking..."
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh() 
                #Create the powershell instance and supply the scriptblock with the other parameters 
                $powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($computer).AddArgument($uiHash).AddArgument($Path).AddArgument($invkClstr)
           
                #Add the runspace into the powershell instance
                $powershell.RunspacePool = $runspaceHash.runspacepool
           
                #Create a temporary collection for each runspace
                $temp = "" | Select-Object PowerShell,Runspace,Computer
                $Temp.Computer = $Computer.computer
                $temp.PowerShell = $powershell
           
                #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
                $temp.Runspace = $powershell.BeginInvoke()
                Write-Verbose ("Adding {0} collection" -f $temp.Computer)
                $jobs.Add($temp) | Out-Null                
            }
            #endregion

        } 
             Else {
                $uiHash.StatusTextBox.Foreground = "Red"
                $uiHash.StatusTextBox.Text = "No server/s selected!"
        }

}

Function Check-ChefLogs{
 Write-Debug ("ComboBox {0}" -f $uiHash.RunOptionComboBox.Text)
        $selectedItems = $uiHash.Listview.SelectedItems
        If ($selectedItems.Count -gt 0) {
        $uiHash.ProgressBar.Maximum = $selectedItems.count
        $uiHash.Listview.ItemsSource | ForEach {$_.Status = $Null; $_.notes = $Null} 
	#region CheckChefLogs
            $uiHash.RunButton.IsEnabled = $False
            $uiHash.StartImage.Source = "$pwd\Images\Start_locked.jpg"
            $uiHash.CancelButton.IsEnabled = $True
            $uiHash.CancelImage.Source = "$pwd\Images\Stop.jpg"            
            $uiHash.StatusTextBox.Foreground = "Black"
            $uiHash.StatusTextBox.Text = "Checking CHEF Logs..."            
            $uiHash.StartTime = (Get-Date)
            
            [Float]$uiHash.ProgressBar.Value = 0
            $scriptBlock = {
                Param (
                    $Computer,
                    $uiHash,
                    $Path
                )               
                $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                    $uiHash.Listview.Items.EditItem($Computer)
                    $computer.Status = "Running"
					$computer.Notes = "Checking CHEF Logs... please wait."
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh() 
                })  
                $Connection = (Test-Connection -ComputerName $Computer.computer -Count 1 -Quiet)
                $server = $Computer.computer
                Try{#$date = Invoke-Command -cn $server -scriptblock{get-date -Format "yyyyMMdd"
                $localTime = Get-WmiObject -Class win32_localtime -ComputerName $server
                $date = get-date -Year $localTime.Year -Month $localTime.Month -Day $localTime.Day -Format "yyyyMMdd"
                }Catch{$dateErr = $_.Exception.Message | Out-String}
                    if($date -ne ""){
                    try{$checkPath = Test-Path -path "\\$server\c$\users\administrator\patch.log"}
                    Catch{$chkPathErr = $_.Exception.Message | out-string}
                        if($checkPath){
                        $getchefLogs = Get-Content -Path "\\$server\c$\users\administrator\patch.log" -tail 10
                            if($?){
                                $chefLogs = $getchefLogs | Select-String $date
                                if($chefLogs -match "Patching operation complete" -or $chefLogs -match "up-to-date"){
                                    $status = "Completed" 
                                    $resout = "Patch Completed"}
                                elseif($chefLogs -match "System is ready for manual reboot"){
                                    $status = "Completed" 
                                    $resout = "System is ready for manual reboot, please reboot when ready."}
                                elseif($chefLogs -match "Patch Failure"){
                                    $getErr = $chefLogs | Select-String -pattern "Patch Failure"
                                    $shwErr = $getErr | ForEach-Object {$_.line.Substring($_.Line.LastIndexOf('Patch Failure:'))}
                                    $resout = $shwErr
                                    $status = "Failed" }
                                elseif($chefLogs -match "Error"){
                                    $getErr = $chefLogs | Select-String -pattern "Error"
                                    $shwErr = $getErr | ForEach-Object {$_.line.Substring($_.Line.LastIndexOf('error:'))}
                                    $resout = $shwErr
                                    $status = "Failed" }
                                else{
                                    $status = "Completed" 
                                    $resout = Get-Content -Path "\\$server\c$\users\administrator\patch.log" -tail 10 | Out-String}
                            }Else{
                                $resout = "No Chef Patch Logs found for today's date!"
                                $status = "Failed"
                            }
                        }Else{
                        $resout = "patch.log not found! Error:" + $chkPathErr
                        $status = "Failed"}
                    }Else{
                    $resout = "Unable to get today's date! Error: " + $dateErr
                    $status = "Failed"}

               $uiHash.ListView.Dispatcher.Invoke("Background",[action]{                    
                    $uiHash.Listview.Items.EditItem($Computer)
                    If ($Connection) {
                        $Computer.status = $status
                        $Computer.notes = $resout
                    } ElseIf (-Not $Connection) {
                        $Computer.Status = "Offline"
                        $Computer.notes = "Check Server"
                    } Else {
                        $Computer.Status = "Unknown"
                        $Computer.notes = "Check Server"
                    } 
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh()
                })
                $uiHash.ProgressBar.Dispatcher.Invoke("Normal",[action]{   
                    $uiHash.ProgressBar.value++  
                })
                $uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
                    #Check to see if find job
                    If ($uiHash.ProgressBar.value -eq $uiHash.ProgressBar.Maximum) { 
                        $End = New-Timespan $uihash.StartTime (Get-Date)                                             
                        $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)
                        $uiHash.RunButton.IsEnabled = $True
                        $uiHash.StartImage.Source = "$pwd\Images\Start.jpg"
                        $uiHash.CancelButton.IsEnabled = $False
                        $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.jpg"                    
                    }
                })  
                
            }

            Write-Verbose ("Creating runspace pool and session states")
            $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
            $runspaceHash.runspacepool = [runspacefactory]::CreateRunspacePool(1, $maxConcurrentJobs, $sessionstate, $Host)
            $runspaceHash.runspacepool.Open()   

            ForEach ($Computer in $selectedItems) {
                $uiHash.Listview.Items.EditItem($Computer)
                $computer.Status = "Pending..."
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh() 
                #Create the powershell instance and supply the scriptblock with the other parameters 
                $powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($computer).AddArgument($uiHash).AddArgument($Path)
           
                #Add the runspace into the powershell instance
                $powershell.RunspacePool = $runspaceHash.runspacepool
           
                #Create a temporary collection for each runspace
                $temp = "" | Select-Object PowerShell,Runspace,Computer
                $Temp.Computer = $Computer.computer
                $temp.PowerShell = $powershell
           
                #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
                $temp.Runspace = $powershell.BeginInvoke()
                Write-Verbose ("Adding {0} collection" -f $temp.Computer)
                $jobs.Add($temp) | Out-Null                
            }
            #endregion

        } 
             Else {
                $uiHash.StatusTextBox.Foreground = "Red"
                $uiHash.StatusTextBox.Text = "No server/s selected!"
        }

}

Function System-Reboot{
        Write-Debug ("ComboBox {0}" -f $uiHash.RunOptionComboBox.Text)
        $selectedItems = $uiHash.Listview.SelectedItems
        If ($selectedItems.Count -gt 0) {
        $uiHash.ProgressBar.Maximum = $selectedItems.count
        $uiHash.Listview.ItemsSource | ForEach {$_.Status = $Null; $_.notes = $Null} 
        #region Rebootss
            #If ((Show-Warning) -eq "Yes") {
                $uiHash.RunButton.IsEnabled = $False
                $uiHash.StartImage.Source = "$pwd\Images\Start_locked.jpg"
                $uiHash.CancelButton.IsEnabled = $True
                $uiHash.CancelImage.Source = "$pwd\Images\Stop.jpg"            
                $uiHash.StatusTextBox.Foreground = "Black"
                $uiHash.StatusTextBox.Text = "Rebooting Server(s)..."            
                $uiHash.StartTime = (Get-Date)
            
                [Float]$uiHash.ProgressBar.Value = 0
                $scriptBlock = {
                    Param (
                        $Computer,
                        $uiHash,
                        $Path
                    )               
                    $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                        $uiHash.Listview.Items.EditItem($Computer)
                        $computer.Status = "Rebooting"
                        $uiHash.Listview.Items.CommitEdit()
                        $uiHash.Listview.Items.Refresh() 
                    })                
                    Set-Location $Path
                    If (Test-Connection -Computer $Computer.computer -count 1 -Quiet) {
                        Try {
                            Restart-Computer -ComputerName $Computer.computer -Force -ea stop
                            Do {
                                Start-Sleep -Seconds 2
                                Write-Verbose ("Waiting for {0} to shutdown..." -f $Computer.computer)
                                $computer.Notes = "Waiting for {0} to shutdown..." -f $Computer.computer
                            }
                            While ((Test-Connection -ComputerName $Computer.computer -Count 1 -Quiet))    
                            Do {
                                Start-Sleep -Seconds 5
                                $i++        
                                Write-Verbose ("{0} down...{1}" -f $Computer.computer, $i)
                                $computer.Notes = "{0} down...{1}" -f $Computer.computer, $i
                                If($i -eq 60) {
                                    Write-Warning ("{0} did not come back online from reboot!" -f $Computer.computer)
                                    $computer.Notes = "{0} did not come back online from reboot!" -f $Computer.computer
                                    $connection = $False
                                }
                            }
                            While (-NOT(Test-Connection -ComputerName $Computer.computer -Count 1 -Quiet))
                            Write-Verbose ("{0} is back up" -f $Computer.computer)
                            $connection = $True
                        } Catch {
                            Write-Warning "$($Error[0])"
                            $connection = $False
                        }
                    }

                    $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{                    
                        $uiHash.Listview.Items.EditItem($Computer)
                        If ($Connection) {
                            $Computer.Status = "Complete"
                            $Computer.Notes = "Online"
                        } ElseIf (-Not $Connection) {
                            $Computer.Status = "Failed"
                        } Else {
                            $Computer.Status = "Unknown"
                        } 
                        $uiHash.Listview.Items.CommitEdit()
                        $uiHash.Listview.Items.Refresh()
                    })
                    $uiHash.ProgressBar.Dispatcher.Invoke("Normal",[action]{   
                        $uiHash.ProgressBar.value++  
                    })
                    $uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
                        #Check to see if find job
                        If ($uiHash.ProgressBar.value -eq $uiHash.ProgressBar.Maximum) { 
                            $End = New-Timespan $uihash.StartTime (Get-Date)                                             
                            $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)
                            $uiHash.RunButton.IsEnabled = $True
                            $uiHash.StartImage.Source = "$pwd\Images\Start.jpg"
                            $uiHash.CancelButton.IsEnabled = $False
                            $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.jpg"                    
                        }
                    })  
                
                }

            Write-Verbose ("Creating runspace pool and session states")
            $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
            $runspaceHash.runspacepool = [runspacefactory]::CreateRunspacePool(1, $maxRebootJobs, $sessionstate, $Host)
            $runspaceHash.runspacepool.Open()  

                ForEach ($Computer in $selectedItems) {
                    $uiHash.Listview.Items.EditItem($Computer)
                    $computer.Status = "Pending Reboot"
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh() 
                    #Create the powershell instance and supply the scriptblock with the other parameters 
                    $powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($computer).AddArgument($uiHash).AddArgument($Path)
           
                    #Add the runspace into the powershell instance
                    $powershell.RunspacePool = $runspaceHash.runspacepool
           
                    #Create a temporary collection for each runspace
                    $temp = "" | Select-Object PowerShell,Runspace,Computer
                    $Temp.Computer = $Computer.computer
                    $temp.PowerShell = $powershell
           
                    #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
                    $temp.Runspace = $powershell.BeginInvoke()
                    Write-Verbose ("Adding {0} collection" -f $temp.Computer)
                    $jobs.Add($temp) | Out-Null
                }                
            #endregion 
            } 
             Else {
                $uiHash.StatusTextBox.Foreground = "Red"
                $uiHash.StatusTextBox.Text = "No server/s selected!"
        }

}

Function RobocopyPrompt{

#region Get Source/Destination Path
                $objForm = New-Object System.Windows.Forms.Form 
                $objForm.Text = "Robocopy"
                $objForm.Size = New-Object System.Drawing.Size(400,200) 
                $objForm.StartPosition = "CenterScreen"
                
                $objForm.KeyPreview = $True
                $objForm.Add_KeyDown({
                    if ($_.KeyCode -eq "Enter" -or $_.KeyCode -eq "Escape"){
                        $objForm.Close()
                    }
                })
                
                $CancelButton = New-Object System.Windows.Forms.Button
                $CancelButton.Location = New-Object System.Drawing.Size(295,120)
                $CancelButton.Size = New-Object System.Drawing.Size(75,23)
                $CancelButton.Text = "Cancel"
                $CancelButton.Add_Click({$objForm.Close()})
                $objForm.Controls.Add($CancelButton)
                
                $objLabel = New-Object System.Windows.Forms.Label
                $objLabel.Location = New-Object System.Drawing.Size(10,20) 
                $objLabel.Size = New-Object System.Drawing.Size(280,20) 
                $objLabel.Text = 'Please enter "SOURCE PATH" at the space below:'
                $objForm.Controls.Add($objLabel) 
                
                $objTextBox = New-Object System.Windows.Forms.TextBox 
                $objTextBox.Location = New-Object System.Drawing.Size(10,40) 
                $objTextBox.Size = New-Object System.Drawing.Size(360,40) 
                $objForm.Controls.Add($objTextBox) 
                
                $objLabel2 = New-Object System.Windows.Forms.Label
                $objLabel2.Location = New-Object System.Drawing.Size(10,70) 
                $objLabel2.Size = New-Object System.Drawing.Size(320,20) 
                $objLabel2.Text = 'Enter "DESTINATION FOLDER NAME" at the space below:'
                $objForm.Controls.Add($objLabel2) 
                
                $objTextBox2 = New-Object System.Windows.Forms.TextBox 
                $objTextBox2.Location = New-Object System.Drawing.Size(10,90) 
                $objTextBox2.Size = New-Object System.Drawing.Size(360,40) 
                $objForm.Controls.Add($objTextBox2) 

                $OKButton = New-Object System.Windows.Forms.Button
                $OKButton.Location = New-Object System.Drawing.Size(220,120)
                $OKButton.Size = New-Object System.Drawing.Size(75,23)
                $OKButton.Text = "OK"
                $OKButton.Add_Click({
                $objForm.Close()
                $Global:adhocScript =  $null
                $Global:adhocBATScript =  $null
                $Global:adhocScript = $objTextBox.text
                $Global:adhocBATScript = $objTextBox2.text
                RoboCopy
                })
                $objForm.Controls.Add($OKButton)
                
                $objForm.Topmost = $True
                
                $objForm.Add_Shown({$objForm.Activate()})
                [void]$objForm.ShowDialog()
               
              #endregion
              
}

Function RoboCopy{
 Write-Debug ("ComboBox {0}" -f $uiHash.RunOptionComboBox.Text)
        $selectedItems = $uiHash.Listview.SelectedItems
        If ($selectedItems.Count -gt 0) {
        $uiHash.ProgressBar.Maximum = $selectedItems.count
        $uiHash.Listview.ItemsSource | ForEach {$_.Status = $Null; $_.notes = $Null} 
        #region CheckChefLogs
            $uiHash.RunButton.IsEnabled = $False
            $uiHash.StartImage.Source = "$pwd\Images\Start_locked.jpg"
            $uiHash.CancelButton.IsEnabled = $True
            $uiHash.CancelImage.Source = "$pwd\Images\Stop.jpg"            
            $uiHash.StatusTextBox.Foreground = "Black"
            $uiHash.StatusTextBox.Text = "Starting robocopy..."            
            $uiHash.StartTime = (Get-Date)
            
            [Float]$uiHash.ProgressBar.Value = 0
            $scriptBlock = {
                Param (
                    $Computer,
                    $uiHash,
                    $Path,
                    $adhocScript,
                    $adhocBATScript
                )               
                $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                    $uiHash.Listview.Items.EditItem($Computer)
                    $computer.Status = "Copying, Please wait..."
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh() 
                })                
              

               
                $sourcePath = $adhocScript
                $targetFolder = $adhocBATScript
                $Connection = (Test-Connection -ComputerName $Computer.computer -Count 1 -Quiet)
                $Server = $Computer.computer
                    
                Try{
                    $Destination = "\\" + $Server + "\C$\" + $targetFolder
                    robocopy "$sourcePath" $Destination /S /Z /R:1 /NFL /NDL /NJH /NJS /nc /ns /np
                    $status = "Completed"
                    $notes = "File(s) successfully copied at c:\" + $targetFolder
                 }Catch{
                    $status = "Failed"
                    $err = $_.Exception.Message 
                 }
                    

                $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{
                    $uiHash.Listview.Items.EditItem($Computer)
                    If ($Connection) {
                        if($status -eq "Completed"){
                            $Computer.Status = "Completed"
                            $Computer.Notes = $notes}
                        else{
                            $Computer.Status = "Failed"
                            $Computer.Notes = $err | out-string}
                    } ElseIf (-Not $Connection) {
                        $Computer.Status = "Offline"
                    } Else {
                        $Computer.Status = "Unknown"
                    } 
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh()
                })
                $uiHash.ProgressBar.Dispatcher.Invoke("Normal",[action]{   
                    $uiHash.ProgressBar.value++  
                })
                $uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
                    #Check to see if find job
                    If ($uiHash.ProgressBar.value -eq $uiHash.ProgressBar.Maximum) { 
                        $End = New-Timespan $uihash.StartTime (Get-Date)                                             
                        $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)
                        $uiHash.RunButton.IsEnabled = $True
                        $uiHash.StartImage.Source = "$pwd\Images\Start.jpg"
                        $uiHash.CancelButton.IsEnabled = $False
                        $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.jpg"                    
                    }
                })  
                
            }

            Write-Verbose ("Creating runspace pool and session states")
            $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
            $runspaceHash.runspacepool = [runspacefactory]::CreateRunspacePool(1, $maxConcurrentJobs, $sessionstate, $Host)
            $runspaceHash.runspacepool.Open()   

            ForEach ($Computer in $selectedItems) {
                $uiHash.Listview.Items.EditItem($Computer)
                $computer.Status = "Pending Robocopy"
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh() 
                #Create the powershell instance and supply the scriptblock with the other parameters 
                $powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($computer).AddArgument($uiHash).AddArgument($Path).AddArgument($adhocScript).AddArgument($adhocBATScript)
           
                #Add the runspace into the powershell instance
                $powershell.RunspacePool = $runspaceHash.runspacepool
           
                #Create a temporary collection for each runspace
                $temp = "" | Select-Object PowerShell,Runspace,Computer
                $Temp.Computer = $Computer.computer
                $temp.PowerShell = $powershell
           
                #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
                $temp.Runspace = $powershell.BeginInvoke()
                Write-Verbose ("Adding {0} collection" -f $temp.Computer)
                $jobs.Add($temp) | Out-Null                
            }
            #endregion

        } 
             Else {
                $uiHash.StatusTextBox.Foreground = "Red"
                $uiHash.StatusTextBox.Text = "No server/s selected!"
        }

}

Function Run-2k8ESULicense{
 Write-Debug ("ComboBox {0}" -f $uiHash.RunOptionComboBox.Text)
        $selectedItems = $uiHash.Listview.SelectedItems
        If ($selectedItems.Count -gt 0) {
        $uiHash.ProgressBar.Maximum = $selectedItems.count
        $uiHash.Listview.ItemsSource | ForEach {$_.Status = $Null; $_.notes = $Null} 
        #region CheckChefLogs
            $uiHash.RunButton.IsEnabled = $False
            $uiHash.StartImage.Source = "$pwd\Images\Start_locked.jpg"
            $uiHash.CancelButton.IsEnabled = $True
            $uiHash.CancelImage.Source = "$pwd\Images\Stop.jpg"            
            $uiHash.StatusTextBox.Foreground = "Black"
            $uiHash.StatusTextBox.Text = "Initializing ESU License..."            
            $uiHash.StartTime = (Get-Date)
            
            [Float]$uiHash.ProgressBar.Value = 0
            $scriptBlock = {
                Param (
                    $Computer,
                    $uiHash,
                    $Path,
                    $ESUreport
                )               
                $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                    $uiHash.Listview.Items.EditItem($Computer)
                    $computer.Status = "Installation and activation in progress, Please wait..."
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh() 
                })   

                Set-Location $path
                . .\Scripts\2k8_ESU_License.ps1                
                $ESURes = @(2k8_ESU_License -Computer $computer.computer)

                $ESUreport.AddRange($ESURes) | Out-Null

                $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{
                    $uiHash.Listview.Items.EditItem($Computer)
                    If ($ESURes[0].Error -eq 0) {
                        $Computer.Status = "Completed"
                        $computer.Notes = $ESURes[0].ESU_License
                    }ElseIf ($ESURes[0].Error -eq "Offline") {
                        $Computer.Status = "Offline"
                    } 
                    Elseif($ESURes[0].Error -gt 0) {
                        $errcnt = $ESURes[0].Error
                        $Computer.Status = "Failed"
                        #$computer.Notes = "There were $errcnt error(s) encountered! Please check ESU report."
                        $computer.Notes = $ESURes[0].Notes
                    }
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh()
                })
                $uiHash.ProgressBar.Dispatcher.Invoke("Normal",[action]{   
                    $uiHash.ProgressBar.value++  
                })
                $uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
                    #Check to see if find job
                    If ($uiHash.ProgressBar.value -eq $uiHash.ProgressBar.Maximum) { 
                        $End = New-Timespan $uihash.StartTime (Get-Date)                                             
                        $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)
                        $uiHash.RunButton.IsEnabled = $True
                        $uiHash.StartImage.Source = "$pwd\Images\Start.jpg"
                        $uiHash.CancelButton.IsEnabled = $False
                        $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.jpg"                    
                    }
                })  
                
            }

            Write-Verbose ("Creating runspace pool and session states")
            $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
            $runspaceHash.runspacepool = [runspacefactory]::CreateRunspacePool(1, $maxConcurrentJobs, $sessionstate, $Host)
            $runspaceHash.runspacepool.Open()   

            ForEach ($Computer in $selectedItems) {
                $uiHash.Listview.Items.EditItem($Computer)
                $computer.Status = "Pending ESU License Activation"
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh() 
                #Create the powershell instance and supply the scriptblock with the other parameters 
                $powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($computer).AddArgument($uiHash).AddArgument($Path).AddArgument($ESUreport)
                
                #Add the runspace into the powershell instance
                $powershell.RunspacePool = $runspaceHash.runspacepool
           
                #Create a temporary collection for each runspace
                $temp = "" | Select-Object PowerShell,Runspace,Computer
                $Temp.Computer = $Computer.computer
                $temp.PowerShell = $powershell
           
                #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
                $temp.Runspace = $powershell.BeginInvoke()
                Write-Verbose ("Adding {0} collection" -f $temp.Computer)
                $jobs.Add($temp) | Out-Null  
                              
            }
            #endregion

        } 
             Else {
                $uiHash.StatusTextBox.Foreground = "Red"
                $uiHash.StatusTextBox.Text = "No server/s selected!"
        }

}

Function Check-SEPver{
 Write-Debug ("ComboBox {0}" -f $uiHash.RunOptionComboBox.Text)
        $selectedItems = $uiHash.Listview.SelectedItems
        If ($selectedItems.Count -gt 0) {
        $uiHash.ProgressBar.Maximum = $selectedItems.count
        $uiHash.Listview.ItemsSource | ForEach {$_.Status = $Null; $_.notes = $Null} 
        #region CheckSEPver
            $uiHash.RunButton.IsEnabled = $False
            $uiHash.StartImage.Source = "$pwd\Images\Start_locked.jpg"
            $uiHash.CancelButton.IsEnabled = $True
            $uiHash.CancelImage.Source = "$pwd\Images\Stop.jpg"            
            $uiHash.StatusTextBox.Foreground = "Black"
            $uiHash.StatusTextBox.Text = "Checking SEP version..."            
            $uiHash.StartTime = (Get-Date)
            
            [Float]$uiHash.ProgressBar.Value = 0
            $scriptBlock = {
                Param (
                    $Computer,
                    $uiHash,
                    $Path
                )               
                $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                    $uiHash.Listview.Items.EditItem($Computer)
                    $computer.Status = "Checking SEP version..."
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh() 
                })                
                Set-Location $Path
                $Connection = (Test-Connection -ComputerName $Computer.computer -Count 1 -Quiet)
                $server = $Computer.computer 
                    try{$status, $resout = Invoke-Command -ComputerName $server -ScriptBlock {
                          try{
                            $SEPVer = (Get-ItemProperty 'HKLM:\SOFTWARE\Symantec\Symantec Endpoint Protection\CurrentVersion' | select "PRODUCTVERSION")."PRODUCTVERSION"
                            $SEPDef = (get-ItemProperty 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Symantec\Symantec Endpoint Protection\CurrentVersion\public-opstate'  | select LatestVirusDefsDate).LatestVirusDefsDate
                            $rev = (get-ItemProperty 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Symantec\Symantec Endpoint Protection\CurrentVersion\public-opstate' | select LatestVirusDefsRevision).LatestVirusDefsRevision
                            $SEPsrvce = (Get-Service -DisplayName "Symantec Endpoint Protection" | select Status).Status
                            if($SEPVer -ne $null){
                                if($SEPsrvce -eq "Running"){
                                   $SEPDir = (Get-ItemProperty 'HKLM:\SOFTWARE\Symantec\Symantec Endpoint Protection\CurrentVersion\Common Client\PathExpansionMap\INSTALLDIR' | select "Path32")."Path32"
                            
                                   $res =  "$SEPVer; Def: $SEPDef r$rev; Dir: $SEPDir"
                                   return "Completed", $res
                                }Else{try{  $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                                            $uiHash.Listview.Items.EditItem($Computer)
                                            $computer.Status = "Running"
                                            $computer.Notes = "Starting SEP services..."
                                            $uiHash.Listview.Items.CommitEdit()
                                            $uiHash.Listview.Items.Refresh()})
                                            Get-Service -ComputerName $server -DisplayName "Symantec Endpoint Protection" | Start-Service
                                            Start-Sleep -s 3
                                           $res =  "$SEPVer; Def: $SEPDef r$rev; Dir: $SEPDir"; return "Completed", $res}
                                       catch{return "Failed", "Failed to start SEP services, please check!"}
                                }
                            }Else{return "Failed", "SEP not installed!"}
                          }Catch{$Status = "Failed"; $resout = $_.Exception.Message }
                        } -ErrorVariable e -ErrorAction SilentlyContinue #EndScriptBlock
                        if($e){$Status = "Failed"; $resout = $e | Out-String}
                    }Catch{$Status = "Failed"; $resout = $_.Exception.Message }

               $uiHash.ListView.Dispatcher.Invoke("Background",[action]{                    
                    $uiHash.Listview.Items.EditItem($Computer)
                    If ($Connection) {
                        $Computer.status = $status
                        $Computer.notes = $resout
                    } ElseIf (-Not $Connection) {
                        $Computer.Status = "Offline"
                    } Else {
                        $Computer.Status = "Unknown"
                    } 
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh()
                })
                $uiHash.ProgressBar.Dispatcher.Invoke("Normal",[action]{   
                    $uiHash.ProgressBar.value++  
                })
                $uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
                    #Check to see if find job
                    If ($uiHash.ProgressBar.value -eq $uiHash.ProgressBar.Maximum) { 
                        $End = New-Timespan $uihash.StartTime (Get-Date)                                             
                        $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)
                        $uiHash.RunButton.IsEnabled = $True
                        $uiHash.StartImage.Source = "$pwd\Images\Start.jpg"
                        $uiHash.CancelButton.IsEnabled = $False
                        $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.jpg"                    
                    }
                })  
                
            }

            Write-Verbose ("Creating runspace pool and session states")
            $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
            $runspaceHash.runspacepool = [runspacefactory]::CreateRunspacePool(1, $maxConcurrentJobs, $sessionstate, $Host)
            $runspaceHash.runspacepool.Open()   

            ForEach ($Computer in $selectedItems) {
                $uiHash.Listview.Items.EditItem($Computer)
                $computer.Status = "Pending SEP version checking"
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh() 
                #Create the powershell instance and supply the scriptblock with the other parameters 
                $powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($computer).AddArgument($uiHash).AddArgument($Path)
           
                #Add the runspace into the powershell instance
                $powershell.RunspacePool = $runspaceHash.runspacepool
           
                #Create a temporary collection for each runspace
                $temp = "" | Select-Object PowerShell,Runspace,Computer
                $Temp.Computer = $Computer.computer
                $temp.PowerShell = $powershell
           
                #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
                $temp.Runspace = $powershell.BeginInvoke()
                Write-Verbose ("Adding {0} collection" -f $temp.Computer)
                $jobs.Add($temp) | Out-Null                
            }
            #endregion

        } 
             Else {
                $uiHash.StatusTextBox.Foreground = "Red"
                $uiHash.StatusTextBox.Text = "No server/s selected!"
        }
        
}

Function Check-Spectre{
 Write-Debug ("ComboBox {0}" -f $uiHash.RunOptionComboBox.Text)
        $selectedItems = $uiHash.Listview.SelectedItems
        If ($selectedItems.Count -gt 0) {
        $uiHash.ProgressBar.Maximum = $selectedItems.count
        $uiHash.Listview.ItemsSource | ForEach {$_.Status = $Null; $_.notes = $Null} 
        #region CheckSpectre
            $uiHash.RunButton.IsEnabled = $False
            $uiHash.StartImage.Source = "$pwd\Images\Start_locked.jpg"
            $uiHash.CancelButton.IsEnabled = $True
            $uiHash.CancelImage.Source = "$pwd\Images\Stop.jpg"            
            $uiHash.StatusTextBox.Foreground = "Black"
            $uiHash.StatusTextBox.Text = "Getting Spectre Meltdown details..."            
            $uiHash.StartTime = (Get-Date)
            
            [Float]$uiHash.ProgressBar.Value = 0
            $scriptBlock = {
                Param (
                    $Computer,
                    $uiHash,
                    $Path
                )               
                $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                    $uiHash.Listview.Items.EditItem($Computer)
                    $computer.Status = "Getting Spectre Meltdown details..."
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh() 
                })                
                Set-Location $Path
                $Connection = (Test-Connection -ComputerName $Computer.computer -Count 1 -Quiet)
                    $server = $Computer.computer
                    $spectre = Invoke-Command -cn $server -ScriptBlock {
                       Try{
                        $Override = (Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management'| select "FeatureSettingsOverride")."FeatureSettingsOverride"
                        $Mask = (Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management' | select "FeatureSettingsOverrideMask")."FeatureSettingsOverrideMask"
                          if ($Override -eq 0 -and $Mask -eq 3){
                            return 1
                          }Else{$Status = "Spectre Meltdown configuration not found!"}
                            return 0
                       }Catch {
                       return $Error | Out-String
                       }}
                       if($spectre -eq 1){
                       $status = "Completed"
                       $notes = "Spectre Meltdown is configured"}
                       elseif($spectre -eq 0){
                       $status = "Completed"
                       $notes = "Spectre Meltdown configuration NOT found!"}
                       else{$Status = "Failed"; $notes = $spectre}
               $uiHash.ListView.Dispatcher.Invoke("Background",[action]{                    
                    $uiHash.Listview.Items.EditItem($Computer)
                    If ($Connection) {
                        $Computer.Status = $status
                        $computer.notes = $notes
                    } ElseIf (-Not $Connection) {
                        $Computer.Status = "Offline"
                    } Else {
                        $Computer.Status = "Failed"
                    } 
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh()
                })
                $uiHash.ProgressBar.Dispatcher.Invoke("Normal",[action]{   
                    $uiHash.ProgressBar.value++  
                })
                $uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
                    #Check to see if find job
                    If ($uiHash.ProgressBar.value -eq $uiHash.ProgressBar.Maximum) { 
                        $End = New-Timespan $uihash.StartTime (Get-Date)                                             
                        $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)
                        $uiHash.RunButton.IsEnabled = $True
                        $uiHash.StartImage.Source = "$pwd\Images\Start.jpg"
                        $uiHash.CancelButton.IsEnabled = $False
                        $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.jpg"                    
                    }
                })  
                
            }

            Write-Verbose ("Creating runspace pool and session states")
            $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
            $runspaceHash.runspacepool = [runspacefactory]::CreateRunspacePool(1, $maxConcurrentJobs, $sessionstate, $Host)
            $runspaceHash.runspacepool.Open()   

            ForEach ($Computer in $selectedItems) {
                $uiHash.Listview.Items.EditItem($Computer)
                $computer.Status = "Pending Spectre Meltdown Check"
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh() 
                #Create the powershell instance and supply the scriptblock with the other parameters 
                $powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($computer).AddArgument($uiHash).AddArgument($Path)
           
                #Add the runspace into the powershell instance
                $powershell.RunspacePool = $runspaceHash.runspacepool
           
                #Create a temporary collection for each runspace
                $temp = "" | Select-Object PowerShell,Runspace,Computer
                $Temp.Computer = $Computer.computer
                $temp.PowerShell = $powershell
           
                #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
                $temp.Runspace = $powershell.BeginInvoke()
                Write-Verbose ("Adding {0} collection" -f $temp.Computer)
                $jobs.Add($temp) | Out-Null                
            }
            #endregion

        } 
             Else {
                $uiHash.StatusTextBox.Foreground = "Red"
                $uiHash.StatusTextBox.Text = "No server/s selected!"
        }

}

Function Check-Exchange{
 Write-Debug ("ComboBox {0}" -f $uiHash.RunOptionComboBox.Text)
        $selectedItems = $uiHash.Listview.SelectedItems
        If ($selectedItems.Count -gt 0) {
        $uiHash.ProgressBar.Maximum = $selectedItems.count
        $uiHash.Listview.ItemsSource | ForEach {$_.Status = $Null; $_.notes = $Null} 
        #region CheckExchange
            $uiHash.RunButton.IsEnabled = $False
            $uiHash.StartImage.Source = "$pwd\Images\Start_locked.jpg"
            $uiHash.CancelButton.IsEnabled = $True
            $uiHash.CancelImage.Source = "$pwd\Images\Stop.jpg"            
            $uiHash.StatusTextBox.Foreground = "Black"
            $uiHash.StatusTextBox.Text = "Checking Exchange Services..."            
            $uiHash.StartTime = (Get-Date)
            
            [Float]$uiHash.ProgressBar.Value = 0
            $scriptBlock = {
                Param (
                    $Computer,
                    $uiHash,
                    $Path
                )               
                $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                    $uiHash.Listview.Items.EditItem($Computer)
                    $computer.Status = "Checking Exchange Services..."
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh() 
                })                
                Set-Location $Path
                $Connection = (Test-Connection -ComputerName $Computer.computer -Count 1 -Quiet)
                    $server = $Computer.computer
                    try{
                        $errcnt = 0
                        $Exservices = Get-Service "MSExchange*" -cn $server | select name, Status | Out-string
                        if($Exservices){
                              foreach ($servies in $Exservices) {
                                if($servies -like "*Stopped*"){$errcnt += 1}
                              }
                              if(!$errcnt -gt 0) {
                              $status = "Completed"
                              $notes = "All Exchange Services are running."
                              }else{
                              $status = "Completed"
                              $notes = "$errcnt Exchange Services are NOT running! `n
                              $Exservices"
                              }
                        }Else{
                        $status = "Completed"
                        $notes = "No Exchange Services found!"}
                    }Catch{
                        $Status = "Failed"
                        $Notes = $Error
                    }
               $uiHash.ListView.Dispatcher.Invoke("Background",[action]{                    
                    $uiHash.Listview.Items.EditItem($Computer)
                    If ($Connection) {
                        $Computer.Status = $status
                        $Computer.Notes = $notes
                    } ElseIf (-Not $Connection) {
                        $Computer.Status = "Offline"
                    } Else {
                        $Computer.Status = "Failed"
                    } 
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh()
                })
                $uiHash.ProgressBar.Dispatcher.Invoke("Normal",[action]{   
                    $uiHash.ProgressBar.value++  
                })
                $uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
                    #Check to see if find job
                    If ($uiHash.ProgressBar.value -eq $uiHash.ProgressBar.Maximum) { 
                        $End = New-Timespan $uihash.StartTime (Get-Date)                                             
                        $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)
                        $uiHash.RunButton.IsEnabled = $True
                        $uiHash.StartImage.Source = "$pwd\Images\Start.jpg"
                        $uiHash.CancelButton.IsEnabled = $False
                        $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.jpg"                    
                    }
                })  
                
            }

            Write-Verbose ("Creating runspace pool and session states")
            $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
            $runspaceHash.runspacepool = [runspacefactory]::CreateRunspacePool(1, $maxConcurrentJobs, $sessionstate, $Host)
            $runspaceHash.runspacepool.Open()   

            ForEach ($Computer in $selectedItems) {
                $uiHash.Listview.Items.EditItem($Computer)
                $computer.Status = "Pending Exchange Services Check"
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh() 
                #Create the powershell instance and supply the scriptblock with the other parameters 
                $powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($computer).AddArgument($uiHash).AddArgument($Path)
           
                #Add the runspace into the powershell instance
                $powershell.RunspacePool = $runspaceHash.runspacepool
           
                #Create a temporary collection for each runspace
                $temp = "" | Select-Object PowerShell,Runspace,Computer
                $Temp.Computer = $Computer.computer
                $temp.PowerShell = $powershell
           
                #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
                $temp.Runspace = $powershell.BeginInvoke()
                Write-Verbose ("Adding {0} collection" -f $temp.Computer)
                $jobs.Add($temp) | Out-Null                
            }
            #endregion

        } 
             Else {
                $uiHash.StatusTextBox.Foreground = "Red"
                $uiHash.StatusTextBox.Text = "No server/s selected!"
        }

}

Function Check-TSM{
 Write-Debug ("ComboBox {0}" -f $uiHash.RunOptionComboBox.Text)
        $selectedItems = $uiHash.Listview.SelectedItems
        If ($selectedItems.Count -gt 0) {
        $uiHash.ProgressBar.Maximum = $selectedItems.count
        $uiHash.Listview.ItemsSource | ForEach {$_.Status = $Null;$_.notes = $Null} 
        #region CheckTSM
            $uiHash.RunButton.IsEnabled = $False
            $uiHash.StartImage.Source = "$pwd\Images\Start_locked.jpg"
            $uiHash.CancelButton.IsEnabled = $True
            $uiHash.CancelImage.Source = "$pwd\Images\Stop.jpg"            
            $uiHash.StatusTextBox.Foreground = "Black"
            $uiHash.StatusTextBox.Text = "Getting TSM details..."            
            $uiHash.StartTime = (Get-Date)
            
            [Float]$uiHash.ProgressBar.Value = 0
            $scriptBlock = {
                Param (
                    $Computer,
                    $uiHash,
                    $Path
                )               
                $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                    $uiHash.Listview.Items.EditItem($Computer)
                    $computer.Status = "Getting TSM details..."
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh() 
                })                
                Set-Location $Path
                $Connection = (Test-Connection -ComputerName $Computer.computer -Count 1 -Quiet)
                    $server = $Computer.computer
                    try{
                        $tsm = (get-service -Name "TSM*proxy*sched*" -cn $server).Status
                        if($tsm){
                        $status = "Completed"
                        $notes = "TSM Proxy scheduler is $tsm"
                        }Else{
                        $status = "Completed"
                        $notes = "No TSM Services!"}
                    }Catch{
                        $status = "Failed"
                        $notes = $Error | out-string
                    }
               $uiHash.ListView.Dispatcher.Invoke("Background",[action]{                    
                    $uiHash.Listview.Items.EditItem($Computer)
                    If ($Connection) {
                        $Computer.Status = "$status"
                        $Computer.notes = "$notes"
                    } ElseIf (-Not $Connection) {
                        $Computer.Status = "Offline"
                    } Else {
                        $Computer.Status = "Failed"
                    } 
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh()
                })
                $uiHash.ProgressBar.Dispatcher.Invoke("Normal",[action]{   
                    $uiHash.ProgressBar.value++  
                })
                $uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
                    #Check to see if find job
                    If ($uiHash.ProgressBar.value -eq $uiHash.ProgressBar.Maximum) { 
                        $End = New-Timespan $uihash.StartTime (Get-Date)                                             
                        $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)
                        $uiHash.RunButton.IsEnabled = $True
                        $uiHash.StartImage.Source = "$pwd\Images\Start.jpg"
                        $uiHash.CancelButton.IsEnabled = $False
                        $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.jpg"                    
                    }
                })  
                
            }

            Write-Verbose ("Creating runspace pool and session states")
            $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
            $runspaceHash.runspacepool = [runspacefactory]::CreateRunspacePool(1, $maxConcurrentJobs, $sessionstate, $Host)
            $runspaceHash.runspacepool.Open()   

            ForEach ($Computer in $selectedItems) {
                $uiHash.Listview.Items.EditItem($Computer)
                $computer.Status = "Pending TSM Check"
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh() 
                #Create the powershell instance and supply the scriptblock with the other parameters 
                $powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($computer).AddArgument($uiHash).AddArgument($Path)
           
                #Add the runspace into the powershell instance
                $powershell.RunspacePool = $runspaceHash.runspacepool
           
                #Create a temporary collection for each runspace
                $temp = "" | Select-Object PowerShell,Runspace,Computer
                $Temp.Computer = $Computer.computer
                $temp.PowerShell = $powershell
           
                #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
                $temp.Runspace = $powershell.BeginInvoke()
                Write-Verbose ("Adding {0} collection" -f $temp.Computer)
                $jobs.Add($temp) | Out-Null                
            }
            #endregion

        } 
             Else {
                $uiHash.StatusTextBox.Foreground = "Red"
                $uiHash.StatusTextBox.Text = "No server/s selected!"
        }

}

function Run-Spectre{
 Write-Debug ("ComboBox {0}" -f $uiHash.RunOptionComboBox.Text)
        $selectedItems = $uiHash.Listview.SelectedItems
        If ($selectedItems.Count -gt 0) {
        $uiHash.ProgressBar.Maximum = $selectedItems.count
        $uiHash.Listview.ItemsSource | ForEach {$_.Status = $Null; $_.notes = $Null}
#region SpectreMeltdownnCheck
            $uiHash.RunButton.IsEnabled = $False
            $uiHash.StartImage.Source = "$pwd\Images\Start_locked.jpg"
            $uiHash.CancelButton.IsEnabled = $True
            $uiHash.CancelImage.Source = "$pwd\Images\Stop.jpg"            
            $uiHash.StatusTextBox.Foreground = "Black"
            $uiHash.StatusTextBox.Text = "Applying Spectre Meltdown..."            
            $uiHash.StartTime = (Get-Date)
            
            [Float]$uiHash.ProgressBar.Value = 0
            $scriptBlock = {
                Param (
                    $Computer,
                    $uiHash,
                    $Path
                )               
                $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                    $uiHash.Listview.Items.EditItem($Computer)
                    $computer.Status = "Applying Spectre Meltdown..."
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh() 
                })                
                Set-Location $Path
                $Connection = (Test-Connection -ComputerName $Computer.computer -Count 1 -Quiet)
                        $server = $Computer.computer
                        
                       #$Flashversion = (Get-ItemProperty  -ComputerName $server 'HKLM:\SOFTWARE\Macromedia\FlashPlayerPlugin')."Version"
                       #if ($Flashversion){
                       #$resout = "Adobe Flash Version: $Flashversion"
                       #Else{
                       #$resout = "No software installed!"}

                        $spectre = Invoke-Command -ComputerName $server -ScriptBlock {
                         try{
                         reg add "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management" /v FeatureSettingsOverride /t REG_DWORD /d 0 /f
                         reg add "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management" /v FeatureSettingsOverrideMask /t REG_DWORD /d 3 /f
                         }Catch{ 
                         $_.exception.message             
                         return 0
                         }
                        } | out-string

                         if ($spectre -ne 0){
                         $status = "Completed"
                         $notes = "$spectre Restart the computer for the changes to take effect"
                         }Else{
                         $status = "Failed"
                         $notes = $spectre
                         }

                $uiHash.ListView.Dispatcher.Invoke("Background",[action]{                    
                    $uiHash.Listview.Items.EditItem($Computer)
                    If ($Connection) {
                        $Computer.Status = $status
                        $computer.notes =  $notes
                    } ElseIf (-Not $Connection) {
                        $Computer.Status = "Offline"
                    } Else {
                        $Computer.Status = "Unknown"
                    } 
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh()
                })
                $uiHash.ProgressBar.Dispatcher.Invoke("Normal",[action]{   
                    $uiHash.ProgressBar.value++  
                })
                $uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
                    #Check to see if find job
                    If ($uiHash.ProgressBar.value -eq $uiHash.ProgressBar.Maximum) { 
                        $End = New-Timespan $uihash.StartTime (Get-Date)                                             
                        $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)
                        $uiHash.RunButton.IsEnabled = $True
                        $uiHash.StartImage.Source = "$pwd\Images\Start.jpg"
                        $uiHash.CancelButton.IsEnabled = $False
                        $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.jpg"                    
                    }
                })  
            }
            Write-Verbose ("Creating runspace pool and session states")
            $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
            $runspaceHash.runspacepool = [runspacefactory]::CreateRunspacePool(1, $maxConcurrentJobs, $sessionstate, $Host)
            $runspaceHash.runspacepool.Open()   

            ForEach ($Computer in $selectedItems) {
                $uiHash.Listview.Items.EditItem($Computer)
                $computer.Status = "Pending Spectre Meltdown configuration"
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh() 
                #Create the powershell instance and supply the scriptblock with the other parameters 
                $powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($computer).AddArgument($uiHash).AddArgument($Path)
           
                #Add the runspace into the powershell instance
                $powershell.RunspacePool = $runspaceHash.runspacepool
           
                #Create a temporary collection for each runspace
                $temp = "" | Select-Object PowerShell,Runspace,Computer
                $Temp.Computer = $Computer.computer
                $temp.PowerShell = $powershell
           
                #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
                $temp.Runspace = $powershell.BeginInvoke()
                Write-Verbose ("Adding {0} collection" -f $temp.Computer)
                $jobs.Add($temp) | Out-Null                
            }#endregion
            } 
             Else {
                $uiHash.StatusTextBox.Foreground = "Red"
                $uiHash.StatusTextBox.Text = "No server/s selected!"
        }
}

function Run-removeTBMR{
 Write-Debug ("ComboBox {0}" -f $uiHash.RunOptionComboBox.Text)
        $selectedItems = $uiHash.Listview.SelectedItems
        If ($selectedItems.Count -gt 0) {
        $uiHash.ProgressBar.Maximum = $selectedItems.count
        $uiHash.Listview.ItemsSource | ForEach {$_.Status = $Null; $_.notes = $Null}
#region SpectreMeltdownnCheck
        #If ((Show-Warning "uninstallTBMR") -eq "Yes") {
            $uiHash.RunButton.IsEnabled = $False
            $uiHash.StartImage.Source = "$pwd\Images\Start_locked.jpg"
            $uiHash.CancelButton.IsEnabled = $True
            $uiHash.CancelImage.Source = "$pwd\Images\Stop.jpg"            
            $uiHash.StatusTextBox.Foreground = "Black"
            $uiHash.StatusTextBox.Text = "Uninstalling TBMR..."            
            $uiHash.StartTime = (Get-Date)
            
            [Float]$uiHash.ProgressBar.Value = 0
            $scriptBlock = {
                Param (
                    $Computer,
                    $uiHash,
                    $Path
                )               
                $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                    $uiHash.Listview.Items.EditItem($Computer)
                    $computer.Status = "Uninstallation of TBMR is running, Please wait..."
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh() 
                })                
                Set-Location $Path
                $Connection = (Test-Connection -ComputerName $Computer.computer -Count 1 -Quiet)
                        $server = $Computer.computer
                        $ScrptBlck = [scriptblock]::Create({
                            try{
                           	$x64 = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
                           	$x32 = 'HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
                           	$TBMRchkr64 = Get-ChildItem -path $x64 | Get-ItemProperty | Where-Object {$_.DisplayName -match 'TBMR*'}
                           	$TBMRchkr32 = Get-ChildItem -path $x32 | Get-ItemProperty | Where-Object {$_.DisplayName -match 'TBMR*'}
                           		if($TBMRchkr64 -ne $null){
                           			$TBMRobj64 = Get-ChildItem -path $x64 | Get-ItemProperty | Where-Object {$_.DisplayName -match 'TBMR*'} | Select-Object -Property  DisplayVersion, UninstallString
                           			$uninst = $TBMRobj64.UninstallString
                                    $TMBRver64 = $TBMRchkr64.DisplayVersion
                           			$uninst = (($uninst -split ' ')[1] -replace '/I','/X') + ' /qn /norestart'
                           			Start-Process msiexec.exe -ArgumentList $uninst -Wait
                           			$retrn64 = 'TBMR '+" $TMBRver64 "+' x64 Successfully Removed!'
                           		}else{$retrn64 = "No TBMR x64 installed!"}
                           		if($TBMRchkr32 -ne $null){
                           			$TBMRobj32 = Get-ChildItem -path $x32 | Get-ItemProperty | Where-Object {$_.DisplayName -match 'TBMR*'} | Select-Object -Property  DisplayVersion, UninstallString
                           			$uninst = $TBMRobj32.UninstallString + ' -silent'
                                    $TMBRver32 = $TBMRchkr32.DisplayVersion
                           			Start-Process -FilePath cmd.exe -ArgumentList '/c', $uninst -Wait
                           			$retrn32 = 'TBMR '+" $TMBRver32 "+' x32 Successfully Removed!'
                           		}else{$retrn32 = "No TBMR x32 installed!"}
                             }Catch{$_.exception.message}	
                             Return "$retrn64 `n $retrn32"
                           })
                        $remotesession = new-pssession -computername $server
                        $getTBMR = Invoke-Command -ScriptBlock $ScrptBlck -Session $remotesession | out-string
                        $remotelastexitcode = invoke-command -ScriptBlock { return $?} -Session $remotesession
                        if ($remotelastexitcode){$status = "Completed"}
                        else{$status = "Failed"}

                $uiHash.ListView.Dispatcher.Invoke("Background",[action]{                    
                    $uiHash.Listview.Items.EditItem($Computer)
                    If ($Connection) {
                        $Computer.Status = $status
                        $computer.notes =  $getTBMR
                    } ElseIf (-Not $Connection) {
                        $Computer.Status = "Offline"
                    } Else {
                        $Computer.Status = "Unknown"
                    } 
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh()
                })
                $uiHash.ProgressBar.Dispatcher.Invoke("Normal",[action]{   
                    $uiHash.ProgressBar.value++  
                })
                $uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
                    #Check to see if find job
                    If ($uiHash.ProgressBar.value -eq $uiHash.ProgressBar.Maximum) { 
                        $End = New-Timespan $uihash.StartTime (Get-Date)                                             
                        $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)
                        $uiHash.RunButton.IsEnabled = $True
                        $uiHash.StartImage.Source = "$pwd\Images\Start.jpg"
                        $uiHash.CancelButton.IsEnabled = $False
                        $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.jpg"                    
                    }
                })  
            }
            Write-Verbose ("Creating runspace pool and session states")
            $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
            $runspaceHash.runspacepool = [runspacefactory]::CreateRunspacePool(1, $maxConcurrentJobs, $sessionstate, $Host)
            $runspaceHash.runspacepool.Open()   

            ForEach ($Computer in $selectedItems) {
                $uiHash.Listview.Items.EditItem($Computer)
                $computer.Status = "Pending Spectre Meltdown configuration"
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh() 
                #Create the powershell instance and supply the scriptblock with the other parameters 
                $powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($computer).AddArgument($uiHash).AddArgument($Path)
           
                #Add the runspace into the powershell instance
                $powershell.RunspacePool = $runspaceHash.runspacepool
           
                #Create a temporary collection for each runspace
                $temp = "" | Select-Object PowerShell,Runspace,Computer
                $Temp.Computer = $Computer.computer
                $temp.PowerShell = $powershell
           
                #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
                $temp.Runspace = $powershell.BeginInvoke()
                Write-Verbose ("Adding {0} collection" -f $temp.Computer)
                $jobs.Add($temp) | Out-Null                
            }
            #endregion
            } 
             Else {
                $uiHash.StatusTextBox.Foreground = "Red"
                $uiHash.StatusTextBox.Text = "No server/s selected!"
        }
}

#start-RunJob function
Function Start-RunJob {    
    Write-Debug ("ComboBox {0}" -f $uiHash.RunOptionComboBox.Text)
    $selectedItems = $uiHash.Listview.SelectedItems
    If ($selectedItems.Count -gt 0) {
        $uiHash.ProgressBar.Maximum = $selectedItems.count
        $uiHash.Listview.ItemsSource | ForEach {$_.Status = $Null; $_.notes = $Null}        
        If ($uiHash.RunOptionComboBox.Text -eq 'Install Patches') {             
            #region Install Patches
            $uiHash.RunButton.IsEnabled = $False
            $uiHash.StartImage.Source = "$pwd\Images\Start_locked.jpg"
            $uiHash.CancelButton.IsEnabled = $True
            $uiHash.CancelImage.Source = "$pwd\Images\Stop.jpg"            
            $uiHash.StatusTextBox.Foreground = "Black"
            $uiHash.StatusTextBox.Text = "Installing Patches for all servers...Please Wait"              
            $uiHash.StartTime = (Get-Date)
            
            [Float]$uiHash.ProgressBar.Value = 0
            $scriptBlock = {
                Param (
                    $Path,
                    $Computer,
                    $installAudit,
                    $uiHash
                )
                $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                    $uiHash.Listview.Items.EditItem($Computer)
                    $computer.Status = "Installing Patches"
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh() 
                })  
                Set-Location $path   
                . .\Scripts\Install-Patches.ps1                
                $clientInstall = @(Install-Patches -Computername $computer.computer)
                $installAudit.AddRange($clientInstall) | Out-Null
                $clientInstalledCount =  @($clientInstall | Where {$_.Status -notmatch "Failed to Install Patch|ERROR"}).Count
                $clientInstalledErrorCount = @($clientInstall | Where {$_.Status -match "Failed to Install Patch|ERROR"}).Count
                $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{
                    $uiHash.Listview.Items.EditItem($Computer)
                    If ($clientInstall[0].Title -eq "NA") {                        
                        $Computer.Installed = 0                        
                    } Else {
                        $Computer.Installed = $clientInstalledCount
                        $Computer.InstallErrors = $clientInstalledErrorCount
                    }
                    $Computer.Status = "Completed"
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh()
                })
                $uiHash.ProgressBar.Dispatcher.Invoke("Normal",[action]{
                    $uiHash.ProgressBar.value++  
                })
                $uiHash.Window.Dispatcher.Invoke("Normal",[action]{
                    #Check to see if find job
                    If ($uiHash.ProgressBar.value -eq $uiHash.ProgressBar.Maximum) {    
                        $End = New-Timespan $uihash.StartTime (Get-Date)                                             
                        $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)                                     
                        $uiHash.RunButton.IsEnabled = $True
                        $uiHash.StartImage.Source = "$pwd\Images\Start.jpg"
                        $uiHash.CancelButton.IsEnabled = $False
                        $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.jpg"                    
                    }
                })                  
            }

            Write-Verbose ("Creating runspace pool and session states")
            $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
            $runspaceHash.runspacepool = [runspacefactory]::CreateRunspacePool(1, $maxConcurrentJobs, $sessionstate, $Host)
            $runspaceHash.runspacepool.Open()  

            ForEach ($Computer in $selectedItems) {
                $uiHash.Listview.Items.EditItem($Computer)
                $computer.Status = "Pending Patch Install"
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh() 
                #Create the powershell instance and supply the scriptblock with the other parameters 
                $powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($Path).AddArgument($computer).AddArgument($installAudit).AddArgument($uiHash)
           
                #Add the runspace into the powershell instance
                $powershell.RunspacePool = $runspaceHash.runspacepool
           
                #Create a temporary collection for each runspace
                $temp = "" | Select-Object PowerShell,Runspace,Computer
                $Temp.Computer = $Computer.computer
                $temp.PowerShell = $powershell
           
                #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
                $temp.Runspace = $powershell.BeginInvoke()
                Write-Verbose ("Adding {0} collection" -f $temp.Computer)
                $jobs.Add($temp) | Out-Null                
            }#endregion  
        } ElseIf ($uiHash.RunOptionComboBox.Text -eq 'Audit Patches') {
            #region Audit Patches
            $uiHash.RunButton.IsEnabled = $False
            $uiHash.StartImage.Source = "$pwd\Images\Start_locked.jpg"
            $uiHash.CancelButton.IsEnabled = $True
            $uiHash.CancelImage.Source = "$pwd\Images\Stop.jpg"            
            $uiHash.StatusTextBox.Foreground = "Black"
            $uiHash.StatusTextBox.Text = "Auditing Patches for all servers...Please Wait"            
            $Global:updatelayout = [Windows.Input.InputEventHandler]{ $uiHash.ProgressBar.UpdateLayout() }
            $uiHash.StartTime = (Get-Date)
            
            [Float]$uiHash.ProgressBar.Value = 0
            $scriptBlock = {
                Param (
                    $Path,
                    $Computer,
                    $updateAudit,
                    $uiHash
                )
                $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                    $uiHash.Listview.Items.EditItem($Computer)
                    $computer.Status = "Auditing Patches"
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh() 
                })  
                Set-Location $path
                . .\Scripts\Get-PendingUpdates.ps1                
                $clientUpdate = @(Get-PendingUpdates -Computer $computer.computer)

                $updateAudit.AddRange($clientUpdate) | Out-Null

                $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{
                    $uiHash.Listview.Items.EditItem($Computer)
                    If ($clientUpdate[0].Title -eq "NA") {                        
                        $Computer.Audited = 0
                        $Computer.Status = "Completed"
                    } ElseIf ($clientUpdate[0].Title -eq "ERROR") {
                        $Computer.Audited = 0
                        $Computer.Status = "Error with Audit"
                    } ElseIf ($clientUpdate[0].Title -eq "OFFLINE") {
                        $Computer.Audited = 0
                        $Computer.Status = "Offline"
                    } Else {
                        $Computer.Audited = $clientUpdate.Count
                        $Computer.Status = "Completed"
                    }
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh()
                })
                $uiHash.ProgressBar.Dispatcher.Invoke("Normal",[action]{
                    $uiHash.ProgressBar.value++ 
                }) 
                    
                $uiHash.Window.Dispatcher.Invoke("Normal",[action]{
                    #Check to see if find job
                    If ($uiHash.ProgressBar.value -eq $uiHash.ProgressBar.Maximum) {    
                        $End = New-Timespan $uihash.StartTime (Get-Date)                                             
                        $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)                                     
                        $uiHash.RunButton.IsEnabled = $True
                        $uiHash.StartImage.Source = "$pwd\Images\Start.jpg"
                        $uiHash.CancelButton.IsEnabled = $False
                        $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.jpg"                    
                    }
                })  
                
            }

            Write-Verbose ("Creating runspace pool and session states")
            $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
            $runspaceHash.runspacepool = [runspacefactory]::CreateRunspacePool(1, $maxConcurrentJobs, $sessionstate, $Host)
            $runspaceHash.runspacepool.Open()    

            ForEach ($Computer in $selectedItems) {
                $uiHash.Listview.Items.EditItem($Computer)
                $computer.Status = "Pending Patch Audit"
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh() 
                #Create the powershell instance and supply the scriptblock with the other parameters 
                $powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($Path).AddArgument($computer).AddArgument($updateAudit).AddArgument($uiHash)
           
                #Add the runspace into the powershell instance
                $powershell.RunspacePool = $runspaceHash.runspacepool
           
                #Create a temporary collection for each runspace
                $temp = "" | Select-Object PowerShell,Runspace,Computer
                $Temp.Computer = $Computer.computer
                $temp.PowerShell = $powershell
           
                #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
                $temp.Runspace = $powershell.BeginInvoke()
                Write-Verbose ("Adding {0} collection" -f $temp.Computer)
                $jobs.Add($temp) | Out-Null                
            }#endregion                 
        } ElseIf ($uiHash.RunOptionComboBox.Text -eq 'Pre-Patch') {
           RunPrePatch              
        }ElseIf ($uiHash.RunOptionComboBox.Text -eq 'Post-Patch') {
           RunPostPatch              
        }ElseIf ($uiHash.RunOptionComboBox.Text -eq 'Reboot Systems') {
            System-Reboot          
        } ElseIf ($uiHash.RunOptionComboBox.Text -eq 'PING/BootUpTime Check') {
            #region LastBootUpTime
            $uiHash.RunButton.IsEnabled = $False
            $uiHash.StartImage.Source = "$pwd\Images\Start_locked.jpg"
            $uiHash.CancelButton.IsEnabled = $True
            $uiHash.CancelImage.Source = "$pwd\Images\Stop.jpg"            
            $uiHash.StatusTextBox.Foreground = "Black"
            $uiHash.StatusTextBox.Text = "Checking server connection..."            
            $uiHash.StartTime = (Get-Date)
            
            [Float]$uiHash.ProgressBar.Value = 0
            $scriptBlock = {
                Param (
                    $Computer,
                    $uiHash,
                    $Path
                )               
                $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                    $uiHash.Listview.Items.EditItem($Computer)
                    $computer.Status = "Checking connection"
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh() 
                })                
                Set-Location $Path
                $Connection = (Test-Connection -ComputerName $Computer.computer -Count 1 -Quiet)
                #$Connection = (New-Object System.Net.Sockets.TCPClient -ArgumentList $env:COMPUTERNAME,3389)
                    Try{
                        $server = $Computer.computer
                        $wmi = gwmi Win32_OperatingSystem -computer $server 
                        $LBTime = $wmi.ConvertToDateTime($wmi.Lastbootuptime) 
                        [TimeSpan]$uptime = New-TimeSpan $LBTime $(get-date)
                        if ($wmi) {
                        $resout = "PING Success : LASTBOOTUPTIME:$LBtime : UPTIME $($uptime.days) Days $($uptime.hours) Hrs $($uptime.minutes) Mins $($uptime.seconds) Secs : RDP Connection is OPEN"
                        } else {
                        $resout = "PING:Success but PSexec not present at the server or WMI error!"}
                        $status = "Completed"
                    }Catch{
                    $status = "Failed"
                    $resout = $_.Exception.Message
                    }
                $uiHash.ListView.Dispatcher.Invoke("Background",[action]{                    
                    $uiHash.Listview.Items.EditItem($Computer)
                    If ($Connection) {
                        $Computer.Status = $status
                        $Computer.Notes = $resout
                    } ElseIf (-Not $Connection) {
                        $Computer.Status = "Offline"
                    } Else {
                        $Computer.Status = "Unknown"
                    } 
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh()
                })
                $uiHash.ProgressBar.Dispatcher.Invoke("Normal",[action]{   
                    $uiHash.ProgressBar.value++  
                })
                $uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
                    #Check to see if find job
                    If ($uiHash.ProgressBar.value -eq $uiHash.ProgressBar.Maximum) { 
                        $End = New-Timespan $uihash.StartTime (Get-Date)                                             
                        $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)
                        $uiHash.RunButton.IsEnabled = $True
                        $uiHash.StartImage.Source = "$pwd\Images\Start.jpg"
                        $uiHash.CancelButton.IsEnabled = $False
                        $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.jpg"                    
                    }
                })  
                
            }

            Write-Verbose ("Creating runspace pool and session states")
            $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
            $runspaceHash.runspacepool = [runspacefactory]::CreateRunspacePool(1, $maxConcurrentJobs, $sessionstate, $Host)
            $runspaceHash.runspacepool.Open()   

            ForEach ($Computer in $selectedItems) {
                $uiHash.Listview.Items.EditItem($Computer)
                $computer.Status = "Pending Network Test"
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh() 
                #Create the powershell instance and supply the scriptblock with the other parameters 
                $powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($computer).AddArgument($uiHash).AddArgument($Path)
           
                #Add the runspace into the powershell instance
                $powershell.RunspacePool = $runspaceHash.runspacepool
           
                #Create a temporary collection for each runspace
                $temp = "" | Select-Object PowerShell,Runspace,Computer
                $Temp.Computer = $Computer.computer
                $temp.PowerShell = $powershell
           
                #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
                $temp.Runspace = $powershell.BeginInvoke()
                Write-Verbose ("Adding {0} collection" -f $temp.Computer)
                $jobs.Add($temp) | Out-Null                
            }
            #endregion           
        } ElseIf ($uiHash.RunOptionComboBox.Text -eq 'Check Pending Reboot') {
            #region Check Pending Reboot
            $uiHash.RunButton.IsEnabled = $False
            $uiHash.StartImage.Source = "$pwd\Images\Start_locked.jpg"
            $uiHash.CancelButton.IsEnabled = $True
            $uiHash.CancelImage.Source = "$pwd\Images\Stop.jpg"            
            $uiHash.StatusTextBox.Foreground = "Black"
            $uiHash.StatusTextBox.Text = "Checking for servers with a pending reboot..."            
            $uiHash.StartTime = (Get-Date)
            
            [Float]$uiHash.ProgressBar.Value = 0
            $scriptBlock = {
                Param (
                    $Computer,
                    $uiHash,
                    $Path
                )               
                $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                    $uiHash.Listview.Items.EditItem($Computer)
                    $computer.Status = "Checking for pending reboot"
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh() 
                })                
                Set-Location $Path
                . .\Scripts\Get-ComputerRebootState.ps1
                $clientRebootRequired = Get-ComputerRebootState -Computer $Computer.computer
                $uiHash.ListView.Dispatcher.Invoke("Background",[action]{                    
                    $uiHash.Listview.Items.EditItem($Computer)
                    If ($clientRebootRequired.RebootRequired -eq $True) {
                        $Computer.Status = "Completed"
                        $Computer.Notes = "Reboot Required"
                    } ElseIf ($clientRebootRequired.RebootRequired -eq $False) {
                        $Computer.Status = "Completed"
                        $Computer.Notes = "No Reboot Required"
                    } ElseIf ($clientRebootRequired.RebootRequired -eq "NA") {
                        $Computer.Status = "Failed"
                        $Computer.Notes = "Unable to determine reboot state"
                    } ElseIf ($clientRebootRequired.RebootRequired -eq "Offline") {
                        $Computer.Status = "Offline"
                    }  
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh()
                })
                $uiHash.ProgressBar.Dispatcher.Invoke("Normal",[action]{   
                    $uiHash.ProgressBar.value++  
                })
                $uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
                    #Check to see if find job
                    If ($uiHash.ProgressBar.value -eq $uiHash.ProgressBar.Maximum) { 
                        $End = New-Timespan $uihash.StartTime (Get-Date)                                             
                        $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)
                        $uiHash.RunButton.IsEnabled = $True
                        $uiHash.StartImage.Source = "$pwd\Images\Start.jpg"
                        $uiHash.CancelButton.IsEnabled = $False
                        $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.jpg"                    
                    }
                })  
                
            }

            Write-Verbose ("Creating runspace pool and session states")
            $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
            $runspaceHash.runspacepool = [runspacefactory]::CreateRunspacePool(1, $maxConcurrentJobs, $sessionstate, $Host)
            $runspaceHash.runspacepool.Open()    

            ForEach ($Computer in $selectedItems) {
                $uiHash.Listview.Items.EditItem($Computer)
                $computer.Status = "Pending Reboot Check"
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh() 
                #Create the powershell instance and supply the scriptblock with the other parameters 
                $powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($computer).AddArgument($uiHash).AddArgument($Path)
           
                #Add the runspace into the powershell instance
                $powershell.RunspacePool = $runspaceHash.runspacepool
           
                #Create a temporary collection for each runspace
                $temp = "" | Select-Object PowerShell,Runspace,Computer
                $Temp.Computer = $Computer.computer
                $temp.PowerShell = $powershell
           
                #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
                $temp.Runspace = $powershell.BeginInvoke()
                Write-Verbose ("Adding {0} collection" -f $temp.Computer)
                $jobs.Add($temp) | Out-Null                
            }#endregion                
        } ElseIf ($uiHash.RunOptionComboBox.Text -eq 'Services Check') {
            #region Check Services
            $uiHash.RunButton.IsEnabled = $False
            $uiHash.StartImage.Source = "$pwd\Images\Start_locked.jpg"
            $uiHash.CancelButton.IsEnabled = $True
            $uiHash.CancelImage.Source = "$pwd\Images\Stop.jpg"            
            $uiHash.StatusTextBox.Foreground = "Black"
            $uiHash.StatusTextBox.Text = "Checking for Non-Running Automatic Services..."            
            $uiHash.StartTime = (Get-Date)
            
            [Float]$uiHash.ProgressBar.Value = 0
            $scriptBlock = {
                Param (
                    $Computer,
                    $uiHash,
                    $Path,
                    $servicesAudit
                )               
                $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                    $uiHash.Listview.Items.EditItem($Computer)
                    $computer.Status = "Checking for non-running services set to Auto"
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh() 
                })                
                Clear-Variable queryError -ErrorAction SilentlyContinue
                Set-Location $Path
                If (Test-Connection -ComputerName $computer.computer -Count 1 -Quiet) {
                    Try {
                        $wmi = @{
                            ErrorAction = 'Stop'
                            Computername = $computer.computer
                            Query = "Select __Server,Name,DisplayName,State,StartMode,ExitCode,Status FROM Win32_Service WHERE StartMode='Auto' AND State!='Running'"
                        }
                        $services = @(Get-WmiObject @wmi)
                    } Catch {
                        $queryError = $_.Exception.Message
                    }
                } Else {
                    $queryError = "Offline"
                }
                If ($services.count -gt 0) {
                    $servicesAudit.AddRange($services) | Out-Null
                }
                $uiHash.ListView.Dispatcher.Invoke("Background",[action]{                    
                    $uiHash.Listview.Items.EditItem($Computer)
                    $Computer.Services = $services.count
                    If ($queryError) {
                        $Computer.Status = 'Failed'
                        $Computer.Notes = $queryError
                    } Else {
                        $Computer.Status = 'Completed'
                    }
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh()
                })
                $uiHash.ProgressBar.Dispatcher.Invoke("Normal",[action]{   
                    $uiHash.ProgressBar.value++  
                })
                $uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
                    #Check to see if find job
                    If ($uiHash.ProgressBar.value -eq $uiHash.ProgressBar.Maximum) { 
                        $End = New-Timespan $uihash.StartTime (Get-Date)                                             
                        $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)
                        $uiHash.RunButton.IsEnabled = $True
                        $uiHash.StartImage.Source = "$pwd\Images\Start.jpg"
                        $uiHash.CancelButton.IsEnabled = $False
                        $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.jpg"                    
                    }
                })  
                
            }

            Write-Verbose ("Creating runspace pool and session states")
            $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
            $runspaceHash.runspacepool = [runspacefactory]::CreateRunspacePool(1, $maxConcurrentJobs, $sessionstate, $Host)
            $runspaceHash.runspacepool.Open()    

            ForEach ($Computer in $selectedItems) {
                $uiHash.Listview.Items.EditItem($Computer)
                $computer.Status = "Pending Service Check"
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh() 
                #Create the powershell instance and supply the scriptblock with the other parameters 
                $powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($computer).AddArgument($uiHash).AddArgument($Path).AddArgument($servicesAudit)
           
                #Add the runspace into the powershell instance
                $powershell.RunspacePool = $runspaceHash.runspacepool
           
                #Create a temporary collection for each runspace
                $temp = "" | Select-Object PowerShell,Runspace,Computer
                $Temp.Computer = $Computer.computer
                $temp.PowerShell = $powershell
           
                #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
                $temp.Runspace = $powershell.BeginInvoke()
                Write-Verbose ("Adding {0} collection" -f $temp.Computer)
                $jobs.Add($temp) | Out-Null                
            }#endregion
        } ElseIf ($uiHash.RunOptionComboBox.Text -eq 'Check CHEF Version') {
            #region CHEFverCheck
            $uiHash.RunButton.IsEnabled = $False
            $uiHash.StartImage.Source = "$pwd\Images\Start_locked.jpg"
            $uiHash.CancelButton.IsEnabled = $True
            $uiHash.CancelImage.Source = "$pwd\Images\Stop.jpg"            
            $uiHash.StatusTextBox.Foreground = "Black"
            $uiHash.StatusTextBox.Text = "Checking CHEF version..."            
            $uiHash.StartTime = (Get-Date)
            
            [Float]$uiHash.ProgressBar.Value = 0
            $scriptBlock = {
                Param (
                    $Computer,
                    $uiHash,
                    $Path,
                    $userVal
                )               
                $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                    $uiHash.Listview.Items.EditItem($Computer)
                    $computer.Status = "Checking, Please wait..."
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh()
                    #$userVal = [Microsoft.VisualBasic.Interaction]::InputBox("Enter invoke command here", "Adhoc")
                })                
                Set-Location $Path
                 
              
                $Connection = (Test-Connection -ComputerName $Computer.computer -Count 1 -Quiet)
                        $server = $Computer.computer
                        try{
                            $CHEFversion = (Get-WmiObject -ComputerName $server -Class Win32_Product | select name, version | where-object {$_.name  -Like "Chef*"})."version"
                            $resout = "Current CHEF Version: $CHEFversion"
                            $status = "Completed"
                        }Catch{
                             $status = "Failed"
                             $resout = $_.Exception.Message
                        }

                $uiHash.ListView.Dispatcher.Invoke("Background",[action]{                    
                    $uiHash.Listview.Items.EditItem($Computer)
                    If ($Connection) {
                        $Computer.Status = $status
                        $Computer.Notes =  $resout
                    } ElseIf (-Not $Connection) {
                        $Computer.Status = "Offline"
                    } Else {
                        $Computer.Status = "Unknown"
                    } 
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh()
                })
                $uiHash.ProgressBar.Dispatcher.Invoke("Normal",[action]{   
                    $uiHash.ProgressBar.value++  
                })
                $uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
                    #Check to see if find job
                    If ($uiHash.ProgressBar.value -eq $uiHash.ProgressBar.Maximum) { 
                        $End = New-Timespan $uihash.StartTime (Get-Date)                                             
                        $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)
                        $uiHash.RunButton.IsEnabled = $True
                        $uiHash.StartImage.Source = "$pwd\Images\Start.jpg"
                        $uiHash.CancelButton.IsEnabled = $False
                        $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.jpg"                    
                    }
                })  
                
            }

            Write-Verbose ("Creating runspace pool and session states")
            $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
            $runspaceHash.runspacepool = [runspacefactory]::CreateRunspacePool(1, $maxConcurrentJobs, $sessionstate, $Host)
            $runspaceHash.runspacepool.Open()   

            ForEach ($Computer in $selectedItems) {
                $uiHash.Listview.Items.EditItem($Computer)
                $computer.Status = "Pending CHEF Version Check"
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh() 
                #Create the powershell instance and supply the scriptblock with the other parameters 
                $powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($computer).AddArgument($uiHash).AddArgument($Path)
           
                #Add the runspace into the powershell instance
                $powershell.RunspacePool = $runspaceHash.runspacepool
           
                #Create a temporary collection for each runspace
                $temp = "" | Select-Object PowerShell,Runspace,Computer
                $Temp.Computer = $Computer.computer
                $temp.PowerShell = $powershell
           
                #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
                $temp.Runspace = $powershell.BeginInvoke()
                Write-Verbose ("Adding {0} collection" -f $temp.Computer)
                $jobs.Add($temp) | Out-Null                
            }
            #endregion
        } ElseIf ($uiHash.RunOptionComboBox.Text -eq 'Check C-drive') {
           #region AdobeFlashVersionCheck
           $uiHash.RunButton.IsEnabled = $False
           $uiHash.StartImage.Source = "$pwd\Images\Start_locked.jpg"
           $uiHash.CancelButton.IsEnabled = $True
           $uiHash.CancelImage.Source = "$pwd\Images\Stop.jpg"            
           $uiHash.StatusTextBox.Foreground = "Black"
           $uiHash.StatusTextBox.Text = "Checking C drive free space..."            
           $uiHash.StartTime = (Get-Date)
           
           [Float]$uiHash.ProgressBar.Value = 0
           $scriptBlock = {
               Param (
                   $Computer,
                   $uiHash,
                   $Path
               )               
               $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                   $uiHash.Listview.Items.EditItem($Computer)
                   $computer.Status = "Checking, Please wait..."
                   $uiHash.Listview.Items.CommitEdit()
                   $uiHash.Listview.Items.Refresh() 
               })                
               Set-Location $Path
               $Connection = (Test-Connection -ComputerName $Computer.computer -Count 1 -Quiet)
                       $server = $Computer.computer
                       #$Airname = (Get-WmiObject -Class win32_product -ComputerName $server -Filter "name LIKE 'Adobe%Flash%'")."name"
                       #$Flashversion = (Get-WmiObject -Class win32_product -ComputerName $server -Filter "name LIKE 'Adobe%Flash%'")."version"
                       #$Airversion = (Get-ItemProperty  -ComputerName $server 'HKLM:\SOFTWARE\Macromedia\FlashPlayerPlugin')."Version"
                       #$disk = Get-WmiObject Win32_LogicalDisk -ComputerName $server -Filter "DeviceID='C:'" | Foreach-Object {$_.Size,$_.FreeSpace}
                       #$disk = Invoke-Command -ComputerName $server {Get-PSDrive C} | Select-Object PSComputerName,Used,Free
                       try{
                             $disk = Get-WmiObject Win32_LogicalDisk -ComputerName $server -Filter "DeviceID='C:'" | Select-Object Size,FreeSpace
                             $csize = [math]::Round($disk.Size / 1GB,1)
                             $cfree = [math]::Round($disk.FreeSpace / 1GB,1)
                             $resout = "Total:$csize GB : Free Space:$cfree GB"
                             $status = "Completed"
                       }Catch{
                             $status = "Failed"
                             $resout = $_.Exception.Message
                       }
               $uiHash.ListView.Dispatcher.Invoke("Background",[action]{                    
                   $uiHash.Listview.Items.EditItem($Computer)
                   If ($Connection) {
                       $Computer.Status = $status
                       $Computer.Notes = $resout
                   } ElseIf (-Not $Connection) {
                       $Computer.Status = "Offline"
                   } Else {
                       $Computer.Status = "Unknown"
                   } 
                   $uiHash.Listview.Items.CommitEdit()
                   $uiHash.Listview.Items.Refresh()
               })
               $uiHash.ProgressBar.Dispatcher.Invoke("Normal",[action]{   
                   $uiHash.ProgressBar.value++  
               })
               $uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
                   #Check to see if find job
                   If ($uiHash.ProgressBar.value -eq $uiHash.ProgressBar.Maximum) { 
                       $End = New-Timespan $uihash.StartTime (Get-Date)                                             
                       $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)
                       $uiHash.RunButton.IsEnabled = $True
                       $uiHash.StartImage.Source = "$pwd\Images\Start.jpg"
                       $uiHash.CancelButton.IsEnabled = $False
                       $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.jpg"                    
                   }
               })  
           }
           Write-Verbose ("Creating runspace pool and session states")
           $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
           $runspaceHash.runspacepool = [runspacefactory]::CreateRunspacePool(1, $maxConcurrentJobs, $sessionstate, $Host)
           $runspaceHash.runspacepool.Open()   
      
           ForEach ($Computer in $selectedItems) {
               $uiHash.Listview.Items.EditItem($Computer)
               $computer.Status = "Pending Adobe Air Version Check"
               $uiHash.Listview.Items.CommitEdit()
               $uiHash.Listview.Items.Refresh() 
               #Create the powershell instance and supply the scriptblock with the other parameters 
               $powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($computer).AddArgument($uiHash).AddArgument($Path)
          
               #Add the runspace into the powershell instance
               $powershell.RunspacePool = $runspaceHash.runspacepool
          
               #Create a temporary collection for each runspace
               $temp = "" | Select-Object PowerShell,Runspace,Computer
               $Temp.Computer = $Computer.computer
               $temp.PowerShell = $powershell
          
               #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
               $temp.Runspace = $powershell.BeginInvoke()
               Write-Verbose ("Adding {0} collection" -f $temp.Computer)
               $jobs.Add($temp) | Out-Null                
           }                                   
       }#en-f
       Else {
        $uiHash.StatusTextBox.Foreground = "Red"
        $uiHash.StatusTextBox.Text = "No server/s selected!"
        }  
}}
      


Function Open-FileDialog {
    $dlg = new-object microsoft.win32.OpenFileDialog
    $dlg.DefaultExt = "*.txt"
    $dlg.Filter = "Text Files |*.txt;*.log"    
    $dlg.InitialDirectory = $path
    [void]$dlg.showdialog()
    Write-Output $dlg.FileName
}

Function Open-DomainDialog {
    $domain = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the LDAP path for the Domain or press OK to use the default domain.", 
    "Domain Query", "$(([adsisearcher]'').SearchRoot.distinguishedName)")  
    If (-Not [string]::IsNullOrEmpty($domain)) {
        Write-Output $domain
    }
}

#RMS
#StartSEPMasterService
Function StartSEPMasterService{
 Write-Debug ("ComboBox {0}" -f $uiHash.RunOptionComboBox.Text)
        $selectedItems = $uiHash.Listview.SelectedItems
        If ($selectedItems.Count -gt 0) {
        $uiHash.ProgressBar.Maximum = $selectedItems.count
        $uiHash.Listview.ItemsSource | ForEach {$_.Status = $Null; $_.notes = $Null} 
              $uiHash.RunButton.IsEnabled = $False
            $uiHash.StartImage.Source = "$pwd\Images\Start_locked.jpg"
            $uiHash.CancelButton.IsEnabled = $True
            $uiHash.CancelImage.Source = "$pwd\Images\Stop.jpg"            
            $uiHash.StatusTextBox.Foreground = "Black"
            $uiHash.StatusTextBox.Text = "Starting SEPMaster Service..."            
            $uiHash.StartTime = (Get-Date)
            
            [Float]$uiHash.ProgressBar.Value = 0
            $scriptBlock = {
                Param (
                    $Computer,
                    $uiHash,
                    $Path
                )               
                $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                    $uiHash.Listview.Items.EditItem($Computer)
                    $computer.Status = "Starting SEPMaster Service..."
					$uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh() 
                })                
                Set-Location $Path
                $Connection = (Test-Connection -ComputerName $Computer.computer -Count 1 -Quiet)
                $server = $Computer.computer 
                    
				$status, $resout = Invoke-Command -ComputerName $server -ScriptBlock{
					Try {
                    Start-Service -Name "SepMasterService"
                    $result = $True
                      			
					return	"Completed", $result
						}Catch{return "Failed", $Error[0]}
					}

               $uiHash.ListView.Dispatcher.Invoke("Background",[action]{                    
                    $uiHash.Listview.Items.EditItem($Computer)
                    If ($Connection) {
                        $Computer.status = $status
                        $Computer.notes = $resout
                    } ElseIf (-Not $Connection) {
                        $Computer.Status = "Offline"
                    } Else {
                        $Computer.Status = "Unknown"
                    } 
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh()
                })
                $uiHash.ProgressBar.Dispatcher.Invoke("Normal",[action]{   
                    $uiHash.ProgressBar.value++  
                })
                $uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
                    #Check to see if find job
                    If ($uiHash.ProgressBar.value -eq $uiHash.ProgressBar.Maximum) { 
                        $End = New-Timespan $uihash.StartTime (Get-Date)                                             
                        $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)
                        $uiHash.RunButton.IsEnabled = $True
                        $uiHash.StartImage.Source = "$pwd\Images\Start.jpg"
                        $uiHash.CancelButton.IsEnabled = $False
                        $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.jpg"                    
                    }
                })  
                
            }

            Write-Verbose ("Creating runspace pool and session states")
            $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
            $runspaceHash.runspacepool = [runspacefactory]::CreateRunspacePool(1, $maxConcurrentJobs, $sessionstate, $Host)
            $runspaceHash.runspacepool.Open()   

            ForEach ($Computer in $selectedItems) {
                $uiHash.Listview.Items.EditItem($Computer)
                $computer.Status = "Pending SEPMaster Service Start.."
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh() 
                #Create the powershell instance and supply the scriptblock with the other parameters 
                $powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($computer).AddArgument($uiHash).AddArgument($Path)
           
                #Add the runspace into the powershell instance
                $powershell.RunspacePool = $runspaceHash.runspacepool
           
                #Create a temporary collection for each runspace
                $temp = "" | Select-Object PowerShell,Runspace,Computer
                $Temp.Computer = $Computer.computer
                $temp.PowerShell = $powershell
           
                #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
                $temp.Runspace = $powershell.BeginInvoke()
                Write-Verbose ("Adding {0} collection" -f $temp.Computer)
                $jobs.Add($temp) | Out-Null                
            }
            #endregion

        } 
             Else {
                $uiHash.StatusTextBox.Foreground = "Red"
                $uiHash.StatusTextBox.Text = "No server/s selected!"
        }
        
}


#SEPHealthCheck
Function SEPHealthCheck{
 Write-Debug ("ComboBox {0}" -f $uiHash.RunOptionComboBox.Text)
        $selectedItems = $uiHash.Listview.SelectedItems
        If ($selectedItems.Count -gt 0) {
        $uiHash.ProgressBar.Maximum = $selectedItems.count
        $uiHash.Listview.ItemsSource | ForEach {$_.Status = $Null; $_.notes = $Null} 
	#RegionStart
            $uiHash.RunButton.IsEnabled = $False
            $uiHash.StartImage.Source = "$pwd\Images\Start_locked.jpg"
            $uiHash.CancelButton.IsEnabled = $True
            $uiHash.CancelImage.Source = "$pwd\Images\Stop.jpg"            
            $uiHash.StatusTextBox.Foreground = "Black"
            $uiHash.StatusTextBox.Text = "Running SEP Health Check..."            
            $uiHash.StartTime = (Get-Date)
            
            [Float]$uiHash.ProgressBar.Value = 0
            $scriptBlock = {
                Param (
                    $Computer,
                    $uiHash,
                    $Path
                )               
                $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                    $uiHash.Listview.Items.EditItem($Computer)
                    $computer.Status = "Running"
					$computer.Notes = "SEP Health Check in progress... please wait."
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh() 
                })                
                Set-Location $Path
                $Connection = (Test-Connection -ComputerName $Computer.computer -Count 1 -Quiet)
                $server = $Computer.computer 
                    
				$status, $resout = Invoke-Command -ComputerName $server -ScriptBlock{
					Try{
			
                        $SEPVer = (Get-ItemProperty 'HKLM:\SOFTWARE\Symantec\Symantec Endpoint Protection\CurrentVersion\' | select "PRODUCTVERSION")."PRODUCTVERSION"
                        $SEPDef = (get-ItemProperty 'HKLM:\SOFTWARE\Symantec\Symantec Endpoint Protection\CurrentVersion\public-opstate'  | select LatestVirusDefsDate).LatestVirusDefsDate
                        $rev = (get-ItemProperty 'HKLM:\SOFTWARE\Symantec\Symantec Endpoint Protection\CurrentVersion\public-opstate' | select LatestVirusDefsRevision).LatestVirusDefsRevision
                        $SEPRegPath = (Get-ItemProperty 'HKLM:\SOFTWARE\Symantec\Symantec Endpoint Protection\CurrentVersion\Common Client\PathExpansionMap\INSTALLDIR' | select "Path32")."Path32"
                        $SEPReboot = (get-ItemProperty 'HKLM:\SOFTWARE\Symantec\Symantec Endpoint Protection\CurrentVersion\public-opstate'  | select RebootReason).RebootReason
                        #$SEPDIR = (get-item $SEPRegPath).parent.parent.Fullname
                        $SEPSVCS = get-wmiobject win32_service | where {$_.Displayname -match "Symantec Endpoint"} | Select Name, State, StartMode
                        $SEPSVCSOUT = $SEPSVCS | Out-string

                        $SEPManConnect = (Get-ItemProperty 'HKLM:\SOFTWARE\Wow6432Node\Symantec\Symantec Endpoint Protection\SMC\SYLINK\SyLink' | select "CommunicationStatus")."CommunicationStatus"
                        if($SEPManConnect)
                        {if($SEPManConnect -like "*Couldn't*")

                        {$SEPMResult = "SEPServer: self-managed"}
                        else
                        {
                        $separator = ";"
                        $parts = $SEPManConnect.split($separator)
                        $SEPServer = $parts[1]
                        $SEPMResult = "SEPServer: $SEPServer"}
                        }
                        else{
                        $SEPMResult = "SEPServer: Not Found."
                        }
                        #$SVCCount = 0
                        #$STPcount = 0

                        if($SEPVer -ne $null)
                        {
                        if($SEPVer -ge "14.3"){
                        $SEPDIR = (get-item $SEPRegPath).parent.parent.Fullname
                        #region
                            if($SEPSVCS)
	                        {
	                        foreach($SEPCount in $SEPSVCS){$SVCCount += 1}
	                        foreach($SEPSVCStatus in $SEPSVCS){if($SEPSVCStatus.State -ne "Running"){$STPcount += 1}}
		                        if($STPcount -gt 0){
            		                        if($SEPReboot -eq 0)
                                                {$SEPResult = "Ver. $SEPVer, Def. $SEPDef r$rev `nDir: $SEPDIR `n$SEPMResult `n`nNO SEP Reboot Required. `n`n$STPcount of $SVCCount services STOPPED. `n$SEPSVCSOUT"}
					                        elseif ($SEPReboot -eq 1)
                                                {$SEPResult = "Ver. $SEPVer, Def. $SEPDef r$rev `nDir: $SEPDIR `n$SEPMResult `n`nThreat Remediation. Reboot Required. `n`n $STPcount of $SVCCount services STOPPED. `n$SEPSVCSOUT"}
                                            elseif ($SEPReboot -eq 2)
			                                    {$SEPResult = "Ver. $SEPVer, Def. $SEPDef r$rev `nDir: $SEPDIR `n$SEPMResult `nProduct Patch. Reboot Required! `n $STPcount of $SVCCount services STOPPED. `n$SEPSVCSOUT"}
		                                    elseif ($SEPReboot -eq 3)
			                                    {$SEPResult = "Ver. $SEPVer, Def. $SEPDef r$rev `nDir: $SEPDIR `n$SEPMResult `nContent Download. Reboot Required! `n $STPcount of $SVCCount services STOPPED. `n$SEPSVCSOUT"}
		                                    elseif ($SEPReboot -eq 4)
			                                    {$SEPResult = "Ver. $SEPVer, Def. $SEPDef r$rev `nDir: $SEPDIR `n$SEPMResult `nInstall Completion. Reboot Required! `n $STPcount of $SVCCount services STOPPED. `n$SEPSVCSOUT"}
		                                    elseif ($SEPReboot -eq 5)
			                                    {$SEPResult = "Ver. $SEPVer, Def. $SEPDef r$rev `nDir: $SEPDIR `n$SEPMResult `nSpcRebootCommand. Reboot Required! `n $STPcount of $SVCCount services STOPPED. `n$SEPSVCSOUT"}
		                                    elseif ($SEPReboot -eq 6)
			                                    {$SEPResult = "Ver. $SEPVer, Def. $SEPDef r$rev `nDir: $SEPDIR `n$SEPMResult `nComponent Malfunction. Reboot Required! `n $STPcount of $SVCCount services STOPPED. `n$SEPSVCSOUT"}
		                                    elseif ($SEPReboot -eq 7)
			                                    {$SEPResult = "Ver. $SEPVer, Def. $SEPDef r$rev `nDir: $SEPDIR `n$SEPMResult `nDriver Malfunction. Reboot Required! `n $STPcount of $SVCCount services STOPPED. `n$SEPSVCSOUT"}
		                                    elseif ($SEPReboot -eq 8)
			                                    {$SEPResult = "Ver. $SEPVer, Def. $SEPDef r$rev `nDir: $SEPDIR `n$SEPMResult `nUnknown. Reboot Required! `n $STPcount of $SVCCount services STOPPED. `n$SEPSVCSOUT"}
                                            else{$SEPResult = "Ver. $SEPVer, Def. $SEPDef r$rev `nDir: $SEPDIR `n$SEPMResult `nUnknown Reason. `n $STPcount of $SVCCount services STOPPED. `n$SEPSVCSOUT"}
                                    #$SEPResult = "$STPcount of $SVCCount services STOPPED. `n $SEPSVCSOUT `n$SEPVer, $SEPDef r$rev, $SEPDIR"}
                                            }
		                        else
		                        {
                                foreach($SEPstart in $SEPSVCS)
			                        {	if($SEPstart.StartMode -ne "Auto"){$SEPResult = "Check Starttype! `n $SEPSVCSOUT `n $SEPVer, $SEPDef r$rev, $SEPDIR `n$SEPMResult"}
				                        else
					                        {
					                        if($SEPReboot -eq 0)
                                                {$SEPResult = "Ver: $SEPVer, Def: $SEPDef r$rev `nDir: $SEPDIR `n$SEPMResult `n`nNo SEP Reboot Required. `nAll $SVCCount SEP Services are running."}
					                        elseif ($SEPReboot -eq 1)
                                                {$SEPResult = "Ver. $SEPVer, Def. $SEPDef r$rev `nDir. $SEPDIR `n$SEPMResult `nThreat Remediation. Reboot Required."}
                                            elseif ($SEPReboot -eq 2)
			                                    {$SEPResult = "Ver. $SEPVer, Def. $SEPDef r$rev `nDir. $SEPDIR `n$SEPMResult `nProduct Patch. Reboot Required!"}
		                                    elseif ($SEPReboot -eq 3)
			                                    {$SEPResult = "Ver. $SEPVer, Def. $SEPDef r$rev `nDir. $SEPDIR `n$SEPMResult `nContent Download. Reboot Required!"}
		                                    elseif ($SEPReboot -eq 4)
			                                    {$SEPResult = "Ver. $SEPVer, Def. $SEPDef r$rev `nDir. $SEPDIR `n$SEPMResult `nInstall Completion. Reboot Required!"}
		                                    elseif ($SEPReboot -eq 5)
			                                    {$SEPResult = "Ver. $SEPVer, Def. $SEPDef r$rev `nDir. $SEPDIR `n$SEPMResult `nSpcRebootCommand. Reboot Required!"}
		                                    elseif ($SEPReboot -eq 6)
			                                    {$SEPResult = "Ver. $SEPVer, Def. $SEPDef r$rev `nDir. $SEPDIR `n$SEPMResult `nComponent Malfunction. Reboot Required!"}
		                                    elseif ($SEPReboot -eq 7)
			                                    {$SEPResult = "Ver. $SEPVer, Def. $SEPDef r$rev `nDir. $SEPDIR `n$SEPMResult `nDriver Malfunction. Reboot Required!"}
		                                    elseif ($SEPReboot -eq 8)
			                                    {$SEPResult = "Ver. $SEPVer, Def. $SEPDef r$rev `nDir. $SEPDIR `n$SEPMResult `nUnknown. Reboot Required!"}
                                            else{$SEPResult = "Ver. $SEPVer, Def. $SEPDef r$rev `nDir. $SEPDIR `n$SEPMResult `nUnknown Reason."}
					                        }
                                       
			                        }
		                        }
	                        }
	                        else{$SEPResult = "SEP service(s) are missing! `nVer:$SEPVer, Def:$SEPDef r$rev, Dir:$SEPDIR `n$SEPMResult"}
                        #endregion
                        #end
                        }
                        else{return "Completed","SEP Version is lower than 14.3, check manually or upgrade SEP."}
                        
                        }
                        
                        else
                        {
	                        if($SEPSVCS -ne $null)
	                        {
	                        foreach($SEPCount in $SEPSVCS){$SVCCount += 1}
	                        foreach($SEPSVCStatus in $SEPSVCS){if($SEPSVCStatus.State -ne "Running"){$STPcount += 1}}
		                        if($STPcount -gt 0){$SEPResult = "No SEP registries found! `n $STPcount of $SVCCount services STOPPED. `n $SEPSVCSOUT"}
		                        else
		                        {
		                        foreach($SEPstart in $SEPSVCS)
			                        {	if($SEPstart.StartMode -ne "Auto"){$SEPResult = "No SEP registries found, Incorrect Startup Type, SEP is corrupted! `n $SEPSVCSOUT"}
				                        else
					                        {
					                        #SEP services are running but 
                                            $SEPResult = "No SEP registries found!"
					                        }
			                        }
		                        }
	                        }
	                        else{
                                $TMDeepSec = Get-WmiObject -Class Win32_Product | Where-Object{$_.Name -eq “Trend Micro Deep Security Agent”}
                                if($TMDeepSec){$SEPResult = "SEP is NOT Installed! `nTrend Micro Deep Security Agent is Installed"}
                                else{$SEPResult = "SEP is NOT Installed! `nTrend is NOT Installed"}
                                }
                        }

                        return "Completed", $SEPResult
	
						}Catch{return "Failed", $Error[0]}
					
					}#EndScriptBlock

               $uiHash.ListView.Dispatcher.Invoke("Background",[action]{                    
                    $uiHash.Listview.Items.EditItem($Computer)
                    If ($Connection) {
                        $Computer.status = $status
                        $Computer.notes = $resout
                    } ElseIf (-Not $Connection) {
                        $Computer.Status = "Offline"
						$Computer.Notes = "Check Server"
                    } Else {
                        $Computer.Status = "Unknown"
						$Computer.Notes = "Check Server"
                    } 
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh()
                })
                $uiHash.ProgressBar.Dispatcher.Invoke("Normal",[action]{   
                    $uiHash.ProgressBar.value++  
                })
                $uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
                    #Check to see if find job
                    If ($uiHash.ProgressBar.value -eq $uiHash.ProgressBar.Maximum) { 
                        $End = New-Timespan $uihash.StartTime (Get-Date)                                             
                        $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)
                        $uiHash.RunButton.IsEnabled = $True
                        $uiHash.StartImage.Source = "$pwd\Images\Start.jpg"
                        $uiHash.CancelButton.IsEnabled = $False
                        $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.jpg"                    
                    }
                })  
                
            }

            Write-Verbose ("Creating runspace pool and session states")
            $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
            $runspaceHash.runspacepool = [runspacefactory]::CreateRunspacePool(1, $maxConcurrentJobs, $sessionstate, $Host)
            $runspaceHash.runspacepool.Open()   

            ForEach ($Computer in $selectedItems) {
                $uiHash.Listview.Items.EditItem($Computer)
                $computer.Status = "Pending..."
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh() 
                #Create the powershell instance and supply the scriptblock with the other parameters 
                $powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($computer).AddArgument($uiHash).AddArgument($Path)
           
                #Add the runspace into the powershell instance
                $powershell.RunspacePool = $runspaceHash.runspacepool
           
                #Create a temporary collection for each runspace
                $temp = "" | Select-Object PowerShell,Runspace,Computer
                $Temp.Computer = $Computer.computer
                $temp.PowerShell = $powershell
           
                #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
                $temp.Runspace = $powershell.BeginInvoke()
                Write-Verbose ("Adding {0} collection" -f $temp.Computer)
                $jobs.Add($temp) | Out-Null                
            }
		#Endregion
        } 
             Else {
                $uiHash.StatusTextBox.Foreground = "Red"
                $uiHash.StatusTextBox.Text = "No server/s selected!"
        }
        
}


#RepairSEP
Function RepairSEP{
 Write-Debug ("ComboBox {0}" -f $uiHash.RunOptionComboBox.Text)
        $selectedItems = $uiHash.Listview.SelectedItems
        If ($selectedItems.Count -gt 0) {
        $uiHash.ProgressBar.Maximum = $selectedItems.count
        $uiHash.Listview.ItemsSource | ForEach {$_.Status = $Null; $_.notes = $Null} 
	#RegionStart
            $uiHash.RunButton.IsEnabled = $False
            $uiHash.StartImage.Source = "$pwd\Images\Start_locked.jpg"
            $uiHash.CancelButton.IsEnabled = $True
            $uiHash.CancelImage.Source = "$pwd\Images\Stop.jpg"            
            $uiHash.StatusTextBox.Foreground = "Black"
            $uiHash.StatusTextBox.Text = "Running SEP Repair..."            
            $uiHash.StartTime = (Get-Date)
            
            [Float]$uiHash.ProgressBar.Value = 0
            $scriptBlock = {
                Param (
                    $Computer,
                    $uiHash,
                    $Path
                )               
                $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                    $uiHash.Listview.Items.EditItem($Computer)
                    $computer.Status = "Running"
					$computer.Notes = "SEP repair in progress... please wait."
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh() 
                })                
                Set-Location $Path
                $Connection = (Test-Connection -ComputerName $Computer.computer -Count 1 -Quiet)
                $server = $Computer.computer 
                    
				$status, $resout = Invoke-Command -ComputerName $server -ScriptBlock{
					Try{
					
                Start-Process "c:\windows\system32\msiexec.exe" '/faumsv "C:\ProgramData\Symantec\Symantec Endpoint Protection\CurrentVersion\Data\Cached Installs\Sep64.msi" /q' -Wait

                #SEPHealthCheckAfterRepair Start
                return "Completed"

                #SEPHealthCheckAfterRepair END

						}Catch{return "Failed", $Error[0]}
					
					}#EndScriptBlock

               $uiHash.ListView.Dispatcher.Invoke("Background",[action]{                    
                    $uiHash.Listview.Items.EditItem($Computer)
                    If ($Connection) {
                        $Computer.status = $status
                        $Computer.notes = $resout
                    } ElseIf (-Not $Connection) {
                        $Computer.Status = "Offline"
						$Computer.Notes = "Check Server"
                    } Else {
                        $Computer.Status = "Unknown"
						$Computer.Notes = "Check Server"
                    } 
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh()
                })
                $uiHash.ProgressBar.Dispatcher.Invoke("Normal",[action]{   
                    $uiHash.ProgressBar.value++  
                })
                $uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
                    #Check to see if find job
                    If ($uiHash.ProgressBar.value -eq $uiHash.ProgressBar.Maximum) { 
                        $End = New-Timespan $uihash.StartTime (Get-Date)                                             
                        $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)
                        $uiHash.RunButton.IsEnabled = $True
                        $uiHash.StartImage.Source = "$pwd\Images\Start.jpg"
                        $uiHash.CancelButton.IsEnabled = $False
                        $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.jpg"                    
                    }
                })  
                
            }

            Write-Verbose ("Creating runspace pool and session states")
            $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
            $runspaceHash.runspacepool = [runspacefactory]::CreateRunspacePool(1, $maxConcurrentJobs, $sessionstate, $Host)
            $runspaceHash.runspacepool.Open()   

            ForEach ($Computer in $selectedItems) {
                $uiHash.Listview.Items.EditItem($Computer)
                $computer.Status = "Pending..."
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh() 
                #Create the powershell instance and supply the scriptblock with the other parameters 
                $powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($computer).AddArgument($uiHash).AddArgument($Path)
           
                #Add the runspace into the powershell instance
                $powershell.RunspacePool = $runspaceHash.runspacepool
           
                #Create a temporary collection for each runspace
                $temp = "" | Select-Object PowerShell,Runspace,Computer
                $Temp.Computer = $Computer.computer
                $temp.PowerShell = $powershell
           
                #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
                $temp.Runspace = $powershell.BeginInvoke()
                Write-Verbose ("Adding {0} collection" -f $temp.Computer)
                $jobs.Add($temp) | Out-Null                
            }
		#Endregion
        } 
             Else {
                $uiHash.StatusTextBox.Foreground = "Red"
                $uiHash.StatusTextBox.Text = "No server/s selected!"
        }
        
}


#EndSepScanService
Function EndSepScanService{
 Write-Debug ("ComboBox {0}" -f $uiHash.RunOptionComboBox.Text)
        $selectedItems = $uiHash.Listview.SelectedItems
        If ($selectedItems.Count -gt 0) {
        $uiHash.ProgressBar.Maximum = $selectedItems.count
        $uiHash.Listview.ItemsSource | ForEach {$_.Status = $Null; $_.notes = $Null} 
	#RegionStart
            $uiHash.RunButton.IsEnabled = $False
            $uiHash.StartImage.Source = "$pwd\Images\Start_locked.jpg"
            $uiHash.CancelButton.IsEnabled = $True
            $uiHash.CancelImage.Source = "$pwd\Images\Stop.jpg"            
            $uiHash.StatusTextBox.Foreground = "Black"
            $uiHash.StatusTextBox.Text = "Running EndSepScanService..."            
            $uiHash.StartTime = (Get-Date)
            
            [Float]$uiHash.ProgressBar.Value = 0
            $scriptBlock = {
                Param (
                    $Computer,
                    $uiHash,
                    $Path
                )               
                $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                    $uiHash.Listview.Items.EditItem($Computer)
                    $computer.Status = "Running"
					$computer.Notes = "EndSepScanService in progress... please wait."
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh() 
                })                
                Set-Location $Path
                $Connection = (Test-Connection -ComputerName $Computer.computer -Count 1 -Quiet)
                $server = $Computer.computer 
                    
				$status, $resout = Invoke-Command -ComputerName $server -ScriptBlock{
					Try{
					
            $process = Get-WmiObject -Class Win32_Service -Filter "Name LIKE 'SepScanService'" | Select-Object -ExpandProperty ProcessId
                    if ($process) {
                    Stop-Process $process -Force
                    return "Completed"}
                    else {return "Completed","No SepScanService found."}

						}Catch{return "Failed", $Error[0]}
					
					}#EndScriptBlock

               $uiHash.ListView.Dispatcher.Invoke("Background",[action]{                    
                    $uiHash.Listview.Items.EditItem($Computer)
                    If ($Connection) {
                        $Computer.status = $status
                        $Computer.notes = $resout
                    } ElseIf (-Not $Connection) {
                        $Computer.Status = "Offline"
						$Computer.Notes = "Check Server"
                    } Else {
                        $Computer.Status = "Unknown"
						$Computer.Notes = "Check Server"
                    } 
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh()
                })
                $uiHash.ProgressBar.Dispatcher.Invoke("Normal",[action]{   
                    $uiHash.ProgressBar.value++  
                })
                $uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
                    #Check to see if find job
                    If ($uiHash.ProgressBar.value -eq $uiHash.ProgressBar.Maximum) { 
                        $End = New-Timespan $uihash.StartTime (Get-Date)                                             
                        $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)
                        $uiHash.RunButton.IsEnabled = $True
                        $uiHash.StartImage.Source = "$pwd\Images\Start.jpg"
                        $uiHash.CancelButton.IsEnabled = $False
                        $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.jpg"                    
                    }
                })  
                
            }

            Write-Verbose ("Creating runspace pool and session states")
            $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
            $runspaceHash.runspacepool = [runspacefactory]::CreateRunspacePool(1, $maxConcurrentJobs, $sessionstate, $Host)
            $runspaceHash.runspacepool.Open()   

            ForEach ($Computer in $selectedItems) {
                $uiHash.Listview.Items.EditItem($Computer)
                $computer.Status = "Pending..."
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh() 
                #Create the powershell instance and supply the scriptblock with the other parameters 
                $powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($computer).AddArgument($uiHash).AddArgument($Path)
           
                #Add the runspace into the powershell instance
                $powershell.RunspacePool = $runspaceHash.runspacepool
           
                #Create a temporary collection for each runspace
                $temp = "" | Select-Object PowerShell,Runspace,Computer
                $Temp.Computer = $Computer.computer
                $temp.PowerShell = $powershell
           
                #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
                $temp.Runspace = $powershell.BeginInvoke()
                Write-Verbose ("Adding {0} collection" -f $temp.Computer)
                $jobs.Add($temp) | Out-Null                
            }
		#Endregion
        } 
             Else {
                $uiHash.StatusTextBox.Foreground = "Red"
                $uiHash.StatusTextBox.Text = "No server/s selected!"
        }
        
}


#SEPVirusDefFoldersCheck
Function SEPVirusDefFoldersCheck{
 Write-Debug ("ComboBox {0}" -f $uiHash.RunOptionComboBox.Text)
        $selectedItems = $uiHash.Listview.SelectedItems
        If ($selectedItems.Count -gt 0) {
        $uiHash.ProgressBar.Maximum = $selectedItems.count
        $uiHash.Listview.ItemsSource | ForEach {$_.Status = $Null; $_.notes = $Null} 
	#RegionStart
            $uiHash.RunButton.IsEnabled = $False
            $uiHash.StartImage.Source = "$pwd\Images\Start_locked.jpg"
            $uiHash.CancelButton.IsEnabled = $True
            $uiHash.CancelImage.Source = "$pwd\Images\Stop.jpg"            
            $uiHash.StatusTextBox.Foreground = "Black"
            $uiHash.StatusTextBox.Text = "Checking SEP Virus Definitions Folder(s)..."            
            $uiHash.StartTime = (Get-Date)
            
            [Float]$uiHash.ProgressBar.Value = 0
            $scriptBlock = {
                Param (
                    $Computer,
                    $uiHash,
                    $Path
                )               
                $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                    $uiHash.Listview.Items.EditItem($Computer)
                    $computer.Status = "Running"
					$computer.Notes = "SEP Virus Def Folder Check in progress... please wait."
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh() 
                })                
                Set-Location $Path
                $Connection = (Test-Connection -ComputerName $Computer.computer -Count 1 -Quiet)
                $server = $Computer.computer 
                    
				$status, $resout = Invoke-Command -ComputerName $server -ScriptBlock{
					Try{
			            
                        $Path = "C:\ProgramData\Symantec\Symantec Endpoint Protection\CurrentVersion\Data\Definitions\SDSDefs\"

                        $Dirs = (Get-ChildItem -Path $Path -Directory| Where-Object -Property Name -notmatch "[a-z]").Name
                        
                        if($Dirs -eq $null){return "Completed","No Virus Definitions Folder(s) Found!"}
                        else{
                            $DirsOut = $Dirs | Out-String
                            return "Completed", "$DirsOut"
                            }
                               
	
						}Catch{return "Failed", $Error[0]}
					
					}#EndScriptBlock

               $uiHash.ListView.Dispatcher.Invoke("Background",[action]{                    
                    $uiHash.Listview.Items.EditItem($Computer)
                    If ($Connection) {
                        $Computer.status = $status
                        $Computer.notes = $resout
                    } ElseIf (-Not $Connection) {
                        $Computer.Status = "Offline"
						$Computer.Notes = "Check Server"
                    } Else {
                        $Computer.Status = "Unknown"
						$Computer.Notes = "Check Server"
                    } 
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh()
                })
                $uiHash.ProgressBar.Dispatcher.Invoke("Normal",[action]{   
                    $uiHash.ProgressBar.value++  
                })
                $uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
                    #Check to see if find job
                    If ($uiHash.ProgressBar.value -eq $uiHash.ProgressBar.Maximum) { 
                        $End = New-Timespan $uihash.StartTime (Get-Date)                                             
                        $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)
                        $uiHash.RunButton.IsEnabled = $True
                        $uiHash.StartImage.Source = "$pwd\Images\Start.jpg"
                        $uiHash.CancelButton.IsEnabled = $False
                        $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.jpg"                    
                    }
                })  
                
            }

            Write-Verbose ("Creating runspace pool and session states")
            $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
            $runspaceHash.runspacepool = [runspacefactory]::CreateRunspacePool(1, $maxConcurrentJobs, $sessionstate, $Host)
            $runspaceHash.runspacepool.Open()   

            ForEach ($Computer in $selectedItems) {
                $uiHash.Listview.Items.EditItem($Computer)
                $computer.Status = "Pending..."
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh() 
                #Create the powershell instance and supply the scriptblock with the other parameters 
                $powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($computer).AddArgument($uiHash).AddArgument($Path)
           
                #Add the runspace into the powershell instance
                $powershell.RunspacePool = $runspaceHash.runspacepool
           
                #Create a temporary collection for each runspace
                $temp = "" | Select-Object PowerShell,Runspace,Computer
                $Temp.Computer = $Computer.computer
                $temp.PowerShell = $powershell
           
                #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
                $temp.Runspace = $powershell.BeginInvoke()
                Write-Verbose ("Adding {0} collection" -f $temp.Computer)
                $jobs.Add($temp) | Out-Null                
            }
		#Endregion
        } 
             Else {
                $uiHash.StatusTextBox.Foreground = "Red"
                $uiHash.StatusTextBox.Text = "No server/s selected!"
        }
        
}


#SEPOLDVirusDefDELETE
Function SEPOLDVirusDefDELETE{
 Write-Debug ("ComboBox {0}" -f $uiHash.RunOptionComboBox.Text)
        $selectedItems = $uiHash.Listview.SelectedItems
        If ($selectedItems.Count -gt 0) {
        $uiHash.ProgressBar.Maximum = $selectedItems.count
        $uiHash.Listview.ItemsSource | ForEach {$_.Status = $Null; $_.notes = $Null} 
	#RegionStart
            $uiHash.RunButton.IsEnabled = $False
            $uiHash.StartImage.Source = "$pwd\Images\Start_locked.jpg"
            $uiHash.CancelButton.IsEnabled = $True
            $uiHash.CancelImage.Source = "$pwd\Images\Stop.jpg"            
            $uiHash.StatusTextBox.Foreground = "Black"
            $uiHash.StatusTextBox.Text = "Deleting Old SEP Virus Definition Folder(s)..."            
            $uiHash.StartTime = (Get-Date)
            
            [Float]$uiHash.ProgressBar.Value = 0
            $scriptBlock = {
                Param (
                    $Computer,
                    $uiHash,
                    $Path
                )               
                $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                    $uiHash.Listview.Items.EditItem($Computer)
                    $computer.Status = "Running"
					$computer.Notes = "Old SEP Virus Def Folder(s) Delete in progress... please wait."
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh() 
                })                
                Set-Location $Path
                $Connection = (Test-Connection -ComputerName $Computer.computer -Count 1 -Quiet)
                $server = $Computer.computer 
                    
				$status, $resout = Invoke-Command -ComputerName $server -ScriptBlock{
					Try{
			            
                        $Path = "C:\ProgramData\Symantec\Symantec Endpoint Protection\CurrentVersion\Data\Definitions\SDSDefs\"
                       
                        $Dirs = Get-ChildItem -Path $Path -Directory| Where-Object -Property Name -notmatch "[a-z]"
                        
                        if($Dirs -eq $null){return "Completed","No Virus Definitions Folder(s) Found!"}
                        else{
                            try{
                                $RemoveDirs = $Dirs | Sort-Object LastWriteTime -Desc |Select-Object -Skip 1 | Remove-Item -Recurse
                                return "Completed","Run Check Definitions Folder(s) to verify."
                                }Catch{return "Failed", $Error[0]}
                            }
                                              
	
						}Catch{return "Failed", $Error[0]}
					
					}#EndScriptBlock

               $uiHash.ListView.Dispatcher.Invoke("Background",[action]{                    
                    $uiHash.Listview.Items.EditItem($Computer)
                    If ($Connection) {
                        $Computer.status = $status
                        $Computer.notes = $resout
                    } ElseIf (-Not $Connection) {
                        $Computer.Status = "Offline"
						$Computer.Notes = "Check Server"
                    } Else {
                        $Computer.Status = "Unknown"
						$Computer.Notes = "Check Server"
                    } 
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh()
                })
                $uiHash.ProgressBar.Dispatcher.Invoke("Normal",[action]{   
                    $uiHash.ProgressBar.value++  
                })
                $uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
                    #Check to see if find job
                    If ($uiHash.ProgressBar.value -eq $uiHash.ProgressBar.Maximum) { 
                        $End = New-Timespan $uihash.StartTime (Get-Date)                                             
                        $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)
                        $uiHash.RunButton.IsEnabled = $True
                        $uiHash.StartImage.Source = "$pwd\Images\Start.jpg"
                        $uiHash.CancelButton.IsEnabled = $False
                        $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.jpg"                    
                    }
                })  
                
            }

            Write-Verbose ("Creating runspace pool and session states")
            $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
            $runspaceHash.runspacepool = [runspacefactory]::CreateRunspacePool(1, $maxConcurrentJobs, $sessionstate, $Host)
            $runspaceHash.runspacepool.Open()   

            ForEach ($Computer in $selectedItems) {
                $uiHash.Listview.Items.EditItem($Computer)
                $computer.Status = "Pending..."
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh() 
                #Create the powershell instance and supply the scriptblock with the other parameters 
                $powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($computer).AddArgument($uiHash).AddArgument($Path)
           
                #Add the runspace into the powershell instance
                $powershell.RunspacePool = $runspaceHash.runspacepool
           
                #Create a temporary collection for each runspace
                $temp = "" | Select-Object PowerShell,Runspace,Computer
                $Temp.Computer = $Computer.computer
                $temp.PowerShell = $powershell
           
                #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
                $temp.Runspace = $powershell.BeginInvoke()
                Write-Verbose ("Adding {0} collection" -f $temp.Computer)
                $jobs.Add($temp) | Out-Null                
            }
		#Endregion
        } 
             Else {
                $uiHash.StatusTextBox.Foreground = "Red"
                $uiHash.StatusTextBox.Text = "No server/s selected!"
        }
        
}


#2023-0624
#TrendMicroDSAHealthCheck
Function TrendMicroDSAHealthCheck{
 Write-Debug ("ComboBox {0}" -f $uiHash.RunOptionComboBox.Text)
        $selectedItems = $uiHash.Listview.SelectedItems
        If ($selectedItems.Count -gt 0) {
        $uiHash.ProgressBar.Maximum = $selectedItems.count
        $uiHash.Listview.ItemsSource | ForEach {$_.Status = $Null; $_.notes = $Null} 
	#RegionStart
            $uiHash.RunButton.IsEnabled = $False
            $uiHash.StartImage.Source = "$pwd\Images\Start_locked.jpg"
            $uiHash.CancelButton.IsEnabled = $True
            $uiHash.CancelImage.Source = "$pwd\Images\Stop.jpg"            
            $uiHash.StatusTextBox.Foreground = "Black"
            $uiHash.StatusTextBox.Text = "Running Trend Micro DSA Health Check..."            
            $uiHash.StartTime = (Get-Date)
            
            [Float]$uiHash.ProgressBar.Value = 0
            $scriptBlock = {
                Param (
                    $Computer,
                    $uiHash,
                    $Path
                )               
                $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                    $uiHash.Listview.Items.EditItem($Computer)
                    $computer.Status = "Running"
					$computer.Notes = "Trend Micro DSA Health Check in progress... please wait."
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh() 
                })                
                Set-Location $Path
                $Connection = (Test-Connection -ComputerName $Computer.computer -Count 1 -Quiet)
                $server = $Computer.computer 
                    
				$status, $resout = Invoke-Command -ComputerName $server -ScriptBlock{
					Try{
			
                        $DSA = (Get-ItemProperty 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*' | where {$_.DisplayName -match "Trend Micro Deep Security Agent"}).'DisplayName'
                        $DSAInstallDate = (Get-ItemProperty 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*' | where {$_.DisplayName -match "Trend Micro Deep Security Agent"}).'InstallDate'
                        $DSAVer = (Get-ItemProperty 'HKLM:\SOFTWARE\TrendMicro\Deep Security Agent\' | select "InstalledVersion")."InstalledVersion"
                        $DSADIR = (Get-ItemProperty 'HKLM:\SOFTWARE\TrendMicro\Deep Security Agent\' | select "InstallationFolder")."InstallationFolder"

                        $DSASVCS = get-wmiobject win32_service | where {$_.Displayname -match "Trend Micro"} | Select Name, State, StartMode
                        $DSASVCSOUT = $DSASVCS | Out-string

                        if($DSA -ne $null)
                        {#DSAInstalled
                            if($DSAVer -ne $null)
                            {#RegistryPresent
                            foreach($DSACount in $DSASVCS){$SVCCount += 1}
	                        foreach($DSASVCStatus in $DSASVCS){if($DSASVCStatus.State -ne "Running"){$STPcount += 1}}
                                if($STPcount -gt 0){$DSAResult = "$STPcount of $SVCCount services STOPPED. `n $DSASVCSOUT `nProduct: $DSA `nInstallDate: $DSAInstallDate `nVer: $DSAVer `nDir: $DSADIR" }
                                else
                                {
                                foreach($DSAstart in $DSASVCS)
                                    {#Checking
                                    if($DSAstart.StartMode -ne "Auto"){$DSAResult = "Check Starttype! `n $DSASVCSOUT `nProduct: $DSA `nInstallDate: $DSAInstallDate  `nVer: $DSAVer `nDir: $DSADIR"}
                                    else{
                                    
                                    #trend_actv
                                    $dsaAgentStatus = & "\Program Files\Trend Micro\Deep Security Agent\dsa_query" -c "GetAgentStatus"
                                    foreach($AgentStatus in $dsaAgentStatus)
			                                    {	if($dsaAgentStatus -match "AgentStatus.dsmDN:" -and $dsaAgentStatus -match "AgentStatus.dsmUrl:"){$trend_actv = "trend_activated: True"}
				                                    else{$trend_actv = "trend_activated: False"}
                                                }
                                    
                                    #trend_am
                                    $DSAComponentInfo = & "\Program Files\Trend Micro\Deep Security Agent\dsa_query"-c GetComponentInfo
                                    foreach($DSAcomponent in $DSAComponentInfo)
			                                    {if($DSAComponentInfo -match "Component.AM.mode: on"){$trend_am = "trend_am: True"}
				                                    else{$trend_am = "trend_am: False"}
                                                   }
                                                   
                                    
                                        $DSAResult = "Product: $DSA `nInstallDate: $DSAInstallDate  `nVer: $DSAVer `nDir: $DSADIR `n$SVCCount of $SVCCount Trend Micro Services are running. `n$trend_actv `n$trend_am"
                                        }
                                    
                                    }#Checking
                                }
                            }#RegistryPresent
                            else{$DSAResult = "NO Trend Micro DSA Services Found"}

                        }#DSAInstalled
                        else
                        {#NOTinstalled
                        
                        if($DSASVCS -ne $null)
                            {#NORegistryFoundSection
                            foreach($DSACount in $DSASVCS){$SVCCount += 1}
	                        foreach($DSASVCStatus in $DSASVCS){if($DSASVCStatus.State -ne "Running"){$STPcount += 1}}
                                if($STPcount -gt 0){$DSAResult = "No Trend Micro DSA registries found! `n$STPcount of $SVCCount services STOPPED. `n $DSASVCSOUT" }
                                else
                                {
                                foreach($DSAstart in $DSASVCS)
                                    {#IncorrectStartupType
                                    if($DSAstart.StartMode -ne "Auto"){$DSAResult = "No DSA registries found, Incorrect Startup Type, DSA is corrupted! `n $DSASVCSOUT"}
                                    else{$DSAResult = "No Trend Micro DSA registries found! Install is corrupted."}
                                    }#IncorrectStartupType
                                }
                            }#NORegistryFoundSection
                            else{$DSAResult = "Trend Micro DSA is NOT Installed"}

                        }#NOTinstalled

                        return "Completed", $DSAResult
	
						}Catch{return "Failed", $Error[0]}
					
					}#EndScriptBlock

               $uiHash.ListView.Dispatcher.Invoke("Background",[action]{                    
                    $uiHash.Listview.Items.EditItem($Computer)
                    If ($Connection) {
                        $Computer.status = $status
                        $Computer.notes = $resout
                    } ElseIf (-Not $Connection) {
                        $Computer.Status = "Offline"
						$Computer.Notes = "Check Server"
                    } Else {
                        $Computer.Status = "Unknown"
						$Computer.Notes = "Check Server"
                    } 
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh()
                })
                $uiHash.ProgressBar.Dispatcher.Invoke("Normal",[action]{   
                    $uiHash.ProgressBar.value++  
                })
                $uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
                    #Check to see if find job
                    If ($uiHash.ProgressBar.value -eq $uiHash.ProgressBar.Maximum) { 
                        $End = New-Timespan $uihash.StartTime (Get-Date)                                             
                        $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)
                        $uiHash.RunButton.IsEnabled = $True
                        $uiHash.StartImage.Source = "$pwd\Images\Start.jpg"
                        $uiHash.CancelButton.IsEnabled = $False
                        $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.jpg"                    
                    }
                })  
                
            }

            Write-Verbose ("Creating runspace pool and session states")
            $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
            $runspaceHash.runspacepool = [runspacefactory]::CreateRunspacePool(1, $maxConcurrentJobs, $sessionstate, $Host)
            $runspaceHash.runspacepool.Open()   

            ForEach ($Computer in $selectedItems) {
                $uiHash.Listview.Items.EditItem($Computer)
                $computer.Status = "Pending..."
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh() 
                #Create the powershell instance and supply the scriptblock with the other parameters 
                $powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($computer).AddArgument($uiHash).AddArgument($Path)
           
                #Add the runspace into the powershell instance
                $powershell.RunspacePool = $runspaceHash.runspacepool
           
                #Create a temporary collection for each runspace
                $temp = "" | Select-Object PowerShell,Runspace,Computer
                $Temp.Computer = $Computer.computer
                $temp.PowerShell = $powershell
           
                #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
                $temp.Runspace = $powershell.BeginInvoke()
                Write-Verbose ("Adding {0} collection" -f $temp.Computer)
                $jobs.Add($temp) | Out-Null                
            }
		#Endregion
        } 
             Else {
                $uiHash.StatusTextBox.Foreground = "Red"
                $uiHash.StatusTextBox.Text = "No server/s selected!"
        }
        
}


#2023-0624
#TMEDRProdAgentDeployment
Function TMEDRProdAgentDeployment{
 Write-Debug ("ComboBox {0}" -f $uiHash.RunOptionComboBox.Text)
        $selectedItems = $uiHash.Listview.SelectedItems
        If ($selectedItems.Count -gt 0) {
        $uiHash.ProgressBar.Maximum = $selectedItems.count
        $uiHash.Listview.ItemsSource | ForEach {$_.Status = $Null; $_.notes = $Null} 
	#RegionStart
            $uiHash.RunButton.IsEnabled = $False
            $uiHash.StartImage.Source = "$pwd\Images\Start_locked.jpg"
            $uiHash.CancelButton.IsEnabled = $True
            $uiHash.CancelImage.Source = "$pwd\Images\Stop.jpg"            
            $uiHash.StatusTextBox.Foreground = "Black"
            $uiHash.StatusTextBox.Text = "Running Trend Micro EDR ProdAgentDeployment..."            
            $uiHash.StartTime = (Get-Date)
            
            [Float]$uiHash.ProgressBar.Value = 0
            $scriptBlock = {
                Param (
                    $Computer,
                    $uiHash,
                    $Path
                )               
                $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                    $uiHash.Listview.Items.EditItem($Computer)
                    $computer.Status = "Running"
					$computer.Notes = "Trend Micro EDR ProdAgentDeployment in progress... please wait."
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh() 
                })                
                Set-Location $Path
                $Connection = (Test-Connection -ComputerName $Computer.computer -Count 1 -Quiet)
                $server = $Computer.computer 
                    
				$status, $resout = Invoke-Command -ComputerName $server -ScriptBlock{
					Try{
	
                        $TMEDRProdAgentDeploymentResult = chef-client -o 'recipe[optum_av::install_trend_edr]'
                        $TrendMicroInstall = $TMEDRProdAgentDeploymentResult | Out-String
                        return "Completed", "$TrendMicroInstall"
	
						}Catch{return "Failed", $Error[0]}
					
					}#EndScriptBlock

               $uiHash.ListView.Dispatcher.Invoke("Background",[action]{                    
                    $uiHash.Listview.Items.EditItem($Computer)
                    If ($Connection) {
                        $Computer.status = $status
                        $Computer.notes = $resout
                    } ElseIf (-Not $Connection) {
                        $Computer.Status = "Offline"
						$Computer.Notes = "Check Server"
                    } Else {
                        $Computer.Status = "Unknown"
						$Computer.Notes = "Check Server"
                    } 
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh()
                })
                $uiHash.ProgressBar.Dispatcher.Invoke("Normal",[action]{   
                    $uiHash.ProgressBar.value++  
                })
                $uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
                    #Check to see if find job
                    If ($uiHash.ProgressBar.value -eq $uiHash.ProgressBar.Maximum) { 
                        $End = New-Timespan $uihash.StartTime (Get-Date)                                             
                        $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)
                        $uiHash.RunButton.IsEnabled = $True
                        $uiHash.StartImage.Source = "$pwd\Images\Start.jpg"
                        $uiHash.CancelButton.IsEnabled = $False
                        $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.jpg"                    
                    }
                })  
                
            }

            Write-Verbose ("Creating runspace pool and session states")
            $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
            $runspaceHash.runspacepool = [runspacefactory]::CreateRunspacePool(1, $maxConcurrentJobs, $sessionstate, $Host)
            $runspaceHash.runspacepool.Open()   

            ForEach ($Computer in $selectedItems) {
                $uiHash.Listview.Items.EditItem($Computer)
                $computer.Status = "Pending..."
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh() 
                #Create the powershell instance and supply the scriptblock with the other parameters 
                $powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($computer).AddArgument($uiHash).AddArgument($Path)
           
                #Add the runspace into the powershell instance
                $powershell.RunspacePool = $runspaceHash.runspacepool
           
                #Create a temporary collection for each runspace
                $temp = "" | Select-Object PowerShell,Runspace,Computer
                $Temp.Computer = $Computer.computer
                $temp.PowerShell = $powershell
           
                #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
                $temp.Runspace = $powershell.BeginInvoke()
                Write-Verbose ("Adding {0} collection" -f $temp.Computer)
                $jobs.Add($temp) | Out-Null                
            }
		#Endregion
        } 
             Else {
                $uiHash.StatusTextBox.Foreground = "Red"
                $uiHash.StatusTextBox.Text = "No server/s selected!"
        }
        
}


#Check Patch Log
Function Check-PatchLogTwentyFive{
 Write-Debug ("ComboBox {0}" -f $uiHash.RunOptionComboBox.Text)
        $selectedItems = $uiHash.Listview.SelectedItems
        If ($selectedItems.Count -gt 0) {
        $uiHash.ProgressBar.Maximum = $selectedItems.count
        $uiHash.Listview.ItemsSource | ForEach {$_.Status = $Null; $_.notes = $Null} 
	#region RunChefClient
			$uiHash.RunButton.IsEnabled = $False
            $uiHash.StartImage.Source = "$pwd\Images\Start_locked.jpg"
            $uiHash.CancelButton.IsEnabled = $True
            $uiHash.CancelImage.Source = "$pwd\Images\Stop.jpg"            
            $uiHash.StatusTextBox.Foreground = "Black"
            $uiHash.StatusTextBox.Text = "Checking Patch Log..."            
            $uiHash.StartTime = (Get-Date)
            
            [Float]$uiHash.ProgressBar.Value = 0
            $scriptBlock = {
                Param (
                    $Computer,
                    $uiHash,
                    $Path
                )               
                $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                    $uiHash.Listview.Items.EditItem($Computer)
                    $computer.Status = "Running"
					$computer.Notes = "Patch Log Check in progress... please wait."
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh() 
                })                
                Set-Location $Path
                $Connection = (Test-Connection -ComputerName $Computer.computer -Count 1 -Quiet)
                $server = $Computer.computer 
	
				$status, $resout = Invoke-Command -ComputerName $server -ScriptBlock{
					Try {
                    
                    $PatchLog25 = Get-Content C:\Users\Administrator\patch.log -Tail 25 | Out-String
                    
					return "Completed", "$PatchLog25"


						}Catch{return "Failed", $Error[0]}
					}

               $uiHash.ListView.Dispatcher.Invoke("Background",[action]{                    
                    $uiHash.Listview.Items.EditItem($Computer)
                    If ($Connection) {
                        $Computer.status = $status
                        $Computer.notes = $resout
                    } ElseIf (-Not $Connection) {
                        $Computer.Status = "Offline"
						$Computer.Notes = "Check Server"
                    } Else {
                        $Computer.Status = "Unknown"
						$Computer.Notes = "Check Server"
                    } 
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh()
                })
                $uiHash.ProgressBar.Dispatcher.Invoke("Normal",[action]{   
                    $uiHash.ProgressBar.value++  
                })
                $uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
                    #Check to see if find job
                    If ($uiHash.ProgressBar.value -eq $uiHash.ProgressBar.Maximum) { 
                        $End = New-Timespan $uihash.StartTime (Get-Date)                                             
                        $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)
                        $uiHash.RunButton.IsEnabled = $True
                        $uiHash.StartImage.Source = "$pwd\Images\Start.jpg"
                        $uiHash.CancelButton.IsEnabled = $False
                        $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.jpg"                    
                    }
                })  
                
            }

            Write-Verbose ("Creating runspace pool and session states")
            $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
            $runspaceHash.runspacepool = [runspacefactory]::CreateRunspacePool(1, $maxConcurrentJobs, $sessionstate, $Host)
            $runspaceHash.runspacepool.Open()   

            ForEach ($Computer in $selectedItems) {
                $uiHash.Listview.Items.EditItem($Computer)
                $computer.Status = "Pending..."
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh() 
                #Create the powershell instance and supply the scriptblock with the other parameters 
                $powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($computer).AddArgument($uiHash).AddArgument($Path)
           
                #Add the runspace into the powershell instance
                $powershell.RunspacePool = $runspaceHash.runspacepool
           
                #Create a temporary collection for each runspace
                $temp = "" | Select-Object PowerShell,Runspace,Computer
                $Temp.Computer = $Computer.computer
                $temp.PowerShell = $powershell
           
                #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
                $temp.Runspace = $powershell.BeginInvoke()
                Write-Verbose ("Adding {0} collection" -f $temp.Computer)
                $jobs.Add($temp) | Out-Null                
            }
            #endregion

        } 
             Else {
                $uiHash.StatusTextBox.Foreground = "Red"
                $uiHash.StatusTextBox.Text = "No server/s selected!"
        }
        
}


#RunChefClient
Function RunChefClient{
 Write-Debug ("ComboBox {0}" -f $uiHash.RunOptionComboBox.Text)
        $selectedItems = $uiHash.Listview.SelectedItems
        If ($selectedItems.Count -gt 0) {
        $uiHash.ProgressBar.Maximum = $selectedItems.count
        $uiHash.Listview.ItemsSource | ForEach {$_.Status = $Null; $_.notes = $Null} 
	#region RunChefClient
			$uiHash.RunButton.IsEnabled = $False
            $uiHash.StartImage.Source = "$pwd\Images\Start_locked.jpg"
            $uiHash.CancelButton.IsEnabled = $True
            $uiHash.CancelImage.Source = "$pwd\Images\Stop.jpg"            
            $uiHash.StatusTextBox.Foreground = "Black"
            $uiHash.StatusTextBox.Text = "Running Chef-Client..."            
            $uiHash.StartTime = (Get-Date)
            
            [Float]$uiHash.ProgressBar.Value = 0
            $scriptBlock = {
                Param (
                    $Computer,
                    $uiHash,
                    $Path
                )               
                $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                    $uiHash.Listview.Items.EditItem($Computer)
                    $computer.Status = "Running"
					$computer.Notes = "Chef-Client run in progress... please wait."
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh() 
                })                
                Set-Location $Path
                $Connection = (Test-Connection -ComputerName $Computer.computer -Count 1 -Quiet)
                $server = $Computer.computer 
	
				$status, $resout = Invoke-Command -ComputerName $server -ScriptBlock{
					Try {
                    $runCC = chef-client
                    $CCout = $runCC | ? {$_ -match "resolving cookbooks for run list:"}
					$CCFin = $runCC | ? {$_ -match "Infra Phase complete,"}
					#$CCerr = $runCC | ? {$_ -match "error:"}
					#$CCFatal = $runCC | ? {$_ -match "fatal:"}


					If ($runCC -ne $null) {
				  	$ErrCount = $null
					foreach ($runCCerr in $runCC){
						if($runCCerr -match "error:"){
						$CCerr = $ErrCount += $runCCerr + "`n"
						}
					}}

					If ($runCC -ne $null) {
				  	$FatalCount = $null
					foreach ($runCCfatal in $runCC){
						if($runCCfatal -match "Fatal:"){
						$CCFatal = $FatalCount += $runCCfatal + "`n"
						}
					}}

					return "Completed", "$CCout `n$CCFin `n$CCerr `n$CCFatal"


						}Catch{return "Failed", $Error[0]}
					}

               $uiHash.ListView.Dispatcher.Invoke("Background",[action]{                    
                    $uiHash.Listview.Items.EditItem($Computer)
                    If ($Connection) {
                        $Computer.status = $status
                        $Computer.notes = $resout
                    } ElseIf (-Not $Connection) {
                        $Computer.Status = "Offline"
						$Computer.Notes = "Check Server"
                    } Else {
                        $Computer.Status = "Unknown"
						$Computer.Notes = "Check Server"
                    } 
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh()
                })
                $uiHash.ProgressBar.Dispatcher.Invoke("Normal",[action]{   
                    $uiHash.ProgressBar.value++  
                })
                $uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
                    #Check to see if find job
                    If ($uiHash.ProgressBar.value -eq $uiHash.ProgressBar.Maximum) { 
                        $End = New-Timespan $uihash.StartTime (Get-Date)                                             
                        $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)
                        $uiHash.RunButton.IsEnabled = $True
                        $uiHash.StartImage.Source = "$pwd\Images\Start.jpg"
                        $uiHash.CancelButton.IsEnabled = $False
                        $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.jpg"                    
                    }
                })  
                
            }

            Write-Verbose ("Creating runspace pool and session states")
            $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
            $runspaceHash.runspacepool = [runspacefactory]::CreateRunspacePool(1, $maxConcurrentJobs, $sessionstate, $Host)
            $runspaceHash.runspacepool.Open()   

            ForEach ($Computer in $selectedItems) {
                $uiHash.Listview.Items.EditItem($Computer)
                $computer.Status = "Pending..."
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh() 
                #Create the powershell instance and supply the scriptblock with the other parameters 
                $powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($computer).AddArgument($uiHash).AddArgument($Path)
           
                #Add the runspace into the powershell instance
                $powershell.RunspacePool = $runspaceHash.runspacepool
           
                #Create a temporary collection for each runspace
                $temp = "" | Select-Object PowerShell,Runspace,Computer
                $Temp.Computer = $Computer.computer
                $temp.PowerShell = $powershell
           
                #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
                $temp.Runspace = $powershell.BeginInvoke()
                Write-Verbose ("Adding {0} collection" -f $temp.Computer)
                $jobs.Add($temp) | Out-Null                
            }
            #endregion

        } 
             Else {
                $uiHash.StatusTextBox.Foreground = "Red"
                $uiHash.StatusTextBox.Text = "No server/s selected!"
        }
        
}

Function Check-ASPVer{
 Write-Debug ("ComboBox {0}" -f $uiHash.RunOptionComboBox.Text)
        $selectedItems = $uiHash.Listview.SelectedItems
        If ($selectedItems.Count -gt 0) {
        $uiHash.ProgressBar.Maximum = $selectedItems.count
        $uiHash.Listview.ItemsSource | ForEach {$_.Status = $Null; $_.notes = $Null} 
	#region Check-ASPVer
            $uiHash.RunButton.IsEnabled = $False
            $uiHash.StartImage.Source = "$pwd\Images\Start_locked.jpg"
            $uiHash.CancelButton.IsEnabled = $True
            $uiHash.CancelImage.Source = "$pwd\Images\Stop.jpg"            
            $uiHash.StatusTextBox.Foreground = "Black"
            $uiHash.StatusTextBox.Text = "Checking ASP .Net Core version..."            
            $uiHash.StartTime = (Get-Date)
            
            [Float]$uiHash.ProgressBar.Value = 0
            $scriptBlock = {
                Param (
                    $Computer,
                    $uiHash,
                    $Path
                )               
                $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                    $uiHash.Listview.Items.EditItem($Computer)
                    $computer.Status = "Running"
					$computer.Notes = "Checking ASP .Net Core version... please wait."
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh() 
                })                
                Set-Location $Path
                $Connection = (Test-Connection -ComputerName $Computer.computer -Count 1 -Quiet)
                $server = $Computer.computer 
                    $status, $resout = Invoke-Command -ComputerName $server -ScriptBlock {
                    try{
						 $ASP = (gci ‘HKLM:\SOFTWARE\Microsoft\ASP.NET Core\Shared Framework’ -recurse | gp -name Version | select Version).Version | ft -AutoSize
						 $ASPout = $ASP | Out-String
						 if($ASP -eq $null){return "Completed", "ASP .Net Core is NOT installed."}
                        else{
                         return "Completed", "ASP .Net Core Version: $ASPout"
                         }
						}Catch{return "Failed", $Error[0]}
					} 

               $uiHash.ListView.Dispatcher.Invoke("Background",[action]{                    
                    $uiHash.Listview.Items.EditItem($Computer)
                    If ($Connection) {
                        $Computer.status = $status
                        $Computer.notes = $resout
                    } ElseIf (-Not $Connection) {
                        $Computer.Status = "Offline"
						$Computer.Notes = "Check Server"
                    } Else {
                        $Computer.Status = "Unknown"
						$Computer.Notes = "Check Server"
                    } 
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh()
                })
                $uiHash.ProgressBar.Dispatcher.Invoke("Normal",[action]{   
                    $uiHash.ProgressBar.value++  
                })
                $uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
                    #Check to see if find job
                    If ($uiHash.ProgressBar.value -eq $uiHash.ProgressBar.Maximum) { 
                        $End = New-Timespan $uihash.StartTime (Get-Date)                                             
                        $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)
                        $uiHash.RunButton.IsEnabled = $True
                        $uiHash.StartImage.Source = "$pwd\Images\Start.jpg"
                        $uiHash.CancelButton.IsEnabled = $False
                        $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.jpg"                    
                    }
                })  
                
            }

            Write-Verbose ("Creating runspace pool and session states")
            $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
            $runspaceHash.runspacepool = [runspacefactory]::CreateRunspacePool(1, $maxConcurrentJobs, $sessionstate, $Host)
            $runspaceHash.runspacepool.Open()   

            ForEach ($Computer in $selectedItems) {
                $uiHash.Listview.Items.EditItem($Computer)
                $computer.Status = "Pending..."
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh() 
                #Create the powershell instance and supply the scriptblock with the other parameters 
                $powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($computer).AddArgument($uiHash).AddArgument($Path)
           
                #Add the runspace into the powershell instance
                $powershell.RunspacePool = $runspaceHash.runspacepool
           
                #Create a temporary collection for each runspace
                $temp = "" | Select-Object PowerShell,Runspace,Computer
                $Temp.Computer = $Computer.computer
                $temp.PowerShell = $powershell
           
                #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
                $temp.Runspace = $powershell.BeginInvoke()
                Write-Verbose ("Adding {0} collection" -f $temp.Computer)
                $jobs.Add($temp) | Out-Null                
            }
            #endregion

        } 
             Else {
                $uiHash.StatusTextBox.Foreground = "Red"
                $uiHash.StatusTextBox.Text = "No server/s selected!"
        }
        
}

Function Check-wsusregistry{
 Write-Debug ("ComboBox {0}" -f $uiHash.RunOptionComboBox.Text)
        $selectedItems = $uiHash.Listview.SelectedItems
        If ($selectedItems.Count -gt 0) {
        $uiHash.ProgressBar.Maximum = $selectedItems.count
        $uiHash.Listview.ItemsSource | ForEach {$_.Status = $Null;$_.notes = $Null} 
       #region CheckWSUSregistry
            $uiHash.RunButton.IsEnabled = $False
            $uiHash.StartImage.Source = "$pwd\Images\Start_locked.jpg"
            $uiHash.CancelButton.IsEnabled = $True
            $uiHash.CancelImage.Source = "$pwd\Images\Stop.jpg"            
            $uiHash.StatusTextBox.Foreground = "Black"
            $uiHash.StatusTextBox.Text = "Checking WSUS Registry details..."            
            $uiHash.StartTime = (Get-Date)
            
            [Float]$uiHash.ProgressBar.Value = 0
            $scriptBlock = {
                Param (
                    $Computer,
                    $uiHash,
                    $Path
                )               
                $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                    $uiHash.Listview.Items.EditItem($Computer)
                    $computer.Status = "Running"
					$computer.Notes = "Checking WSUS Registry details... please wait."
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh() 
                })                
                Set-Location $Path
                $Connection = (Test-Connection -ComputerName $Computer.computer -Count 1 -Quiet)
                    $server = $Computer.computer
                    try{
                        $checkWUreg = Get-ItemProperty -Path HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate | Select-Object -Property WUServer,TargetGroup | FL
                        if($checkWUreg){
                        $status = "Completed"
                        $notes = "WSUS Settings:", $checkWUreg | Out-String
                        }Else{
                        $status = "Failed"
                        $notes = "No WSUS Registry Settings Configured!"}
                    }Catch{
                        $status = "Failed"
                        $notes = $($Error[0]) | out-string
                    }
               $uiHash.ListView.Dispatcher.Invoke("Background",[action]{                    
                    $uiHash.Listview.Items.EditItem($Computer)
                    If ($Connection) {
                        $Computer.Status = "$status"
                        $Computer.notes = "$notes"
                    } ElseIf (-Not $Connection) {
                        $Computer.Status = "Offline"
						$Computer.Notes = "Check Server"
                    } Else {
                        $Computer.Status = "Failed"
						$Computer.Notes = "Check Server"
                    } 
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh()
                })
                $uiHash.ProgressBar.Dispatcher.Invoke("Normal",[action]{   
                    $uiHash.ProgressBar.value++  
                })
                $uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
                    #Check to see if find job
                    If ($uiHash.ProgressBar.value -eq $uiHash.ProgressBar.Maximum) { 
                        $End = New-Timespan $uihash.StartTime (Get-Date)                                             
                        $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)
                        $uiHash.RunButton.IsEnabled = $True
                        $uiHash.StartImage.Source = "$pwd\Images\Start.jpg"
                        $uiHash.CancelButton.IsEnabled = $False
                        $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.jpg"                    
                    }
                })  
                
            }

            Write-Verbose ("Creating runspace pool and session states")
            $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
            $runspaceHash.runspacepool = [runspacefactory]::CreateRunspacePool(1, $maxConcurrentJobs, $sessionstate, $Host)
            $runspaceHash.runspacepool.Open()   

            ForEach ($Computer in $selectedItems) {
                $uiHash.Listview.Items.EditItem($Computer)
                $computer.Status = "Pending..."
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh() 
                #Create the powershell instance and supply the scriptblock with the other parameters 
                $powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($computer).AddArgument($uiHash).AddArgument($Path)
           
                #Add the runspace into the powershell instance
                $powershell.RunspacePool = $runspaceHash.runspacepool
           
                #Create a temporary collection for each runspace
                $temp = "" | Select-Object PowerShell,Runspace,Computer
                $Temp.Computer = $Computer.computer
                $temp.PowerShell = $powershell
           
                #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
                $temp.Runspace = $powershell.BeginInvoke()
                Write-Verbose ("Adding {0} collection" -f $temp.Computer)
                $jobs.Add($temp) | Out-Null                
            }
		#endregion

        } 
             Else {
                $uiHash.StatusTextBox.Foreground = "Red"
                $uiHash.StatusTextBox.Text = "No server/s selected!"
        }

}

Function Check-DotNetRuntimeOnOS{
 Write-Debug ("ComboBox {0}" -f $uiHash.RunOptionComboBox.Text)
        $selectedItems = $uiHash.Listview.SelectedItems
        If ($selectedItems.Count -gt 0) {
        $uiHash.ProgressBar.Maximum = $selectedItems.count
        $uiHash.Listview.ItemsSource | ForEach {$_.Status = $Null; $_.notes = $Null} 
	#region CheckDotNetRuntimeOnOS
            $uiHash.RunButton.IsEnabled = $False
            $uiHash.StartImage.Source = "$pwd\Images\Start_locked.jpg"
            $uiHash.CancelButton.IsEnabled = $True
            $uiHash.CancelImage.Source = "$pwd\Images\Stop.jpg"            
            $uiHash.StatusTextBox.Foreground = "Black"
            $uiHash.StatusTextBox.Text = "Checking Latest DotNet Version..."            
            $uiHash.StartTime = (Get-Date)
            
            [Float]$uiHash.ProgressBar.Value = 0
            $scriptBlock = {
                Param (
                    $Computer,
                    $uiHash,
                    $Path
                )               
                $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                    $uiHash.Listview.Items.EditItem($Computer)
                    $computer.Status = "Running"
					$computer.Notes = "Checking Latest DotNet Version... please wait."
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh() 
                })                
                Set-Location $Path
                $Connection = (Test-Connection -ComputerName $Computer.computer -Count 1 -Quiet)
                $server = $Computer.computer 
                    
				$status, $resout = Invoke-Command -ComputerName $server -ScriptBlock{
					Try{
					$Release = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full' | select "Release")."Release"

						If ($Release -ge 528040)
							{$result = "Latest Dotnet Version: 4.8.0"}
							elseif ($Release -ge 461808)
								{$Result = "Latest Dotnet Version: 4.7.2"}
							elseif ($Release -ge 461308)
								{$Result = "Latest Dotnet Version: 4.7.1"}
							elseif ($Release -ge 460798) 
								{$result = "Latest Dotnet Version: 4.7.0"}
							elseif ($Release -ge 394802) 
								{$result = "Latest Dotnet Version: 4.6.2"}
							elseif ($Release -ge 394254) 
								{$result = "Latest Dotnet Version: 4.6.1"}
							elseif ($Release -ge 393295) 
								{$result = "Latest Dotnet Version: 4.6.0"}
							elseif ($Release -ge 379893) 
								{$result = "Latest Dotnet Version: 4.5.2"}
							elseif ($Release -ge 378675) 
								{$result = "Latest Dotnet Version: 4.5.1"}
							elseif ($Release -ge 378389) 
								{$result = "Latest Dotnet Version: 4.5.0"}
						else
							{$result = "No 4.5 or later version detected."}
						
						return	"Completed", $result
						}Catch{return "Failed", $Error[0]}
						
					}

               $uiHash.ListView.Dispatcher.Invoke("Background",[action]{                    
                    $uiHash.Listview.Items.EditItem($Computer)
                    If ($Connection) {
                        $Computer.status = $status
                        $Computer.notes = $resout
                    } ElseIf (-Not $Connection) {
                        $Computer.Status = "Offline"
						$Computer.Notes = "Check Server"
                    } Else {
                        $Computer.Status = "Unknown"
						$Computer.Notes = "Check Server"
                    } 
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh()
                })
                $uiHash.ProgressBar.Dispatcher.Invoke("Normal",[action]{   
                    $uiHash.ProgressBar.value++  
                })
                $uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
                    #Check to see if find job
                    If ($uiHash.ProgressBar.value -eq $uiHash.ProgressBar.Maximum) { 
                        $End = New-Timespan $uihash.StartTime (Get-Date)                                             
                        $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)
                        $uiHash.RunButton.IsEnabled = $True
                        $uiHash.StartImage.Source = "$pwd\Images\Start.jpg"
                        $uiHash.CancelButton.IsEnabled = $False
                        $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.jpg"                    
                    }
                })  
                
            }

            Write-Verbose ("Creating runspace pool and session states")
            $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
            $runspaceHash.runspacepool = [runspacefactory]::CreateRunspacePool(1, $maxConcurrentJobs, $sessionstate, $Host)
            $runspaceHash.runspacepool.Open()   

            ForEach ($Computer in $selectedItems) {
                $uiHash.Listview.Items.EditItem($Computer)
                $computer.Status = "Pending..."
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh() 
                #Create the powershell instance and supply the scriptblock with the other parameters 
                $powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($computer).AddArgument($uiHash).AddArgument($Path)
           
                #Add the runspace into the powershell instance
                $powershell.RunspacePool = $runspaceHash.runspacepool
           
                #Create a temporary collection for each runspace
                $temp = "" | Select-Object PowerShell,Runspace,Computer
                $Temp.Computer = $Computer.computer
                $temp.PowerShell = $powershell
           
                #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
                $temp.Runspace = $powershell.BeginInvoke()
                Write-Verbose ("Adding {0} collection" -f $temp.Computer)
                $jobs.Add($temp) | Out-Null                
            }
            #endregion

        } 
             Else {
                $uiHash.StatusTextBox.Foreground = "Red"
                $uiHash.StatusTextBox.Text = "No server/s selected!"
        }
        
}

Function Check-AllDotNet{
 Write-Debug ("ComboBox {0}" -f $uiHash.RunOptionComboBox.Text)
        $selectedItems = $uiHash.Listview.SelectedItems
        If ($selectedItems.Count -gt 0) {
        $uiHash.ProgressBar.Maximum = $selectedItems.count
        $uiHash.Listview.ItemsSource | ForEach {$_.Status = $Null;$_.notes = $Null} 
        #region Check-AllDotNet
            $uiHash.RunButton.IsEnabled = $False
            $uiHash.StartImage.Source = "$pwd\Images\Start_locked.jpg"
            $uiHash.CancelButton.IsEnabled = $True
            $uiHash.CancelImage.Source = "$pwd\Images\Stop.jpg"            
            $uiHash.StatusTextBox.Foreground = "Black"
            $uiHash.StatusTextBox.Text = "Checking All DotNet Versions..."            
            $uiHash.StartTime = (Get-Date)
            
            [Float]$uiHash.ProgressBar.Value = 0
            $scriptBlock = {
                Param (
                    $Computer,
                    $uiHash,
                    $Path
                )               
                $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                    $uiHash.Listview.Items.EditItem($Computer)
                    $computer.Status = "Running"
					$computer.Notes = "Checking All DotNet Versions... please wait."
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh() 
                })                
                Set-Location $Path
                $Connection = (Test-Connection -ComputerName $Computer.computer -Count 1 -Quiet)
                    $server = $Computer.computer

                    $status, $resout = Invoke-Command -ComputerName $server -ScriptBlock{
                    try{
                        $Release = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full' | select "Release")."Release"

					If ($Release -ge 528040)
						{$result = "4.8.0"}
						elseif ($Release -ge 461808)
							{$Result = "4.7.2"}
						elseif ($Release -ge 461308)
							{$Result = "4.7.1"}
						elseif ($Release -ge 460798) 
							{$result = "4.7.0"}
						elseif ($Release -ge 394802) 
							{$result = "4.6.2"}
						elseif ($Release -ge 394254) 
							{$result = "4.6.1"}
						elseif ($Release -ge 393295) 
							{$result = "4.6.0"}
						elseif ($Release -ge 379893) 
							{$result = "4.5.2"}
						elseif ($Release -ge 378675) 
							{$result = "4.5.1"}
						elseif ($Release -ge 378389) 
							{$result = "4.5.0"}
					else
						{$Result = "No 4.5 or later version detected."}
        #$Result = "No 4.5 or later version detected."
       #Write-Output  $Dotnetchkres


                $DOTNET10 = (Get-ItemProperty 'HKLM:\Software\Microsoft\.NETFramework\Policy\v1.0\3705' | Select-Object Install).Install #REG_SZ equals 1
	            $DOTNET11 = (Get-ItemProperty 'HKLM:\Software\Microsoft\NET Framework Setup\NDP\v1.1.4322' | Select-Object Install).Install #REG_DWORD equals 1
	            $DOTNET20 = (Get-ItemProperty 'HKLM:\Software\Microsoft\NET Framework Setup\NDP\v2.0.50727' | Select-Object Install).Install #REG_DWORD equals 1
	            $DOTNET30 = (Get-ItemProperty 'HKLM:\Software\Microsoft\NET Framework Setup\NDP\v3.0\Setup' | Select-Object InstallSuccess).InstallSuccess #REG_DWORD equals 1
	            $DOTNET35 = (Get-ItemProperty 'HKLM:\Software\Microsoft\NET Framework Setup\NDP\v3.5' | Select-Object Install).Install #REG_DWORD equals 1
	            $DOTNET40ClientProfile = (Get-ItemProperty 'HKLM:\Software\Microsoft\NET Framework Setup\NDP\v4\Client' | Select-Object Install).Install #REG_DWORD equals 1
	            $DOTNET40FullProfile = (Get-ItemProperty 'HKLM:\Software\Microsoft\NET Framework Setup\NDP\v4\Full' | Select	Install).Install #REG_DWORD equals 1
	
	
	                    if($DOTNET10 -eq $null)
	                    	{$DOTNET10Status = "Dotnet 1.0: None"}
	                    else{
	                    	if($DOTNET10 -eq 1)
	                    		{$DOTNET10Status = "Dotnet 1.0: Installed"}
	                    	else{$DOTNET10Status = "Dotnet 1.0: None"}
	                    	}
	                    #Write-Output $DOTNET10Status
	                    	
	                    if($DOTNET11 -eq $null)
	                    	{$DOTNET11Status = "Dotnet 1.1: None"}
	                    else{
	                    	if($DOTNET11 -eq 1)
	                    		{$DOTNET11Status = "Dotnet 1.1: Installed"}
	                    	else
	                    		{$DOTNET11Status = "Dotnet 1.1: None"}
	                    	}
	                   # Write-Output $DOTNET11Status
	                    	
	                    if($DOTNET20 -eq $null)
	                    	{$DOTNET20Status = "Dotnet 2.0: None"}
	                    else{
	                    	if($DOTNET20 -eq 1)
	                    		{$DOTNET20Status = "Dotnet 2.0: Installed"}
	                    	else
	                    		{$DOTNET20Status = "Dotnet 2.0: None"}
	                    	}
	                    #Write-Output $DOTNET20Status
	                    
	                    if($DOTNET30 -eq $null)
	                    	{$DOTNET30Status = "Dotnet 3.0: None"}
	                    else{
	                    	if($DOTNET30 -eq 1)
	                    		{$DOTNET30Status = "Dotnet 3.0: Installed"}
	                    	else
	                    		{$DOTNET30Status = "Dotnet 3.0: None"}
	                    	}
	                   # Write-Output $DOTNET30Status

	                    if($DOTNET35 -eq $null)
	                    	{$DOTNET35Status = "Dotnet 3.5: None"}
	                    else{
	                    	if($DOTNET35 -eq 1)
	                    		{$DOTNET35Status = "Dotnet 3.5: Installed"}
	                    	else
	                    		{$DOTNET35Status = "Dotnet 3.5: None"}
	                    	}
	                   # Write-Output $DOTNET35Status
	                    
	                    if($DOTNET40ClientProfile -eq $null)
	                    	{$DOTNET40ClientProfileStatus = "Dotnet 4.0 Client: None"}
	                    else{
	                    	if($DOTNET40ClientProfile -eq 1)
	                    		{$DOTNET40ClientProfileStatus = "Dotnet 4.0 Client: Installed"}
	                    	else
	                    		{$DOTNET40ClientProfile = "Dotnet 4.0 Client: None"}
	                    	}
	                   # Write-Output $DOTNET40ClientProfileStatus	
	                    
	                    if($DOTNET40FullProfile -eq $null)
	                    	{$DOTNET40FullProfileStatus = "Dotnet 4.0 Full: None"}
	                    else{
	                    	if($DOTNET40FullProfile -eq 1)
	                    		{$DOTNET40FullProfileStatus = "Dotnet 4.0 Full: Installed"}
	                    	else
	                    		{$DOTNET40FullProfileStatus = "Dotnet 4.0 Full: None"}
	                    	}


						}Catch{return "Failed", $Error[0]}

                       return "Completed","Latest DotNet: $Result `n$DOTNET10Status `n$DOTNET11Status `n$DOTNET20Status `n$DOTNET30Status `n$DOTNET35Status `n$DOTNET40ClientProfileStatus `n$DOTNET40FullProfileStatus "
						
					}

                      
               $uiHash.ListView.Dispatcher.Invoke("Background",[action]{                    
                    $uiHash.Listview.Items.EditItem($Computer)
                    If ($Connection) {
                        $Computer.Status = "$status"
                        $Computer.notes =  $resout #"  #$DOTNET10,$DOTNET11,$DOTNET20,$DOTNET30,$DOTNET35,$DOTNET40ClientProfile,$DOTNET40FullProfile"
                    } ElseIf (-Not $Connection) {
                        $Computer.Status = "Offline"
						$Computer.Notes = "Check Server"
                    } Else {
                        $Computer.Status = "Failed"
						$Computer.Notes = "Check Server"
                    } 
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh()
                })
                $uiHash.ProgressBar.Dispatcher.Invoke("Normal",[action]{   
                    $uiHash.ProgressBar.value++  
                })
                $uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
                    #Check to see if find job
                    If ($uiHash.ProgressBar.value -eq $uiHash.ProgressBar.Maximum) { 
                        $End = New-Timespan $uihash.StartTime (Get-Date)                                             
                        $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)
                        $uiHash.RunButton.IsEnabled = $True
                        $uiHash.StartImage.Source = "$pwd\Images\Start.jpg"
                        $uiHash.CancelButton.IsEnabled = $False
                        $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.jpg"                    
                    }
                })  
                
               }                    

            Write-Verbose ("Creating runspace pool and session states")
            $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
            $runspaceHash.runspacepool = [runspacefactory]::CreateRunspacePool(1, $maxConcurrentJobs, $sessionstate, $Host)
            $runspaceHash.runspacepool.Open()   

            ForEach ($Computer in $selectedItems) {
                $uiHash.Listview.Items.EditItem($Computer)
                $computer.Status = "Pending..."
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh() 
                #Create the powershell instance and supply the scriptblock with the other parameters 
                $powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($computer).AddArgument($uiHash).AddArgument($Path)
           
                #Add the runspace into the powershell instance
                $powershell.RunspacePool = $runspaceHash.runspacepool
           
                #Create a temporary collection for each runspace
                $temp = "" | Select-Object PowerShell,Runspace,Computer
                $Temp.Computer = $Computer.computer
                $temp.PowerShell = $powershell
           
                #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
                $temp.Runspace = $powershell.BeginInvoke()
                Write-Verbose ("Adding {0} collection" -f $temp.Computer)
                $jobs.Add($temp) | Out-Null                
            }
            #endregionCheck-AllDotNet

        }
             Else {
                $uiHash.StatusTextBox.Foreground = "Red"
                $uiHash.StatusTextBox.Text = "No server/s selected!"}
   }    
  
Function Check-DotNetAddOns{
 Write-Debug ("ComboBox {0}" -f $uiHash.RunOptionComboBox.Text)
        $selectedItems = $uiHash.Listview.SelectedItems
        If ($selectedItems.Count -gt 0) {
        $uiHash.ProgressBar.Maximum = $selectedItems.count
        $uiHash.Listview.ItemsSource | ForEach {$_.Status = $Null; $_.notes = $Null} 
	#region DotNetAddonsCheck
            $uiHash.RunButton.IsEnabled = $False
            $uiHash.StartImage.Source = "$pwd\Images\Start_locked.jpg"
            $uiHash.CancelButton.IsEnabled = $True
            $uiHash.CancelImage.Source = "$pwd\Images\Stop.jpg"            
            $uiHash.StatusTextBox.Foreground = "Black"
            $uiHash.StatusTextBox.Text = "Checking DotNet AddOns..."            
            $uiHash.StartTime = (Get-Date)
            
            [Float]$uiHash.ProgressBar.Value = 0
            $scriptBlock = {
                Param (
                    $Computer,
                    $uiHash,
                    $Path
                )               
                $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                    $uiHash.Listview.Items.EditItem($Computer)
                    $computer.Status = "Running"
					$computer.Notes = "Checking DotNet AddOns... please wait."
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh() 
                })                
                Set-Location $Path
                $Connection = (Test-Connection -ComputerName $Computer.computer -Count 1 -Quiet)
                $server = $Computer.computer 
                    
				$status, $resout = Invoke-Command -ComputerName $server -ScriptBlock{
					Try{
						$DotNetAddOns = (Get-ItemProperty HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* |  Select-Object DisplayName | where { $_.DisplayName -like "*.NET*"}).DisplayName | FL
						$result = $DotNetAddOns | Out-String
						if($DotNetAddOns -eq $null)
                        {return "Completed", "There are NO DotNet AddOns Installed."}
                        else{
                        return	"Completed", $result
                        }
						}Catch{return "Failed", $Error[0]}
						
					}

               $uiHash.ListView.Dispatcher.Invoke("Background",[action]{                    
                    $uiHash.Listview.Items.EditItem($Computer)
                    If ($Connection) {
                        $Computer.status = $status
                        $Computer.notes = $resout
                    } ElseIf (-Not $Connection) {
                        $Computer.Status = "Offline"
						$Computer.Notes = "Check Server"
                    } Else {
                        $Computer.Status = "Unknown"
						$Computer.Notes = "Check Server"
                    } 
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh()
                })
                $uiHash.ProgressBar.Dispatcher.Invoke("Normal",[action]{   
                    $uiHash.ProgressBar.value++  
                })
                $uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
                    #Check to see if find job
                    If ($uiHash.ProgressBar.value -eq $uiHash.ProgressBar.Maximum) { 
                        $End = New-Timespan $uihash.StartTime (Get-Date)                                             
                        $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)
                        $uiHash.RunButton.IsEnabled = $True
                        $uiHash.StartImage.Source = "$pwd\Images\Start.jpg"
                        $uiHash.CancelButton.IsEnabled = $False
                        $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.jpg"                    
                    }
                })  
                
            }

            Write-Verbose ("Creating runspace pool and session states")
            $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
            $runspaceHash.runspacepool = [runspacefactory]::CreateRunspacePool(1, $maxConcurrentJobs, $sessionstate, $Host)
            $runspaceHash.runspacepool.Open()   

            ForEach ($Computer in $selectedItems) {
                $uiHash.Listview.Items.EditItem($Computer)
                $computer.Status = "Pending..."
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh() 
                #Create the powershell instance and supply the scriptblock with the other parameters 
                $powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($computer).AddArgument($uiHash).AddArgument($Path)
           
                #Add the runspace into the powershell instance
                $powershell.RunspacePool = $runspaceHash.runspacepool
           
                #Create a temporary collection for each runspace
                $temp = "" | Select-Object PowerShell,Runspace,Computer
                $Temp.Computer = $Computer.computer
                $temp.PowerShell = $powershell
           
                #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
                $temp.Runspace = $powershell.BeginInvoke()
                Write-Verbose ("Adding {0} collection" -f $temp.Computer)
                $jobs.Add($temp) | Out-Null                
            }
            #endregion DotNetAddOnsCheck

        } 
             Else {
                $uiHash.StatusTextBox.Foreground = "Red"
                $uiHash.StatusTextBox.Text = "No server/s selected!"
        }
        
}

Function Check-MsOfficeVersion{
 Write-Debug ("ComboBox {0}" -f $uiHash.RunOptionComboBox.Text)
        $selectedItems = $uiHash.Listview.SelectedItems
        If ($selectedItems.Count -gt 0) {
        $uiHash.ProgressBar.Maximum = $selectedItems.count
        $uiHash.Listview.ItemsSource | ForEach {$_.Status = $Null; $_.notes = $Null} 
	#region CheckMsOfficeVersion
            $uiHash.RunButton.IsEnabled = $False
            $uiHash.StartImage.Source = "$pwd\Images\Start_locked.jpg"
            $uiHash.CancelButton.IsEnabled = $True
            $uiHash.CancelImage.Source = "$pwd\Images\Stop.jpg"            
            $uiHash.StatusTextBox.Foreground = "Black"
            $uiHash.StatusTextBox.Text = "Checking MS Office version..."            
            $uiHash.StartTime = (Get-Date)
            
            [Float]$uiHash.ProgressBar.Value = 0
            $scriptBlock = {
                Param (
                    $Computer,
                    $uiHash,
                    $Path
                )               
                $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                    $uiHash.Listview.Items.EditItem($Computer)
                    $computer.Status = "Running"
					$computer.Notes = "Checking MS Office version... please wait."
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh() 
                })                
                Set-Location $Path
                $Connection = (Test-Connection -ComputerName $Computer.computer -Count 1 -Quiet)
                $server = $Computer.computer 
                   # try{
                    $status, $resout = Invoke-Command -ComputerName $server -ScriptBlock {
                          try{
                            $MSOfficeVer = (Get-WmiObject win32_product | where{$_.Name -like "Microsoft Office Professional*"} | select Name).name
                            if($MSOfficeVer -ne $null){
                                   $res = $MSOfficeVer
                                   return "Completed", $res
                            }Else{return "Failed", "No MS Office installed"}
                          }Catch{$Status = "Failed"; $resout = $_.Exception.Message }
                        #} -ErrorVariable e -ErrorAction SilentlyContinue
                        #if($e){$Status = "Failed"; $resout = $e | Out-String}
                   # }Catch{$Status = "Failed"; $resout = $_.Exception.Message }
}
               $uiHash.ListView.Dispatcher.Invoke("Background",[action]{                    
                    $uiHash.Listview.Items.EditItem($Computer)
                    If ($Connection) {
                        $Computer.status = $status
                        $Computer.notes = $resout
                    } ElseIf (-Not $Connection) {
                        $Computer.Status = "Offline"
						$Computer.Notes = "Check Server"
                    } Else {
                        $Computer.Status = "Unknown"
						$Computer.Notes = "Check Server"
                    } 
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh()
                })
                $uiHash.ProgressBar.Dispatcher.Invoke("Normal",[action]{   
                    $uiHash.ProgressBar.value++  
                })
                $uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
                    #Check to see if find job
                    If ($uiHash.ProgressBar.value -eq $uiHash.ProgressBar.Maximum) { 
                        $End = New-Timespan $uihash.StartTime (Get-Date)                                             
                        $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)
                        $uiHash.RunButton.IsEnabled = $True
                        $uiHash.StartImage.Source = "$pwd\Images\Start.jpg"
                        $uiHash.CancelButton.IsEnabled = $False
                        $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.jpg"                    
                    }
                })  
                
            }

            Write-Verbose ("Creating runspace pool and session states")
            $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
            $runspaceHash.runspacepool = [runspacefactory]::CreateRunspacePool(1, $maxConcurrentJobs, $sessionstate, $Host)
            $runspaceHash.runspacepool.Open()   

            ForEach ($Computer in $selectedItems) {
                $uiHash.Listview.Items.EditItem($Computer)
                $computer.Status = "Pending..."
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh() 
                #Create the powershell instance and supply the scriptblock with the other parameters 
                $powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($computer).AddArgument($uiHash).AddArgument($Path)
           
                #Add the runspace into the powershell instance
                $powershell.RunspacePool = $runspaceHash.runspacepool
           
                #Create a temporary collection for each runspace
                $temp = "" | Select-Object PowerShell,Runspace,Computer
                $Temp.Computer = $Computer.computer
                $temp.PowerShell = $powershell
           
                #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
                $temp.Runspace = $powershell.BeginInvoke()
                Write-Verbose ("Adding {0} collection" -f $temp.Computer)
                $jobs.Add($temp) | Out-Null                
            }
            #endregion

        } 
             Else {
                $uiHash.StatusTextBox.Foreground = "Red"
                $uiHash.StatusTextBox.Text = "No server/s selected!"
        }
        
}

Function Check-CBSLogs{
 Write-Debug ("ComboBox {0}" -f $uiHash.RunOptionComboBox.Text)
        $selectedItems = $uiHash.Listview.SelectedItems
        If ($selectedItems.Count -gt 0) {
        $uiHash.ProgressBar.Maximum = $selectedItems.count
        $uiHash.Listview.ItemsSource | ForEach {$_.Status = $Null; $_.notes = $Null} 
        #region CheckCBSLogs
            $uiHash.RunButton.IsEnabled = $False
            $uiHash.StartImage.Source = "$pwd\Images\Start_locked.jpg"
            $uiHash.CancelButton.IsEnabled = $True
            $uiHash.CancelImage.Source = "$pwd\Images\Stop.jpg"            
            $uiHash.StatusTextBox.Foreground = "Black"
            $uiHash.StatusTextBox.Text = "Checking CBS logs..."            
            $uiHash.StartTime = (Get-Date)
            
            [Float]$uiHash.ProgressBar.Value = 0
            $scriptBlock = {
                Param (
                    $Computer,
                    $uiHash,
                    $Path
                )               
                $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                    $uiHash.Listview.Items.EditItem($Computer)
                    $computer.Status = "Running"
					$computer.Notes = "Checking CBS logs, Please wait..."
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh() 
                })                
                Set-Location $Path
                $Connection = (Test-Connection -ComputerName $Computer.computer -Count 1 -Quiet)
                $server = $Computer.computer 
                    $status, $resout = Invoke-Command -ComputerName $server -ScriptBlock {
                          try{
							$CBSlogs = Get-ChildItem -Path C:\Windows\Logs\CBS -Include CBS.log -Recurse | Select-String -Pattern "store corruption"
     						$SXS = Get-ChildItem -Path C:\Windows\Logs\CBS -Include CBS.log -Recurse | Select-String -Pattern "ComponentAnalyzerEvaluateSelfUpdate"
	 
	  if($CBSlogs -ne $null){
           $KB = $null
           $LogSplt = $CBSlogs -split " "
           foreach ($words in $LogSplt){
                if($words -match "KB"){
                    $wordSplt = $words -split "_"
                    foreach($wrd in $wordSplt){
                         if($wrd -match "KB"){
                            $KB += "corruption found", $wrd.Substring(0,9) + "`n"
                         }
                    }
                }

           }
				#return "Completed", $KB
		 If ($SXS -ne $null) {
	  	$SXS1 = $null
		$LogSplt1 = $SXS -split " "
		foreach ($words1 in $LogSplt1){
			if($words1 -match "0x80073712"){
							$SXS1 += "Error SXS Component", $words1.Substring(0,10) + "`n"
						}
                    }
				#return "Completed", $SXS1
		   }
			return "Completed", "$KB $SXS1"
		}
		ElseIf ($SXS -ne $null) {
	  	$SXS1 = $null
		$LogSplt1 = $SXS -split " "
		foreach ($words1 in $LogSplt1){
			if($words1 -match "0x80073712"){
							$SXS1 += "Error SXS Component", $words1.Substring(0,10) + "`n"
						}
                    }
				return "Completed", $SXS1
		   }   
		   
	  Else{
           return "Completed", "No corrupt KB and SXS found"
      }                    
						  }Catch{$Status = "Failed"; $resout = $_.Exception.Message }
                        }
               $uiHash.ListView.Dispatcher.Invoke("Background",[action]{                    
                    $uiHash.Listview.Items.EditItem($Computer)
                    If ($Connection) {
                        $Computer.status = $status
                        $Computer.notes = $resout
                    } ElseIf (-Not $Connection) {
                        $Computer.Status = "Offline"
						$Computer.Notes = "Check Server"
                    } Else {
                        $Computer.Status = "Unknown"
						$Computer.Notes = "Check Server"
                    } 
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh()
                })
                $uiHash.ProgressBar.Dispatcher.Invoke("Normal",[action]{   
                    $uiHash.ProgressBar.value++  
                })
                $uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
                    #Check to see if find job
                    If ($uiHash.ProgressBar.value -eq $uiHash.ProgressBar.Maximum) { 
                        $End = New-Timespan $uihash.StartTime (Get-Date)                                             
                        $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)
                        $uiHash.RunButton.IsEnabled = $True
                        $uiHash.StartImage.Source = "$pwd\Images\Start.jpg"
                        $uiHash.CancelButton.IsEnabled = $False
                        $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.jpg"                    
                    }
                })  
                
            }

            Write-Verbose ("Creating runspace pool and session states")
            $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
            $runspaceHash.runspacepool = [runspacefactory]::CreateRunspacePool(1, $maxConcurrentJobs, $sessionstate, $Host)
            $runspaceHash.runspacepool.Open()   

            ForEach ($Computer in $selectedItems) {
                $uiHash.Listview.Items.EditItem($Computer)
                $computer.Status = "Pending..."
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh() 
                #Create the powershell instance and supply the scriptblock with the other parameters 
                $powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($computer).AddArgument($uiHash).AddArgument($Path)
           
                #Add the runspace into the powershell instance
                $powershell.RunspacePool = $runspaceHash.runspacepool
           
                #Create a temporary collection for each runspace
                $temp = "" | Select-Object PowerShell,Runspace,Computer
                $Temp.Computer = $Computer.computer
                $temp.PowerShell = $powershell
           
                #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
                $temp.Runspace = $powershell.BeginInvoke()
                Write-Verbose ("Adding {0} collection" -f $temp.Computer)
                $jobs.Add($temp) | Out-Null                
            }
            #endregion

        } 
             Else {
                $uiHash.StatusTextBox.Foreground = "Red"
                $uiHash.StatusTextBox.Text = "No server/s selected!"
        }
        
}


#Check CHEF Environment Variables
Function Check-CHEFenvVar {
 Write-Debug ("ComboBox {0}" -f $uiHash.RunOptionComboBox.Text)
        $selectedItems = $uiHash.Listview.SelectedItems
        If ($selectedItems.Count -gt 0) {
        $uiHash.ProgressBar.Maximum = $selectedItems.count
        $uiHash.Listview.ItemsSource | ForEach {$_.Status = $Null; $_.notes = $Null} 
        #region
            $uiHash.RunButton.IsEnabled = $False
            $uiHash.StartImage.Source = "$pwd\Images\Start_locked.jpg"
            $uiHash.CancelButton.IsEnabled = $True
            $uiHash.CancelImage.Source = "$pwd\Images\Stop.jpg"            
            $uiHash.StatusTextBox.Foreground = "Black"
            $uiHash.StatusTextBox.Text = "Checking CHEF environment variables..."            
            $uiHash.StartTime = (Get-Date)
            
            [Float]$uiHash.ProgressBar.Value = 0
            $scriptBlock = {
                Param (
                    $Computer,
                    $uiHash,
                    $Path
                )               
                $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                    $uiHash.Listview.Items.EditItem($Computer)
                    $computer.Status = "Running"
					$computer.Notes = "Checking CHEF environment variables... please wait."
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh() 
                })                
                Set-Location $Path
                $Connection = (Test-Connection -ComputerName $Computer.computer -Count 1 -Quiet)
                $server = $Computer.computer 
                    $status, $resout = Invoke-Command -ComputerName $server -ScriptBlock {
                          try{
							
                                $pathContent = [Environment]::GetEnvironmentVariable('path', 'Machine')
                                $pathContentValues = $pathContent -split ';'
                                $myPath = ("C:\Windows\system32","C:\Windows","C:\Windows\System32\Wbem","C:\Windows\System32\WindowsPowerShell\v1.0\","C:\opscode\chef\bin\")
                                if ($pathContent -ne $null)
                                {
                                    $pathValues = @($pathContentValues)
                                    $myPathValues = @($myPath)
                                    $missing = (Compare-Object $pathValues $myPathValues | Where-Object { $_.SideIndicator -eq '=>' } | Foreach-Object { $_.InputObject }) | Out-String
                                    $missingvalues = $missing -split ' '
                                    if($missing){
                                            Return "Completed","Below System Environment Variable(s) are Missing for CHEF:`n$missingvalues"
                                            }
                                  else
                                  {
                                    Return "Completed","System Environment Variables for CHEF are present."
                                  }

                                  }
						  
						  }Catch{return "Failed", $Error[0]}
                        } 

               $uiHash.ListView.Dispatcher.Invoke("Background",[action]{                    
                    $uiHash.Listview.Items.EditItem($Computer)
                    If ($Connection) {
                        $Computer.status = $status
                        $Computer.notes = $resout
                    } ElseIf (-Not $Connection) {
                        $Computer.Status = "Offline"
						$Computer.Notes = "Check Server"
                    } Else {
                        $Computer.Status = "Unknown"
						$Computer.Notes = "Check Server"
                    } 
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh()
                })
                $uiHash.ProgressBar.Dispatcher.Invoke("Normal",[action]{   
                    $uiHash.ProgressBar.value++  
                })
                $uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
                    #Check to see if find job
                    If ($uiHash.ProgressBar.value -eq $uiHash.ProgressBar.Maximum) { 
                        $End = New-Timespan $uihash.StartTime (Get-Date)                                             
                        $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)
                        $uiHash.RunButton.IsEnabled = $True
                        $uiHash.StartImage.Source = "$pwd\Images\Start.jpg"
                        $uiHash.CancelButton.IsEnabled = $False
                        $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.jpg"                    
                    }
                })  
                
            }

            Write-Verbose ("Creating runspace pool and session states")
            $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
            $runspaceHash.runspacepool = [runspacefactory]::CreateRunspacePool(1, $maxConcurrentJobs, $sessionstate, $Host)
            $runspaceHash.runspacepool.Open()   

            ForEach ($Computer in $selectedItems) {
                $uiHash.Listview.Items.EditItem($Computer)
                $computer.Status = "Pending..."
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh() 
                #Create the powershell instance and supply the scriptblock with the other parameters 
                $powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($computer).AddArgument($uiHash).AddArgument($Path)
           
                #Add the runspace into the powershell instance
                $powershell.RunspacePool = $runspaceHash.runspacepool
           
                #Create a temporary collection for each runspace
                $temp = "" | Select-Object PowerShell,Runspace,Computer
                $Temp.Computer = $Computer.computer
                $temp.PowerShell = $powershell
           
                #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
                $temp.Runspace = $powershell.BeginInvoke()
                Write-Verbose ("Adding {0} collection" -f $temp.Computer)
                $jobs.Add($temp) | Out-Null                
            }
            #endregion

        } 
             Else {
                $uiHash.StatusTextBox.Foreground = "Red"
                $uiHash.StatusTextBox.Text = "No server/s selected!"
        }
        
}



#Check Add Missing CHEF Env Variables
Function AddMissingCHEFenvVar {
 Write-Debug ("ComboBox {0}" -f $uiHash.RunOptionComboBox.Text)
        $selectedItems = $uiHash.Listview.SelectedItems
        If ($selectedItems.Count -gt 0) {
        $uiHash.ProgressBar.Maximum = $selectedItems.count
        $uiHash.Listview.ItemsSource | ForEach {$_.Status = $Null; $_.notes = $Null} 
        #region
            $uiHash.RunButton.IsEnabled = $False
            $uiHash.StartImage.Source = "$pwd\Images\Start_locked.jpg"
            $uiHash.CancelButton.IsEnabled = $True
            $uiHash.CancelImage.Source = "$pwd\Images\Stop.jpg"            
            $uiHash.StatusTextBox.Foreground = "Black"
            $uiHash.StatusTextBox.Text = "Checking Patching Reboot Flag..."            
            $uiHash.StartTime = (Get-Date)
            
            [Float]$uiHash.ProgressBar.Value = 0
            $scriptBlock = {
                Param (
                    $Computer,
                    $uiHash,
                    $Path
                )               
                $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                    $uiHash.Listview.Items.EditItem($Computer)
                    $computer.Status = "Running"
					$computer.Notes = "Checking Patching Reboot Flag... please wait."
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh() 
                })                
                Set-Location $Path
                $Connection = (Test-Connection -ComputerName $Computer.computer -Count 1 -Quiet)
                $server = $Computer.computer 
                    $status, $resout = Invoke-Command -ComputerName $server -ScriptBlock {
                          try{
							
                            $pathContent = [Environment]::GetEnvironmentVariable('path', 'Machine')
                            $pathContentValues = $pathContent -split ';'
                            $myPath = ("C:\Windows\system32","C:\Windows","C:\Windows\System32\Wbem","C:\Windows\System32\WindowsPowerShell\v1.0\","C:\opscode\chef\bin\")
                            if ($pathContent -ne $null)
                            {
                                $pathValues = @($pathContentValues)
                                $myPathValues = @($myPath)
                                $missing = (Compare-Object $pathValues $myPathValues | Where-Object { $_.SideIndicator -eq '=>' } | Foreach-Object { $_.InputObject })
                                $missingvalues = $missing -join ';'
                                if($missing){
                                        $newPath = "$pathContent;$missingvalues" -replace ';+', ';'
                                        [System.Environment]::SetEnvironmentVariable("Path",$newPath,"Machine")
                                        Return "Completed","Below System Environment Variable(s) were added for CHEF:`n$missingvalues"

                                        }


                              else
                              {
                                Return "Completed","System Environment Variables for CHEF are present."
                              }

                              } 

						  }Catch{return "Failed", $Error[0]}
                        } 

               $uiHash.ListView.Dispatcher.Invoke("Background",[action]{                    
                    $uiHash.Listview.Items.EditItem($Computer)
                    If ($Connection) {
                        $Computer.status = $status
                        $Computer.notes = $resout
                    } ElseIf (-Not $Connection) {
                        $Computer.Status = "Offline"
						$Computer.Notes = "Check Server"
                    } Else {
                        $Computer.Status = "Unknown"
						$Computer.Notes = "Check Server"
                    } 
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh()
                })
                $uiHash.ProgressBar.Dispatcher.Invoke("Normal",[action]{   
                    $uiHash.ProgressBar.value++  
                })
                $uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
                    #Check to see if find job
                    If ($uiHash.ProgressBar.value -eq $uiHash.ProgressBar.Maximum) { 
                        $End = New-Timespan $uihash.StartTime (Get-Date)                                             
                        $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)
                        $uiHash.RunButton.IsEnabled = $True
                        $uiHash.StartImage.Source = "$pwd\Images\Start.jpg"
                        $uiHash.CancelButton.IsEnabled = $False
                        $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.jpg"                    
                    }
                })  
                
            }

            Write-Verbose ("Creating runspace pool and session states")
            $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
            $runspaceHash.runspacepool = [runspacefactory]::CreateRunspacePool(1, $maxConcurrentJobs, $sessionstate, $Host)
            $runspaceHash.runspacepool.Open()   

            ForEach ($Computer in $selectedItems) {
                $uiHash.Listview.Items.EditItem($Computer)
                $computer.Status = "Pending..."
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh() 
                #Create the powershell instance and supply the scriptblock with the other parameters 
                $powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($computer).AddArgument($uiHash).AddArgument($Path)
           
                #Add the runspace into the powershell instance
                $powershell.RunspacePool = $runspaceHash.runspacepool
           
                #Create a temporary collection for each runspace
                $temp = "" | Select-Object PowerShell,Runspace,Computer
                $Temp.Computer = $Computer.computer
                $temp.PowerShell = $powershell
           
                #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
                $temp.Runspace = $powershell.BeginInvoke()
                Write-Verbose ("Adding {0} collection" -f $temp.Computer)
                $jobs.Add($temp) | Out-Null                
            }
            #endregion

        } 
             Else {
                $uiHash.StatusTextBox.Foreground = "Red"
                $uiHash.StatusTextBox.Text = "No server/s selected!"
        }
        
}



#Check Patching Reboot Flag
Function Check-PatchingRebootFlag {
 Write-Debug ("ComboBox {0}" -f $uiHash.RunOptionComboBox.Text)
        $selectedItems = $uiHash.Listview.SelectedItems
        If ($selectedItems.Count -gt 0) {
        $uiHash.ProgressBar.Maximum = $selectedItems.count
        $uiHash.Listview.ItemsSource | ForEach {$_.Status = $Null; $_.notes = $Null} 
        #region Check Patching Reboot Flag
            $uiHash.RunButton.IsEnabled = $False
            $uiHash.StartImage.Source = "$pwd\Images\Start_locked.jpg"
            $uiHash.CancelButton.IsEnabled = $True
            $uiHash.CancelImage.Source = "$pwd\Images\Stop.jpg"            
            $uiHash.StatusTextBox.Foreground = "Black"
            $uiHash.StatusTextBox.Text = "Checking Patching Reboot Flag..."            
            $uiHash.StartTime = (Get-Date)
            
            [Float]$uiHash.ProgressBar.Value = 0
            $scriptBlock = {
                Param (
                    $Computer,
                    $uiHash,
                    $Path
                )               
                $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                    $uiHash.Listview.Items.EditItem($Computer)
                    $computer.Status = "Running"
					$computer.Notes = "Checking Patching Reboot Flag... please wait."
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh() 
                })                
                Set-Location $Path
                $Connection = (Test-Connection -ComputerName $Computer.computer -Count 1 -Quiet)
                $server = $Computer.computer 
                    $status, $resout = Invoke-Command -ComputerName $server -ScriptBlock {
                          try{
							$wmi = gwmi Win32_OperatingSystem
							$LBTime = $wmi.ConvertToDateTime($wmi.Lastbootuptime)
                            $localtime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
							$RebootFlag = Test-Path -path "C:\Users\Administrator\reboot_requested.txt"
							If ($RebootFlag)
							{
							$RFModified = (get-item C:\Users\Administrator\reboot_requested.txt | select "LastWriteTime").LastWriteTime
							return "Completed", "Reboot Flag Last Modified: $RFModified `nLast Boot Time: $LBTime `n`nLocal Time: $localtime"
							}
							else
							{
							return "Completed", "No reboot flag found `nLast Boot Time: $LBTime `n`nLocal Time: $localtime"
							}
						  
						  }Catch{return "Failed", $Error[0]}
                        } 

               $uiHash.ListView.Dispatcher.Invoke("Background",[action]{                    
                    $uiHash.Listview.Items.EditItem($Computer)
                    If ($Connection) {
                        $Computer.status = $status
                        $Computer.notes = $resout
                    } ElseIf (-Not $Connection) {
                        $Computer.Status = "Offline"
						$Computer.Notes = "Check Server"
                    } Else {
                        $Computer.Status = "Unknown"
						$Computer.Notes = "Check Server"
                    } 
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh()
                })
                $uiHash.ProgressBar.Dispatcher.Invoke("Normal",[action]{   
                    $uiHash.ProgressBar.value++  
                })
                $uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
                    #Check to see if find job
                    If ($uiHash.ProgressBar.value -eq $uiHash.ProgressBar.Maximum) { 
                        $End = New-Timespan $uihash.StartTime (Get-Date)                                             
                        $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)
                        $uiHash.RunButton.IsEnabled = $True
                        $uiHash.StartImage.Source = "$pwd\Images\Start.jpg"
                        $uiHash.CancelButton.IsEnabled = $False
                        $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.jpg"                    
                    }
                })  
                
            }

            Write-Verbose ("Creating runspace pool and session states")
            $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
            $runspaceHash.runspacepool = [runspacefactory]::CreateRunspacePool(1, $maxConcurrentJobs, $sessionstate, $Host)
            $runspaceHash.runspacepool.Open()   

            ForEach ($Computer in $selectedItems) {
                $uiHash.Listview.Items.EditItem($Computer)
                $computer.Status = "Pending..."
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh() 
                #Create the powershell instance and supply the scriptblock with the other parameters 
                $powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($computer).AddArgument($uiHash).AddArgument($Path)
           
                #Add the runspace into the powershell instance
                $powershell.RunspacePool = $runspaceHash.runspacepool
           
                #Create a temporary collection for each runspace
                $temp = "" | Select-Object PowerShell,Runspace,Computer
                $Temp.Computer = $Computer.computer
                $temp.PowerShell = $powershell
           
                #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
                $temp.Runspace = $powershell.BeginInvoke()
                Write-Verbose ("Adding {0} collection" -f $temp.Computer)
                $jobs.Add($temp) | Out-Null                
            }
            #endregion

        } 
             Else {
                $uiHash.StatusTextBox.Foreground = "Red"
                $uiHash.StatusTextBox.Text = "No server/s selected!"
        }
        
}


#Delete Reboot Flag
Function Delete-RebootFlag{
 Write-Debug ("ComboBox {0}" -f $uiHash.RunOptionComboBox.Text)
        $selectedItems = $uiHash.Listview.SelectedItems
        If ($selectedItems.Count -gt 0) {
        $uiHash.ProgressBar.Maximum = $selectedItems.count
        $uiHash.Listview.ItemsSource | ForEach {$_.Status = $Null; $_.notes = $Null} 
        #region DeleterebootFlag
            $uiHash.RunButton.IsEnabled = $False
            $uiHash.StartImage.Source = "$pwd\Images\Start_locked.jpg"
            $uiHash.CancelButton.IsEnabled = $True
            $uiHash.CancelImage.Source = "$pwd\Images\Stop.jpg"            
            $uiHash.StatusTextBox.Foreground = "Black"
            $uiHash.StatusTextBox.Text = "Deleting reboot flag and running chef-client..."            
            $uiHash.StartTime = (Get-Date)
            
            [Float]$uiHash.ProgressBar.Value = 0
            $scriptBlock = {
                Param (
                    $Computer,
                    $uiHash,
                    $Path
                )               
                $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                    $uiHash.Listview.Items.EditItem($Computer)
                    $computer.Status = "Running"
					$computer.Notes = "Deleting reboot flag and running chef-client..."
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh() 
                })                
                Set-Location $Path
                $Connection = (Test-Connection -ComputerName $Computer.computer -Count 1 -Quiet)
                $server = $Computer.computer 
                    $status, $resout = Invoke-Command -ComputerName $server -ScriptBlock {
                          try{
                          
							$wmi = gwmi Win32_OperatingSystem
							$LBTime = $wmi.ConvertToDateTime($wmi.Lastbootuptime) 
							$RebootFlag = Test-Path -path "C:\Users\Administrator\reboot_requested.txt"
							If ($RebootFlag)
							{
							del c:\Users\Administrator\reboot_requested.txt
                            $runCC = chef-client
							$ccout = $runCC | ? {$_ -match "resolving cookbooks for run list:"} | out-string
							return "Completed", "Reboot_requested.txt deleted `nChef-Client run complete `n$ccout"
							}
							else
							{
							return "Completed", "No reboot flag found `nLast Boot Time: $LBTime"
                            
							}
						  #del c:\Users\Administrator\reboot_requested.txt
							#$runCC = chef-client
							#$ccout = $runCC | ? {$_ -match "resolving cookbooks for run list:"} | out-string
						 
						# return "Completed", "Reboot_requested.txt deleted `nChef-Client run complete `n$ccout"
						  }Catch{return "Failed", $Error[0]}
                        } 

               $uiHash.ListView.Dispatcher.Invoke("Background",[action]{                    
                    $uiHash.Listview.Items.EditItem($Computer)
                    If ($Connection) {
                        $Computer.status = $status
                        $Computer.notes = $resout
                    } ElseIf (-Not $Connection) {
                        $Computer.Status = "Offline"
						$Computer.Notes = "Check Server"
                    } Else {
                        $Computer.Status = "Unknown"
						$Computer.Notes = "Check Server"
                    } 
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh()
                })
                $uiHash.ProgressBar.Dispatcher.Invoke("Normal",[action]{   
                    $uiHash.ProgressBar.value++  
                })
                $uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
                    #Check to see if find job
                    If ($uiHash.ProgressBar.value -eq $uiHash.ProgressBar.Maximum) { 
                        $End = New-Timespan $uihash.StartTime (Get-Date)                                             
                        $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)
                        $uiHash.RunButton.IsEnabled = $True
                        $uiHash.StartImage.Source = "$pwd\Images\Start.jpg"
                        $uiHash.CancelButton.IsEnabled = $False
                        $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.jpg"                    
                    }
                })  
                
            }

            Write-Verbose ("Creating runspace pool and session states")
            $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
            $runspaceHash.runspacepool = [runspacefactory]::CreateRunspacePool(1, $maxConcurrentJobs, $sessionstate, $Host)
            $runspaceHash.runspacepool.Open()   

            ForEach ($Computer in $selectedItems) {
                $uiHash.Listview.Items.EditItem($Computer)
                $computer.Status = "Pending..."
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh() 
                #Create the powershell instance and supply the scriptblock with the other parameters 
                $powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($computer).AddArgument($uiHash).AddArgument($Path)
           
                #Add the runspace into the powershell instance
                $powershell.RunspacePool = $runspaceHash.runspacepool
           
                #Create a temporary collection for each runspace
                $temp = "" | Select-Object PowerShell,Runspace,Computer
                $Temp.Computer = $Computer.computer
                $temp.PowerShell = $powershell
           
                #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
                $temp.Runspace = $powershell.BeginInvoke()
                Write-Verbose ("Adding {0} collection" -f $temp.Computer)
                $jobs.Add($temp) | Out-Null                
            }
            #endregion

        } 
             Else {
                $uiHash.StatusTextBox.Foreground = "Red"
                $uiHash.StatusTextBox.Text = "No server/s selected!"
        }
        
}


#Check Ruby Process Event
Function Check-RubyProcess {
 Write-Debug ("ComboBox {0}" -f $uiHash.RunOptionComboBox.Text)
        $selectedItems = $uiHash.Listview.SelectedItems
        If ($selectedItems.Count -gt 0) {
        $uiHash.ProgressBar.Maximum = $selectedItems.count
        $uiHash.Listview.ItemsSource | ForEach {$_.Status = $Null; $_.notes = $Null} 
        #region Check Ruby Process
            $uiHash.RunButton.IsEnabled = $False
            $uiHash.StartImage.Source = "$pwd\Images\Start_locked.jpg"
            $uiHash.CancelButton.IsEnabled = $True
            $uiHash.CancelImage.Source = "$pwd\Images\Stop.jpg"            
            $uiHash.StatusTextBox.Foreground = "Black"
            $uiHash.StatusTextBox.Text = "Checking Ruby Interpreter Process..."            
            $uiHash.StartTime = (Get-Date)
            
            [Float]$uiHash.ProgressBar.Value = 0
            $scriptBlock = {
                Param (
                    $Computer,
                    $uiHash,
                    $Path
                )               
                $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                    $uiHash.Listview.Items.EditItem($Computer)
                    $computer.Status = "Running"
					$computer.Notes = "Checking Ruby Interpreter Process... please wait."
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh() 
                })                
                Set-Location $Path
                $Connection = (Test-Connection -ComputerName $Computer.computer -Count 1 -Quiet)
                $server = $Computer.computer 
                    $status, $resout = Invoke-Command -ComputerName $server -ScriptBlock {
                          try{
							

                                $process = Get-WmiObject Win32_Process -Filter "Name LIKE 'Ruby.exe'"
                                #$pCount = 0
                                foreach($p in $process){$pCount += 1 }
                                if($pCount -gt 0){return "Completed", "$pCount Ruby Interpreter Process Found"}
                                else {return "Completed", "No Ruby Interpreter Process Found"}


						  }Catch{return "Failed", $Error[0]}
                        } 

               $uiHash.ListView.Dispatcher.Invoke("Background",[action]{                    
                    $uiHash.Listview.Items.EditItem($Computer)
                    If ($Connection) {
                        $Computer.status = $status
                        $Computer.notes = $resout
                    } ElseIf (-Not $Connection) {
                        $Computer.Status = "Offline"
						$Computer.Notes = "Check Server"
                    } Else {
                        $Computer.Status = "Unknown"
						$Computer.Notes = "Check Server"
                    } 
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh()
                })
                $uiHash.ProgressBar.Dispatcher.Invoke("Normal",[action]{   
                    $uiHash.ProgressBar.value++  
                })
                $uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
                    #Check to see if find job
                    If ($uiHash.ProgressBar.value -eq $uiHash.ProgressBar.Maximum) { 
                        $End = New-Timespan $uihash.StartTime (Get-Date)                                             
                        $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)
                        $uiHash.RunButton.IsEnabled = $True
                        $uiHash.StartImage.Source = "$pwd\Images\Start.jpg"
                        $uiHash.CancelButton.IsEnabled = $False
                        $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.jpg"                    
                    }
                })  
                
            }

            Write-Verbose ("Creating runspace pool and session states")
            $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
            $runspaceHash.runspacepool = [runspacefactory]::CreateRunspacePool(1, $maxConcurrentJobs, $sessionstate, $Host)
            $runspaceHash.runspacepool.Open()   

            ForEach ($Computer in $selectedItems) {
                $uiHash.Listview.Items.EditItem($Computer)
                $computer.Status = "Pending..."
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh() 
                #Create the powershell instance and supply the scriptblock with the other parameters 
                $powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($computer).AddArgument($uiHash).AddArgument($Path)
           
                #Add the runspace into the powershell instance
                $powershell.RunspacePool = $runspaceHash.runspacepool
           
                #Create a temporary collection for each runspace
                $temp = "" | Select-Object PowerShell,Runspace,Computer
                $Temp.Computer = $Computer.computer
                $temp.PowerShell = $powershell
           
                #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
                $temp.Runspace = $powershell.BeginInvoke()
                Write-Verbose ("Adding {0} collection" -f $temp.Computer)
                $jobs.Add($temp) | Out-Null                
            }
            #endregion

        } 
             Else {
                $uiHash.StatusTextBox.Foreground = "Red"
                $uiHash.StatusTextBox.Text = "No server/s selected!"
        }
        
}


#End Ruby Process Event
Function End-RubyProcess {
 Write-Debug ("ComboBox {0}" -f $uiHash.RunOptionComboBox.Text)
        $selectedItems = $uiHash.Listview.SelectedItems
        If ($selectedItems.Count -gt 0) {
        $uiHash.ProgressBar.Maximum = $selectedItems.count
        $uiHash.Listview.ItemsSource | ForEach {$_.Status = $Null; $_.notes = $Null} 
        #region END Ruby Process
            $uiHash.RunButton.IsEnabled = $False
            $uiHash.StartImage.Source = "$pwd\Images\Start_locked.jpg"
            $uiHash.CancelButton.IsEnabled = $True
            $uiHash.CancelImage.Source = "$pwd\Images\Stop.jpg"            
            $uiHash.StatusTextBox.Foreground = "Black"
            $uiHash.StatusTextBox.Text = "Terminating Ruby Interpreter Process..."            
            $uiHash.StartTime = (Get-Date)
            
            [Float]$uiHash.ProgressBar.Value = 0
            $scriptBlock = {
                Param (
                    $Computer,
                    $uiHash,
                    $Path
                )               
                $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                    $uiHash.Listview.Items.EditItem($Computer)
                    $computer.Status = "Running"
					$computer.Notes = "Terminating Ruby Interpreter Process... please wait."
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh() 
                })                
                Set-Location $Path
                $Connection = (Test-Connection -ComputerName $Computer.computer -Count 1 -Quiet)
                $server = $Computer.computer 
                    $status, $resout = Invoke-Command -ComputerName $server -ScriptBlock {
                          try{
							
                                $process = (Get-WmiObject -Class Win32_Process -Filter "Name LIKE 'Ruby.exe'").ProcessId
                                foreach($p in $process){$pCount += 1}
                                if($pCount -gt 0){
                                    foreach($p in $process){
                                        Stop-Process $process -Force
    	                                return "Completed", "$pCount Ruby Interpreter Process Terminated"}
                                    }
	                             else {return "Completed", "No Ruby Interpreter Process Found"}


						  }Catch{return "Failed", $Error[0]}
                        } 

               $uiHash.ListView.Dispatcher.Invoke("Background",[action]{                    
                    $uiHash.Listview.Items.EditItem($Computer)
                    If ($Connection) {
                        $Computer.status = $status
                        $Computer.notes = $resout
                    } ElseIf (-Not $Connection) {
                        $Computer.Status = "Offline"
						$Computer.Notes = "Check Server"
                    } Else {
                        $Computer.Status = "Unknown"
						$Computer.Notes = "Check Server"
                    } 
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh()
                })
                $uiHash.ProgressBar.Dispatcher.Invoke("Normal",[action]{   
                    $uiHash.ProgressBar.value++  
                })
                $uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
                    #Check to see if find job
                    If ($uiHash.ProgressBar.value -eq $uiHash.ProgressBar.Maximum) { 
                        $End = New-Timespan $uihash.StartTime (Get-Date)                                             
                        $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)
                        $uiHash.RunButton.IsEnabled = $True
                        $uiHash.StartImage.Source = "$pwd\Images\Start.jpg"
                        $uiHash.CancelButton.IsEnabled = $False
                        $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.jpg"                    
                    }
                })  
                
            }

            Write-Verbose ("Creating runspace pool and session states")
            $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
            $runspaceHash.runspacepool = [runspacefactory]::CreateRunspacePool(1, $maxConcurrentJobs, $sessionstate, $Host)
            $runspaceHash.runspacepool.Open()   

            ForEach ($Computer in $selectedItems) {
                $uiHash.Listview.Items.EditItem($Computer)
                $computer.Status = "Pending..."
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh() 
                #Create the powershell instance and supply the scriptblock with the other parameters 
                $powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($computer).AddArgument($uiHash).AddArgument($Path)
           
                #Add the runspace into the powershell instance
                $powershell.RunspacePool = $runspaceHash.runspacepool
           
                #Create a temporary collection for each runspace
                $temp = "" | Select-Object PowerShell,Runspace,Computer
                $Temp.Computer = $Computer.computer
                $temp.PowerShell = $powershell
           
                #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
                $temp.Runspace = $powershell.BeginInvoke()
                Write-Verbose ("Adding {0} collection" -f $temp.Computer)
                $jobs.Add($temp) | Out-Null                
            }
            #endregion

        } 
             Else {
                $uiHash.StatusTextBox.Foreground = "Red"
                $uiHash.StatusTextBox.Text = "No server/s selected!"
        }
        
}



#Message prompt for WinDefend Run Actions, WinUpdateFix, Install-IIS, Uninstall-IIS
Function Show-ExecWarning ([string]$execstring) {
    $sc = $uiHash.Listview.selecteditems.count
    $title = "Running $chefstring Warning"
    $message = "You are about to run $execstring on $sc server(s)! `nAre you sure you want to proceed?"
    $button = [System.Windows.Forms.MessageBoxButtons]::YesNo
    $icon = [Windows.Forms.MessageBoxIcon]::Warning
    $msgbox = [windows.forms.messagebox]::Show($message,$title,$button,$icon)
    if($execstring -eq "WinDefendOffboard"){
        if ($msgbox -eq "Yes") {Offboard-WinDefend}
        }
		elseif($execstring -eq "WinDefendUninstall"){
        if ($msgbox -eq "Yes") {WinDefendUninstall}
		}
        elseif($execstring -eq "Install-Win2012ESUYear1ProductKey"){
        if ($msgbox -eq "Yes") {Install-Win2012ESUYear1ProductKey}
		}
		elseif($execstring -eq "RepairSEP"){
        if ($msgbox -eq "Yes") {RepairSEP}
		}
		elseif($execstring -eq "EndSepScanService"){
        if ($msgbox -eq "Yes") {EndSepScanService}
		}
		elseif($execstring -eq "SEPOLDVirusDefDELETE"){
        if ($msgbox -eq "Yes") {SEPOLDVirusDefDELETE}
		}
		elseif($execstring -eq "TMEDRProdAgentDeployment"){
        if ($msgbox -eq "Yes") {TMEDRProdAgentDeployment}
		}
        elseif($execstring -eq "Install-IIS"){
        if ($msgbox -eq "Yes") {Install-IIS}
		}
		elseif($execstring -eq "Uninstall-IIS"){
        if ($msgbox -eq "Yes") {Uninstall-IIS}
		}
        elseif($execstring -eq "Delete-RebootFlag"){
        if ($msgbox -eq "Yes") {Delete-RebootFlag}
		}
        elseif($execstring -eq "END-RubyProcess"){
        if ($msgbox -eq "Yes") {END-RubyProcess}
		}
        elseif($execstring -eq "AddMissingCHEFEnvVar"){
        if ($msgbox -eq "Yes") {AddMissingCHEFEnvVar}
		}
		elseif($execstring -eq "WinUpdateFix-MultipleServers"){
        if ($msgbox -eq "Yes") {WinUpdateFix-MultipleServers}
		}
		else{
        if ($msgbox -eq "Yes") {WinUpdateFix}
    	}
}

#WindowsUpdateFix on Multiple Servers
Function WinUpdateFix-MultipleServers {
 Write-Debug ("ComboBox {0}" -f $uiHash.RunOptionComboBox.Text)
    $selectedItems = $uiHash.Listview.SelectedItems
    If ($selectedItems.Count -gt 0) {
        $uiHash.ProgressBar.Maximum = $selectedItems.count
        $uiHash.Listview.ItemsSource | ForEach {$_.Status = $Null; $_.notes = $Null} 
	#Region WinUpdateFix-MultipleServers
            $uiHash.RunButton.IsEnabled = $False
            $uiHash.StartImage.Source = "$pwd\Images\Start_locked.jpg"
            $uiHash.CancelButton.IsEnabled = $True
            $uiHash.CancelImage.Source = "$pwd\Images\Stop.jpg"            
            $uiHash.StatusTextBox.Foreground = "Black"
            $uiHash.StatusTextBox.Text = "Running WinUpdateFix on Multiple Servers..."            
            $uiHash.StartTime = (Get-Date)
            [Float]$uiHash.ProgressBar.Value = 0
            #StartWinUpdateFix
			$scriptBlock = {
                Param (
                    $Computer,
                    $uiHash,
                    $Path
                    
                )               
                          
                Set-Location $Path
                $Connection = (Test-Connection -ComputerName $Computer.computer -Count 1 -Quiet)
                $server = $Computer.computer
                if($Connection){

                $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                    $uiHash.Listview.Items.EditItem($Computer)
                    $computer.Status = "Running"
                    $computer.Notes = "Copying files..."
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh()
                }) 
                    try{
							$SPath = "\\unpiox03pn\IntelSystems\Lifecycle\Master_Scripts\Windows_Update_Fix"
                        	$DPath = "\\" + $server + "\c$\temp\Windows_Update_Fix"
                        #if($LASTEXITCODE -le 1 -or $LASTEXITCODE -eq 3 -or $LASTEXITCODE -eq 5 -or $LASTEXITCODE -eq 7){
                           $chkWUFIXpath = Test-Path -path "\\$server\c$\temp\Windows_Update_Fix\*"
                           if($chkWUFIXpath){
                            $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                            $uiHash.Listview.Items.EditItem($Computer)
                            $computer.Status = "Completed"
                            $computer.Notes = "Folder Exists"
                            $uiHash.Listview.Items.CommitEdit()
                            $uiHash.Listview.Items.Refresh()})
                            $RunWUFIX = $True
														
                           }#here
						   else{
								robocopy "$SPath" "$DPath" /S /Z /R:1 /NFL /NDL /NJH /NJS /nc /ns /np
                            $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                            $uiHash.Listview.Items.EditItem($Computer)
                            $computer.Status = "Completed"
                            $computer.Notes = "Robocopy Complete"
                            $uiHash.Listview.Items.CommitEdit()
                            $uiHash.Listview.Items.Refresh()})
                            $RunWUFIX = $True
							 #}
						   #else{$getStatus = "Failed"; $getNotes = "Robocopy Failed!"; $RunWUFIX = $False}
						   #}
                        	}
						   #}Else{$getStatus = "Failed"; $getNotes = "LastExitCode = $LASTEXITCODE" + "$($error[0])"; $RunWUFIX = $False}
                    }Catch{$getStatus = "Failed"; $getNotes = $($error[0])}
                    
                    If($RunWUFIX){
                        
                            $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                                $uiHash.Listview.Items.EditItem($Computer)
                                $computer.Status = "Running"
                                $computer.Notes = "Windows Update Fix in progress on multiple servers... please wait."
                                $uiHash.Listview.Items.CommitEdit()
                                $uiHash.Listview.Items.Refresh()
                            })
                            try{
                                #$session = New-PSSession -ComputerName $server
                                #enter-pssession -session $session
                                
                                #$runEXE = Invoke-Command -cn $server -ScriptBlock{
								Invoke-Command -cn $server -ScriptBlock{
                                    C:\temp\Windows_Update_Fix\Windows_Update_fix.bat | Out-File -FilePath C:\temp\Windows_Update_Fix\WinUpdateFix.Log
									#$OUT = Get-content -path  "C:\temp\Windows_Update_Fix\WinUpdateFix.Log"
                                    #return $OUT
									
                                } -ErrorVariable $EXEerr -ErrorAction SilentlyContinue
                               
                                    $getStatus = "Completed"; $getNotes = "Log is located in \\$server\C$\temp\Windows_Update_Fix\WinUpdateFix.Log"
                                
                            }Catch{$getStatus = "Failed"; $getNotes = $Error[0]}
                        #}
   
                    }
                }
                $uiHash.ListView.Dispatcher.Invoke("Background",[action]{                    
                    $uiHash.Listview.Items.EditItem($Computer)
                    $Computer.Notes = $null
                    If ($Connection) {
                        $Computer.Status = "Completed"
                        $Computer.Notes = $getNotes
                    } ElseIf (-Not $Connection) {
                        $Computer.Status = "Offline"
						$Computer.Notes = "Check Server"
                    } Else {
                        $Computer.Status = "Unknown"
						$Computer.Notes = "Check Server"
                    } 
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh()
                })
                $uiHash.ProgressBar.Dispatcher.Invoke("Normal",[action]{   
                    $uiHash.ProgressBar.value++  
                })
                $uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
                    #Check to see if find job
                    If ($uiHash.ProgressBar.value -eq $uiHash.ProgressBar.Maximum) { 
                        $End = New-Timespan $uihash.StartTime (Get-Date)                                             
                        $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)
                        $uiHash.RunButton.IsEnabled = $True
                        $uiHash.StartImage.Source = "$pwd\Images\Start.jpg"
                        $uiHash.CancelButton.IsEnabled = $False
                        $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.jpg"                    
                    }
                })  
            }
		#EndWinUpdateFix-MultipleServers
            Write-Verbose ("Creating runspace pool and session states")
            $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
            $runspaceHash.runspacepool = [runspacefactory]::CreateRunspacePool(1, $maxConcurrentJobs, $sessionstate, $Host)
            $runspaceHash.runspacepool.Open()   

            ForEach ($Computer in $selectedItems) {
                $uiHash.Listview.Items.EditItem($Computer)
                $computer.Status = "Pending..."
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh() 
                #Create the powershell instance and supply the scriptblock with the other parameters 
                $powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($computer).AddArgument($uiHash).AddArgument($Path)
           
                #Add the runspace into the powershell instance
                $powershell.RunspacePool = $runspaceHash.runspacepool
           
                #Create a temporary collection for each runspace
                $temp = "" | Select-Object PowerShell,Runspace,Computer
                $Temp.Computer = $Computer.computer
                $temp.PowerShell = $powershell
           
                #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
                $temp.Runspace = $powershell.BeginInvoke()
                Write-Verbose ("Adding {0} collection" -f $temp.Computer)
                $jobs.Add($temp) | Out-Null                
            }
    #Endregion WinUpdateFix-MultipleServers
            }Else {
                $uiHash.StatusTextBox.Foreground = "Red"
                $uiHash.StatusTextBox.Text = "No server/s selected!"
            }
}

#WindowsUpdateFix Single Server
Function WinUpdateFix {
 Write-Debug ("ComboBox {0}" -f $uiHash.RunOptionComboBox.Text)
    $selectedItems = $uiHash.Listview.SelectedItems
    If ($selectedItems.Count -gt 0) {
        $uiHash.ProgressBar.Maximum = $selectedItems.count
        $uiHash.Listview.ItemsSource | ForEach {$_.Status = $Null; $_.notes = $Null} 
	#Region WinUpdateFix
            $uiHash.RunButton.IsEnabled = $False
            $uiHash.StartImage.Source = "$pwd\Images\Start_locked.jpg"
            $uiHash.CancelButton.IsEnabled = $True
            $uiHash.CancelImage.Source = "$pwd\Images\Stop.jpg"            
            $uiHash.StatusTextBox.Foreground = "Black"
            $uiHash.StatusTextBox.Text = "Running WinUpdateFix..."            
            $uiHash.StartTime = (Get-Date)
            [Float]$uiHash.ProgressBar.Value = 0
            #StartWinUpdateFix
			$scriptBlock = {
                Param (
                    $Computer,
                    $uiHash,
                    $Path
                    
                )               
                          
                Set-Location $Path
                $Connection = (Test-Connection -ComputerName $Computer.computer -Count 1 -Quiet)
                $server = $Computer.computer
                if($Connection){

                $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                    $uiHash.Listview.Items.EditItem($Computer)
                    $computer.Status = "Running"
                    $computer.Notes = "Copying files..."
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh()
                }) 
                    try{
							$SPath = "\\unpiox03pn\IntelSystems\Lifecycle\Master_Scripts\Windows_Update_Fix"
                        	$DPath = "\\" + $server + "\c$\temp\Windows_Update_Fix"
                        #if($LASTEXITCODE -le 1 -or $LASTEXITCODE -eq 3 -or $LASTEXITCODE -eq 5 -or $LASTEXITCODE -eq 7){
                           $chkWUFIXpath = Test-Path -path "\\$server\c$\temp\Windows_Update_Fix\*"
                           if($chkWUFIXpath){
                            $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                            $uiHash.Listview.Items.EditItem($Computer)
                            $computer.Status = "Completed"
                            $computer.Notes = "Folder Exists"
                            $uiHash.Listview.Items.CommitEdit()
                            $uiHash.Listview.Items.Refresh()})
                            $RunWUFIX = $True
														
                           }#here
						   else{
								robocopy "$SPath" "$DPath" /S /Z /R:1 /NFL /NDL /NJH /NJS /nc /ns /np
                            $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                            $uiHash.Listview.Items.EditItem($Computer)
                            $computer.Status = "Completed"
                            $computer.Notes = "Robocopy Complete"
                            $uiHash.Listview.Items.CommitEdit()
                            $uiHash.Listview.Items.Refresh()})
                            $RunWUFIX = $True
							 #}
						   #else{$getStatus = "Failed"; $getNotes = "Robocopy Failed!"; $RunWUFIX = $False}
						   #}
                        	}
						   #}Else{$getStatus = "Failed"; $getNotes = "LastExitCode = $LASTEXITCODE" + "$($error[0])"; $RunWUFIX = $False}
                    }Catch{$getStatus = "Failed"; $getNotes = $($error[0])}
                    
                    If($RunWUFIX){
                        
                            $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                                $uiHash.Listview.Items.EditItem($Computer)
                                $computer.Status = "Running"
                                $computer.Notes = "Windows Update Fix in progress... please wait."
                                $uiHash.Listview.Items.CommitEdit()
                                $uiHash.Listview.Items.Refresh()
                            })
                            try{
                                #$session = New-PSSession -ComputerName $server
                                #enter-pssession -session $session
                                
                                $runEXE = Invoke-Command -cn $server -ScriptBlock{
                                    C:\temp\Windows_Update_Fix\Windows_Update_fix.bat | Out-File -FilePath C:\temp\Windows_Update_Fix\WinUpdateFix.Log
									$OUT = Get-content -path  "C:\temp\Windows_Update_Fix\WinUpdateFix.Log"
                                    return $OUT
									
                                } -ErrorVariable $EXEerr -ErrorAction SilentlyContinue
                               if($runEXE){
									$runEXE | Out-GridView -Title 'WinUpdateFix Log'
									#$getStatus = "Completed"; $getNotes = $runEXE
                                }else{
                                    $getStatus = "Failed"; $getNotes = "Check server"
                                }
                            }Catch{$getStatus = "Failed"; $getNotes = $Error[0]}
                        #}
   
                    }
                }
                $uiHash.ListView.Dispatcher.Invoke("Background",[action]{                    
                    $uiHash.Listview.Items.EditItem($Computer)
                    $Computer.Notes = $null
                    If ($Connection) {
                        $Computer.Status = "Completed"
                        #$Computer.Notes = $getNotes
                    } ElseIf (-Not $Connection) {
                        $Computer.Status = "Offline"
						$Computer.Notes = "Check Server"
                    } Else {
                        $Computer.Status = "Unknown"
						$Computer.Notes = "Check Server"
                    } 
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh()
                })
                $uiHash.ProgressBar.Dispatcher.Invoke("Normal",[action]{   
                    $uiHash.ProgressBar.value++  
                })
                $uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
                    #Check to see if find job
                    If ($uiHash.ProgressBar.value -eq $uiHash.ProgressBar.Maximum) { 
                        $End = New-Timespan $uihash.StartTime (Get-Date)                                             
                        $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)
                        $uiHash.RunButton.IsEnabled = $True
                        $uiHash.StartImage.Source = "$pwd\Images\Start.jpg"
                        $uiHash.CancelButton.IsEnabled = $False
                        $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.jpg"                    
                    }
                })  
            }
		#EndWinUpdateFix
            Write-Verbose ("Creating runspace pool and session states")
            $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
            $runspaceHash.runspacepool = [runspacefactory]::CreateRunspacePool(1, $maxConcurrentJobs, $sessionstate, $Host)
            $runspaceHash.runspacepool.Open()   

            ForEach ($Computer in $selectedItems) {
                $uiHash.Listview.Items.EditItem($Computer)
                $computer.Status = "Pending..."
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh() 
                #Create the powershell instance and supply the scriptblock with the other parameters 
                $powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($computer).AddArgument($uiHash).AddArgument($Path)
           
                #Add the runspace into the powershell instance
                $powershell.RunspacePool = $runspaceHash.runspacepool
           
                #Create a temporary collection for each runspace
                $temp = "" | Select-Object PowerShell,Runspace,Computer
                $Temp.Computer = $Computer.computer
                $temp.PowerShell = $powershell
           
                #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
                $temp.Runspace = $powershell.BeginInvoke()
                Write-Verbose ("Adding {0} collection" -f $temp.Computer)
                $jobs.Add($temp) | Out-Null                
            }
    #Endregion WinUpdateFix
            }Else {
                $uiHash.StatusTextBox.Foreground = "Red"
                $uiHash.StatusTextBox.Text = "No server/s selected!"
            }
}

#Check EDGE
Function Check-EDGE{
 Write-Debug ("ComboBox {0}" -f $uiHash.RunOptionComboBox.Text)
        $selectedItems = $uiHash.Listview.SelectedItems
        If ($selectedItems.Count -gt 0) {
        $uiHash.ProgressBar.Maximum = $selectedItems.count
        $uiHash.Listview.ItemsSource | ForEach {$_.Status = $Null; $_.notes = $Null} 
	    #Region CheckEDGE
            $uiHash.RunButton.IsEnabled = $False
            $uiHash.StartImage.Source = "$pwd\Images\Start_locked.jpg"
            $uiHash.CancelButton.IsEnabled = $True
            $uiHash.CancelImage.Source = "$pwd\Images\Stop.jpg"            
            $uiHash.StatusTextBox.Foreground = "Black"
            $uiHash.StatusTextBox.Text = "Checking Microsoft Edge..."            
            $uiHash.StartTime = (Get-Date)
            
            [Float]$uiHash.ProgressBar.Value = 0
            $scriptBlock = {
                Param (
                    $Computer,
                    $uiHash,
                    $Path
                )               
                $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                    $uiHash.Listview.Items.EditItem($Computer)
                    $computer.Status = "Running"
					$computer.Notes = "Checking MS Edge... please wait."
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh() 
                })                
                Set-Location $Path
                $Connection = (Test-Connection -ComputerName $Computer.computer -Count 1 -Quiet)
                $server = $Computer.computer 
                    
				$status, $resout = Invoke-Command -ComputerName $server -ScriptBlock{
					Try{

                        $EdgeCHK = (Get-ItemProperty 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\{56EB18F8-B008-4CBD-B6D2-8C97FE7E9062}').pv
                        $EdgeUpdateReg = (Get-ItemProperty 'HKLM:\SOFTWARE\Policies\Microsoft\EdgeUpdate\' | select "Update{56EB18F8-B008-4CBD-B6D2-8C97FE7E9062}")."Update{56EB18F8-B008-4CBD-B6D2-8C97FE7E9062}"
                        $EdgeUpdateChkReg = (Get-ItemProperty 'HKLM:\SOFTWARE\Policies\Microsoft\EdgeUpdate\' | select "AutoUpdateCheckPeriodMinutes")."AutoUpdateCheckPeriodMinutes"
		
                        if($EdgeCHK)
	                        {
		                        if($EdgeUpdateReg -eq 0 -and $EdgeUpdateChkReg -eq 0)
				                        {
                                        #$EdgeStatus = "Edge $EdgeCHK is installed. Registry is Set."
                                        return "Completed", "Microsoft Edge $EdgeCHK is installed."
                                        }
				                        else{
				                        #$EdgeStatus = "Edge $EdgeCHK is installed. Need to set Registry."
                                        return "Completed", "Microsoft Edge $EdgeCHK is installed. Need to set registry for Edge."
				                        }
	                        }
                        else{
	                        #$EdgeStatus = "Edge is NOT installed."
                            return "Completed", "Microsoft Edge is NOT installed."

	                        }
                       
						}Catch{return "Failed", $Error[0]}
						
					}#EndScriptBlock

               $uiHash.ListView.Dispatcher.Invoke("Background",[action]{                    
                    $uiHash.Listview.Items.EditItem($Computer)
                    If ($Connection) {
                        $Computer.status = $status
                        $Computer.notes = $resout
                    } ElseIf (-Not $Connection) {
                        $Computer.Status = "Offline"
						$Computer.Notes = "Check Server"
                    } Else {
                        $Computer.Status = "Unknown"
						$Computer.Notes = "Check Server"
                    } 
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh()
                })
                $uiHash.ProgressBar.Dispatcher.Invoke("Normal",[action]{   
                    $uiHash.ProgressBar.value++  
                })
                $uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
                    #Check to see if find job
                    If ($uiHash.ProgressBar.value -eq $uiHash.ProgressBar.Maximum) { 
                        $End = New-Timespan $uihash.StartTime (Get-Date)                                             
                        $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)
                        $uiHash.RunButton.IsEnabled = $True
                        $uiHash.StartImage.Source = "$pwd\Images\Start.jpg"
                        $uiHash.CancelButton.IsEnabled = $False
                        $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.jpg"                    
                    }
                })  
                
            }

            Write-Verbose ("Creating runspace pool and session states")
            $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
            $runspaceHash.runspacepool = [runspacefactory]::CreateRunspacePool(1, $maxConcurrentJobs, $sessionstate, $Host)
            $runspaceHash.runspacepool.Open()   

            ForEach ($Computer in $selectedItems) {
                $uiHash.Listview.Items.EditItem($Computer)
                $computer.Status = "Pending..."
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh() 
                #Create the powershell instance and supply the scriptblock with the other parameters 
                $powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($computer).AddArgument($uiHash).AddArgument($Path)
           
                #Add the runspace into the powershell instance
                $powershell.RunspacePool = $runspaceHash.runspacepool
           
                #Create a temporary collection for each runspace
                $temp = "" | Select-Object PowerShell,Runspace,Computer
                $Temp.Computer = $Computer.computer
                $temp.PowerShell = $powershell
           
                #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
                $temp.Runspace = $powershell.BeginInvoke()
                Write-Verbose ("Adding {0} collection" -f $temp.Computer)
                $jobs.Add($temp) | Out-Null                
            }
		#Endregion CheckEDGE
		#EndCheckEDGE
        } 
             Else {
                $uiHash.StatusTextBox.Foreground = "Red"
                $uiHash.StatusTextBox.Text = "No server/s selected!"
        }
        
}

#Check IIS
Function Check-IIS{
 Write-Debug ("ComboBox {0}" -f $uiHash.RunOptionComboBox.Text)
        $selectedItems = $uiHash.Listview.SelectedItems
        If ($selectedItems.Count -gt 0) {
        $uiHash.ProgressBar.Maximum = $selectedItems.count
        $uiHash.Listview.ItemsSource | ForEach {$_.Status = $Null; $_.notes = $Null} 
	#Region CheckIIS
            $uiHash.RunButton.IsEnabled = $False
            $uiHash.StartImage.Source = "$pwd\Images\Start_locked.jpg"
            $uiHash.CancelButton.IsEnabled = $True
            $uiHash.CancelImage.Source = "$pwd\Images\Stop.jpg"            
            $uiHash.StatusTextBox.Foreground = "Black"
            $uiHash.StatusTextBox.Text = "Checking IIS..."            
            $uiHash.StartTime = (Get-Date)
            
            [Float]$uiHash.ProgressBar.Value = 0
            $scriptBlock = {
                Param (
                    $Computer,
                    $uiHash,
                    $Path
                )               
                $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                    $uiHash.Listview.Items.EditItem($Computer)
                    $computer.Status = "Running"
					$computer.Notes = "Checking IIS... please wait."
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh() 
                })                
                Set-Location $Path
                $Connection = (Test-Connection -ComputerName $Computer.computer -Count 1 -Quiet)
                $server = $Computer.computer 
                    
				$status, $resout = Invoke-Command -ComputerName $server -ScriptBlock{
					Try{
					$IISFeat = (Get-WindowsFeature -Name Web-Server | Select-Object Installed).Installed
					if($IISFeat)
						{
						#get-itemproperty HKLM:\SOFTWARE\Microsoft\InetStp\  | select setupstring,versionstring
                        $IISVer = [System.Diagnostics.FileVersionInfo]::GetVersionInfo("$env:SystemRoot\system32\inetsrv\InetMgr.exe").ProductVersion
                        return "Completed", "IIS is Installed: '$IISFeat' `nIIS Ver: $IISVer"
						
                        }
												
					Else{return "Completed", "IIS is Installed: '$IISFeat'"}
				
						}Catch{return "Failed", $Error[0]}
						
					}#EndScriptBlock

               $uiHash.ListView.Dispatcher.Invoke("Background",[action]{                    
                    $uiHash.Listview.Items.EditItem($Computer)
                    If ($Connection) {
                        $Computer.status = $status
                        $Computer.notes = $resout
                    } ElseIf (-Not $Connection) {
                        $Computer.Status = "Offline"
						$Computer.Notes = "Check Server"
                    } Else {
                        $Computer.Status = "Unknown"
						$Computer.Notes = "Check Server"
                    } 
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh()
                })
                $uiHash.ProgressBar.Dispatcher.Invoke("Normal",[action]{   
                    $uiHash.ProgressBar.value++  
                })
                $uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
                    #Check to see if find job
                    If ($uiHash.ProgressBar.value -eq $uiHash.ProgressBar.Maximum) { 
                        $End = New-Timespan $uihash.StartTime (Get-Date)                                             
                        $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)
                        $uiHash.RunButton.IsEnabled = $True
                        $uiHash.StartImage.Source = "$pwd\Images\Start.jpg"
                        $uiHash.CancelButton.IsEnabled = $False
                        $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.jpg"                    
                    }
                })  
                
            }

            Write-Verbose ("Creating runspace pool and session states")
            $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
            $runspaceHash.runspacepool = [runspacefactory]::CreateRunspacePool(1, $maxConcurrentJobs, $sessionstate, $Host)
            $runspaceHash.runspacepool.Open()   

            ForEach ($Computer in $selectedItems) {
                $uiHash.Listview.Items.EditItem($Computer)
                $computer.Status = "Pending..."
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh() 
                #Create the powershell instance and supply the scriptblock with the other parameters 
                $powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($computer).AddArgument($uiHash).AddArgument($Path)
           
                #Add the runspace into the powershell instance
                $powershell.RunspacePool = $runspaceHash.runspacepool
           
                #Create a temporary collection for each runspace
                $temp = "" | Select-Object PowerShell,Runspace,Computer
                $Temp.Computer = $Computer.computer
                $temp.PowerShell = $powershell
           
                #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
                $temp.Runspace = $powershell.BeginInvoke()
                Write-Verbose ("Adding {0} collection" -f $temp.Computer)
                $jobs.Add($temp) | Out-Null                
            }
		#Endregion CheckIIS
		#EndCheckIIS
        } 
             Else {
                $uiHash.StatusTextBox.Foreground = "Red"
                $uiHash.StatusTextBox.Text = "No server/s selected!"
        }
        
}

#Install IIS
Function Install-IIS{
 Write-Debug ("ComboBox {0}" -f $uiHash.RunOptionComboBox.Text)
        $selectedItems = $uiHash.Listview.SelectedItems
        If ($selectedItems.Count -gt 0) {
        $uiHash.ProgressBar.Maximum = $selectedItems.count
        $uiHash.Listview.ItemsSource | ForEach {$_.Status = $Null; $_.notes = $Null} 
	#Region InstallIIS
            $uiHash.RunButton.IsEnabled = $False
            $uiHash.StartImage.Source = "$pwd\Images\Start_locked.jpg"
            $uiHash.CancelButton.IsEnabled = $True
            $uiHash.CancelImage.Source = "$pwd\Images\Stop.jpg"            
            $uiHash.StatusTextBox.Foreground = "Black"
            $uiHash.StatusTextBox.Text = "Installing IIS..."            
            $uiHash.StartTime = (Get-Date)
            
            [Float]$uiHash.ProgressBar.Value = 0
            $scriptBlock = {
                Param (
                    $Computer,
                    $uiHash,
                    $Path
                )               
                $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                    $uiHash.Listview.Items.EditItem($Computer)
                    $computer.Status = "Running"
					$computer.Notes = "Installing IIS... please wait."
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh() 
                })                
                Set-Location $Path
                $Connection = (Test-Connection -ComputerName $Computer.computer -Count 1 -Quiet)
                $server = $Computer.computer 
                    
				$status, $resout = Invoke-Command -ComputerName $server -ScriptBlock{
					Try{
					$IISFeat = (Get-WindowsFeature -Name Web-Server | Select-Object Installed).Installed
					if($IISFeat)
						{
						return "Completed", "IIS is Installed. No action is required"
						}
												
					Else{					
						$IISinstall = Install-WindowsFeature -Name Web-Server -IncludeManagementTools | Select-Object Success, RestartNeeded, ExitCode, FeatureResult | FL
						$IISRes = $IISinstall | Out-String
						return "Completed", "IIS Installation Result:$IISRes"}
				
						}Catch{return "Failed", $Error[0]}
					
					}#EndScriptBlock

               $uiHash.ListView.Dispatcher.Invoke("Background",[action]{                    
                    $uiHash.Listview.Items.EditItem($Computer)
                    If ($Connection) {
                        $Computer.status = $status
                        $Computer.notes = $resout
                    } ElseIf (-Not $Connection) {
                        $Computer.Status = "Offline"
						$Computer.Notes = "Check Server"
                    } Else {
                        $Computer.Status = "Unknown"
						$Computer.Notes = "Check Server"
                    } 
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh()
                })
                $uiHash.ProgressBar.Dispatcher.Invoke("Normal",[action]{   
                    $uiHash.ProgressBar.value++  
                })
                $uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
                    #Check to see if find job
                    If ($uiHash.ProgressBar.value -eq $uiHash.ProgressBar.Maximum) { 
                        $End = New-Timespan $uihash.StartTime (Get-Date)                                             
                        $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)
                        $uiHash.RunButton.IsEnabled = $True
                        $uiHash.StartImage.Source = "$pwd\Images\Start.jpg"
                        $uiHash.CancelButton.IsEnabled = $False
                        $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.jpg"                    
                    }
                })  
                
            }

            Write-Verbose ("Creating runspace pool and session states")
            $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
            $runspaceHash.runspacepool = [runspacefactory]::CreateRunspacePool(1, $maxConcurrentJobs, $sessionstate, $Host)
            $runspaceHash.runspacepool.Open()   

            ForEach ($Computer in $selectedItems) {
                $uiHash.Listview.Items.EditItem($Computer)
                $computer.Status = "Pending..."
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh() 
                #Create the powershell instance and supply the scriptblock with the other parameters 
                $powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($computer).AddArgument($uiHash).AddArgument($Path)
           
                #Add the runspace into the powershell instance
                $powershell.RunspacePool = $runspaceHash.runspacepool
           
                #Create a temporary collection for each runspace
                $temp = "" | Select-Object PowerShell,Runspace,Computer
                $Temp.Computer = $Computer.computer
                $temp.PowerShell = $powershell
           
                #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
                $temp.Runspace = $powershell.BeginInvoke()
                Write-Verbose ("Adding {0} collection" -f $temp.Computer)
                $jobs.Add($temp) | Out-Null                
            }
		#Endregion InstallIIS
		#EndInstallIIS
        } 
             Else {
                $uiHash.StatusTextBox.Foreground = "Red"
                $uiHash.StatusTextBox.Text = "No server/s selected!"
        }
        
}



#Win2012ESUYear1LicenseCheck
Function Check-Win2012ESUYear1License{
 Write-Debug ("ComboBox {0}" -f $uiHash.RunOptionComboBox.Text)
        $selectedItems = $uiHash.Listview.SelectedItems
        If ($selectedItems.Count -gt 0) {
        $uiHash.ProgressBar.Maximum = $selectedItems.count
        $uiHash.Listview.ItemsSource | ForEach {$_.Status = $Null; $_.notes = $Null} 
	#Region Check-Win2012ESUYear1License
            $uiHash.RunButton.IsEnabled = $False
            $uiHash.StartImage.Source = "$pwd\Images\Start_locked.jpg"
            $uiHash.CancelButton.IsEnabled = $True
            $uiHash.CancelImage.Source = "$pwd\Images\Stop.jpg"            
            $uiHash.StatusTextBox.Foreground = "Black"
            $uiHash.StatusTextBox.Text = "Checking Win2012 ESU-Year1 License..."            
            $uiHash.StartTime = (Get-Date)
            
            [Float]$uiHash.ProgressBar.Value = 0
            $scriptBlock = {
                Param (
                    $Computer,
                    $uiHash,
                    $Path
                )               
                $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                    $uiHash.Listview.Items.EditItem($Computer)
                    $computer.Status = "Running"
					$computer.Notes = "Checking Win2012 ESU-Year1 License... please wait."
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh() 
                })                
                Set-Location $Path
                $Connection = (Test-Connection -ComputerName $Computer.computer -Count 1 -Quiet)
                $server = $Computer.computer 
                    
				$status, $resout = Invoke-Command -ComputerName $server -ScriptBlock{
					Try{

                        $LicChck = cmd.exe /c "cscript c:\Windows\System32\slmgr.vbs /dli c0a2ea62-12ad-435b-ab4f-c9bfab48dbc4" | Select-String -pattern "License"
                        $getStatus = $LicChck | ForEach-Object {$_.line.Substring($_.Line.LastIndexOf(':')+2)} | select -Last 1
                        if($getStatus -eq $null){$LicStatus = "Product Key not found!"}
                        Else{$LicStatus = $getStatus}
                        Return "Completed", "ESU-Year1 Status: $LicStatus"

						}Catch{return "Failed", $Error[0]}
						
					}

               $uiHash.ListView.Dispatcher.Invoke("Background",[action]{                    
                    $uiHash.Listview.Items.EditItem($Computer)
                    If ($Connection) {
                        $Computer.status = $status
                        $Computer.notes = $resout
                    } ElseIf (-Not $Connection) {
                        $Computer.Status = "Offline"
						$Computer.Notes = "Check Server"
                    } Else {
                        $Computer.Status = "Unknown"
						$Computer.Notes = "Check Server"
                    } 
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh()
                })
                $uiHash.ProgressBar.Dispatcher.Invoke("Normal",[action]{   
                    $uiHash.ProgressBar.value++  
                })
                $uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
                    #Check to see if find job
                    If ($uiHash.ProgressBar.value -eq $uiHash.ProgressBar.Maximum) { 
                        $End = New-Timespan $uihash.StartTime (Get-Date)                                             
                        $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)
                        $uiHash.RunButton.IsEnabled = $True
                        $uiHash.StartImage.Source = "$pwd\Images\Start.jpg"
                        $uiHash.CancelButton.IsEnabled = $False
                        $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.jpg"                    
                    }
                })  
                
            }

            Write-Verbose ("Creating runspace pool and session states")
            $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
            $runspaceHash.runspacepool = [runspacefactory]::CreateRunspacePool(1, $maxConcurrentJobs, $sessionstate, $Host)
            $runspaceHash.runspacepool.Open()   

            ForEach ($Computer in $selectedItems) {
                $uiHash.Listview.Items.EditItem($Computer)
                $computer.Status = "Pending..."
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh() 
                #Create the powershell instance and supply the scriptblock with the other parameters 
                $powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($computer).AddArgument($uiHash).AddArgument($Path)
           
                #Add the runspace into the powershell instance
                $powershell.RunspacePool = $runspaceHash.runspacepool
           
                #Create a temporary collection for each runspace
                $temp = "" | Select-Object PowerShell,Runspace,Computer
                $Temp.Computer = $Computer.computer
                $temp.PowerShell = $powershell
           
                #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
                $temp.Runspace = $powershell.BeginInvoke()
                Write-Verbose ("Adding {0} collection" -f $temp.Computer)
                $jobs.Add($temp) | Out-Null                
            }
		#Endregion Check-Win2012ESUYear1License
        } 
             Else {
                $uiHash.StatusTextBox.Foreground = "Red"
                $uiHash.StatusTextBox.Text = "No server/s selected!"
        }
        
}


#Install-Win2012ESUYear1ProductKey
Function Install-Win2012ESUYear1ProductKey{
 Write-Debug ("ComboBox {0}" -f $uiHash.RunOptionComboBox.Text)
        $selectedItems = $uiHash.Listview.SelectedItems
        If ($selectedItems.Count -gt 0) {
        $uiHash.ProgressBar.Maximum = $selectedItems.count
        $uiHash.Listview.ItemsSource | ForEach {$_.Status = $Null; $_.notes = $Null} 
	#Region Install-Win2012ESUYear1ProductKey
            $uiHash.RunButton.IsEnabled = $False
            $uiHash.StartImage.Source = "$pwd\Images\Start_locked.jpg"
            $uiHash.CancelButton.IsEnabled = $True
            $uiHash.CancelImage.Source = "$pwd\Images\Stop.jpg"            
            $uiHash.StatusTextBox.Foreground = "Black"
            $uiHash.StatusTextBox.Text = "Installing Win2012 ESU-Year1 Product Key..."            
            $uiHash.StartTime = (Get-Date)
            
            [Float]$uiHash.ProgressBar.Value = 0
            $scriptBlock = {
                Param (
                    $Computer,
                    $uiHash,
                    $Path
                )               
                $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                    $uiHash.Listview.Items.EditItem($Computer)
                    $computer.Status = "Running"
					$computer.Notes = "Installing Win2012 ESU-Year1 Product Key... please wait."
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh() 
                })                
                Set-Location $Path
                $Connection = (Test-Connection -ComputerName $Computer.computer -Count 1 -Quiet)
                $server = $Computer.computer 
                    
				$status, $resout = Invoke-Command -ComputerName $server -ScriptBlock{
					Try{


                        $InstallKey = cscript /Nologo "C:\Windows\System32\slmgr.vbs" /ipk NPY7Q-MJ7BR-VGQDH-XB82Q-H2483
                        if($InstallKey -eq $null){$InstallStatus = "Please retry!"}
                        Else{$InstallStatus = $InstallKey}
                        Return "Completed", "$InstallKey"

						}Catch{return "Failed", $Error[0]}
						
					}

               $uiHash.ListView.Dispatcher.Invoke("Background",[action]{                    
                    $uiHash.Listview.Items.EditItem($Computer)
                    If ($Connection) {
                        $Computer.status = $status
                        $Computer.notes = $resout
                    } ElseIf (-Not $Connection) {
                        $Computer.Status = "Offline"
						$Computer.Notes = "Check Server"
                    } Else {
                        $Computer.Status = "Unknown"
						$Computer.Notes = "Check Server"
                    } 
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh()
                })
                $uiHash.ProgressBar.Dispatcher.Invoke("Normal",[action]{   
                    $uiHash.ProgressBar.value++  
                })
                $uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
                    #Check to see if find job
                    If ($uiHash.ProgressBar.value -eq $uiHash.ProgressBar.Maximum) { 
                        $End = New-Timespan $uihash.StartTime (Get-Date)                                             
                        $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)
                        $uiHash.RunButton.IsEnabled = $True
                        $uiHash.StartImage.Source = "$pwd\Images\Start.jpg"
                        $uiHash.CancelButton.IsEnabled = $False
                        $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.jpg"                    
                    }
                })  
                
            }

            Write-Verbose ("Creating runspace pool and session states")
            $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
            $runspaceHash.runspacepool = [runspacefactory]::CreateRunspacePool(1, $maxConcurrentJobs, $sessionstate, $Host)
            $runspaceHash.runspacepool.Open()   

            ForEach ($Computer in $selectedItems) {
                $uiHash.Listview.Items.EditItem($Computer)
                $computer.Status = "Pending..."
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh() 
                #Create the powershell instance and supply the scriptblock with the other parameters 
                $powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($computer).AddArgument($uiHash).AddArgument($Path)
           
                #Add the runspace into the powershell instance
                $powershell.RunspacePool = $runspaceHash.runspacepool
           
                #Create a temporary collection for each runspace
                $temp = "" | Select-Object PowerShell,Runspace,Computer
                $Temp.Computer = $Computer.computer
                $temp.PowerShell = $powershell
           
                #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
                $temp.Runspace = $powershell.BeginInvoke()
                Write-Verbose ("Adding {0} collection" -f $temp.Computer)
                $jobs.Add($temp) | Out-Null                
            }
		#Endregion Install-Win2012ESUYear1ProductKey
        } 
             Else {
                $uiHash.StatusTextBox.Foreground = "Red"
                $uiHash.StatusTextBox.Text = "No server/s selected!"
        }
        
}



#Offboard Windows Defender
Function Offboard-WinDefend {
 Write-Debug ("ComboBox {0}" -f $uiHash.RunOptionComboBox.Text)
    $selectedItems = $uiHash.Listview.SelectedItems
    If ($selectedItems.Count -gt 0) {
        $uiHash.ProgressBar.Maximum = $selectedItems.count
        $uiHash.Listview.ItemsSource | ForEach {$_.Status = $Null; $_.notes = $Null} 
	
	#region Offboard-WinDefend
            $uiHash.RunButton.IsEnabled = $False
            $uiHash.StartImage.Source = "$pwd\Images\Start_locked.jpg"
            $uiHash.CancelButton.IsEnabled = $True
            $uiHash.CancelImage.Source = "$pwd\Images\Stop.jpg"            
            $uiHash.StatusTextBox.Foreground = "Black"
            $uiHash.StatusTextBox.Text = "Offboarding Windows Defender..."            
            $uiHash.StartTime = (Get-Date)
            [Float]$uiHash.ProgressBar.Value = 0
            $scriptBlock = {
                Param (
                    $Computer,
                    $uiHash,
                    $Path
                    
                )               
                          
                Set-Location $Path
                $Connection = (Test-Connection -ComputerName $Computer.computer -Count 1 -Quiet)
                $server = $Computer.computer
                if($Connection){

                $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                    $uiHash.Listview.Items.EditItem($Computer)
                    $computer.Status = "Running"
                    $computer.Notes = "Copying Script..."
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh()
                }) 
                    try{
                        $SPath = "\\unpiox03pn\IntelSystems\Lifecycle\Offboarding Defender"
                        $DPath = "\\" + $server + "\c$\temp\OffboardingDefender"
                        robocopy "$SPath" $DPath /S /Z /R:1 /NFL /NDL /NJH /NJS /nc /ns /np
                    
                        if($LASTEXITCODE -le 1 -or $LASTEXITCODE -eq 3 -or $LASTEXITCODE -eq 5 -or $LASTEXITCODE -eq 7){
                           $chkOBSCRpath = Test-Path -path "\\$server\c$\temp\OffboardingDefender\*"
                           if($chkOBSCRpath){
                            $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                            $uiHash.Listview.Items.EditItem($Computer)
                            $computer.Status = "Completed"
                            $computer.Notes = "Successfully copied"
                            $uiHash.Listview.Items.CommitEdit()
                            $uiHash.Listview.Items.Refresh()})
                            $StartOffboard = $True
                           }else{$getStatus = "Failed"; $getNotes = "Robocopy Failed!"; $StartOffboard = $False}
                        }Else{$getStatus = "Failed"; $getNotes = "LastExitCode = $LASTEXITCODE" + "$($error[0])"; $StartOffboard = $False}
                    }Catch{$getStatus = "Failed"; $getNotes = $($error[0])}
				
				#Region RunOffboardScriptProper
 					If($StartOffboard){
                        #$InstallDIR = (Get-ChildItem HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\* | 
                        #      % { Get-ItemProperty $_.PsPath } | Select DisplayName,InstallLocation | Where-Object DisplayName -like "Symantec Endpoint*").InstallLocation
                        #$connectiontimeout = 0
                       			$uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                                $uiHash.Listview.Items.EditItem($Computer)
                                $computer.Status = "Running"
                                $computer.Notes = "Offboarding Windows Defender... please wait."
                                $uiHash.Listview.Items.CommitEdit()
                                $uiHash.Listview.Items.Refresh()
                            })
                            try{
                                #$session = New-PSSession -ComputerName $server
                                #enter-pssession -session $session
                                $connectiontimeout = 0
                                $installationtimeout = 0
                                $runEXE, $runErr = Invoke-Command -cn $server -ScriptBlock{
                                    C:\temp\OffboardingDefender\*.cmd | Out-String
									#C:\temp\VRV\*.cmd
									#REPLACE WITH PROPER SCRIPT
                                    return $?, $error
                                } -ErrorVariable $EXEerr -ErrorAction SilentlyContinue
                                if($runEXE){#01
                                        
                                        $getStatus = "Completed"; $getNotes = "$runEXE"
                                }#01        

								
								else{
                                    $getStatus = "Failed"; $getNotes = $runErr
                                }
                            }Catch{$getStatus = "Failed"; $getNotes = $Error[0]}
                       
                    }
				
				}
				#EndRunOffboardScriptProper
                $uiHash.ListView.Dispatcher.Invoke("Background",[action]{                    
                    $uiHash.Listview.Items.EditItem($Computer)
                    $Computer.Notes = $null
                    If ($Connection) {
                        $Computer.Status = $getStatus
                        $Computer.Notes = $getNotes
                    } ElseIf (-Not $Connection) {
                        $Computer.Status = "Offline"
						$Computer.Notes = "Check Server"
                    } Else {
                        $Computer.Status = "Unknown"
						$Computer.Notes = "Check Server"
                    } 
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh()
                })
                $uiHash.ProgressBar.Dispatcher.Invoke("Normal",[action]{   
                    $uiHash.ProgressBar.value++  
                })
                $uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
                    #Check to see if find job
                    If ($uiHash.ProgressBar.value -eq $uiHash.ProgressBar.Maximum) { 
                        $End = New-Timespan $uihash.StartTime (Get-Date)                                             
                        $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)
                        $uiHash.RunButton.IsEnabled = $True
                        $uiHash.StartImage.Source = "$pwd\Images\Start.jpg"
                        $uiHash.CancelButton.IsEnabled = $False
                        $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.jpg"                    
                    }
                })  
            }
            Write-Verbose ("Creating runspace pool and session states")
            $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
            $runspaceHash.runspacepool = [runspacefactory]::CreateRunspacePool(1, $maxConcurrentJobs, $sessionstate, $Host)
            $runspaceHash.runspacepool.Open()   

            ForEach ($Computer in $selectedItems) {
                $uiHash.Listview.Items.EditItem($Computer)
                $computer.Status = "Pending..."
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh() 
                #Create the powershell instance and supply the scriptblock with the other parameters 
                $powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($computer).AddArgument($uiHash).AddArgument($Path)
           
                #Add the runspace into the powershell instance
                $powershell.RunspacePool = $runspaceHash.runspacepool
           
                #Create a temporary collection for each runspace
                $temp = "" | Select-Object PowerShell,Runspace,Computer
                $Temp.Computer = $Computer.computer
                $temp.PowerShell = $powershell
           
                #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
                $temp.Runspace = $powershell.BeginInvoke()
                Write-Verbose ("Adding {0} collection" -f $temp.Computer)
                $jobs.Add($temp) | Out-Null                
            }
    #Endregion Offboard-WinDefend
    #EndOffboard-WinDefend
			}Else {
                $uiHash.StatusTextBox.Foreground = "Red"
                $uiHash.StatusTextBox.Text = "No server/s selected!"
            }
}

#Check Windows Defender
Function Check-WinDefend{
 Write-Debug ("ComboBox {0}" -f $uiHash.RunOptionComboBox.Text)
        $selectedItems = $uiHash.Listview.SelectedItems
        If ($selectedItems.Count -gt 0) {
        $uiHash.ProgressBar.Maximum = $selectedItems.count
        $uiHash.Listview.ItemsSource | ForEach {$_.Status = $Null; $_.notes = $Null} 
	#Region CheckWinDefend
            $uiHash.RunButton.IsEnabled = $False
            $uiHash.StartImage.Source = "$pwd\Images\Start_locked.jpg"
            $uiHash.CancelButton.IsEnabled = $True
            $uiHash.CancelImage.Source = "$pwd\Images\Stop.jpg"            
            $uiHash.StatusTextBox.Foreground = "Black"
            $uiHash.StatusTextBox.Text = "Checking Windows Defender..."            
            $uiHash.StartTime = (Get-Date)
            
            [Float]$uiHash.ProgressBar.Value = 0
            $scriptBlock = {
                Param (
                    $Computer,
                    $uiHash,
                    $Path
                )               
                $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                    $uiHash.Listview.Items.EditItem($Computer)
                    $computer.Status = "Running"
					$computer.Notes = "Checking Windows Defender... please wait."
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh() 
                })                
                Set-Location $Path
                $Connection = (Test-Connection -ComputerName $Computer.computer -Count 1 -Quiet)
                $server = $Computer.computer 
                    
				$status, $resout = Invoke-Command -ComputerName $server -ScriptBlock{
					Try{
					Function Check-WinDefendServices
					{
					#$WinDefendSVCS = Get-Service -DisplayName "*Defender*" | select Name, Status | FT -AutoSize
					$WinDefendSVCS = Get-Service -Name WdNisSvc,Sense,Windefend | select Name, Status | FT -AutoSize
					$WinDefendSVCSOut = $WinDefendSVCS | Out-string
					if($WinDefendSVCS)
					{return $WinDefendSVCSOut}
					Else{return "No Windows Defender Services found!"}
					}

					$OS = (Get-WmiObject Win32_OperatingSystem).Version
					$WDSRV = Check-WinDefendServices

					if($OS -eq "6.3.9600")
					{
					$WDKey = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{560DAF36-DFFB-1F81-2966-2966F507C4D6}'
					$WinDef2012R2 = (Get-ItemProperty $WDKey | Select-Object DisplayName).DisplayName
						if($WinDef2012R2)
							{return "Completed", "OS: Win2012 R2 `n`n$WinDef2012R2 is installed. `n$WDSRV"}
						else
							{return "Completed", "OS: Win2012 R2 `n`nWindows Defender is NOT Installed. `n$WDSRV"}
					}
					elseif($OS -eq "10.0.14393")
						{
						$WinDefenderFeat = (Get-WindowsFeature -Name Windows-Defender | Select-Object Installed).Installed
						return "Completed", "OS: Win2016 `n`nWindows Defender is Installed: '$WinDefenderFeat' `n $WDSRV"
						}
					elseif($OS -eq "10.0.17763")
						{
						$WinDefenderFeat = (Get-WindowsFeature -Name Windows-Defender | Select-Object Installed).Installed
						return "Completed", "OS: Win2019 `n`nWindows Defender is Installed: '$WinDefenderFeat' `n $WDSRV"
						}
												
					Else{return "Failed", "WinDefend is NOT supported by OS!"}
				
						}Catch{return "Failed", $Error[0]}
						
					}

               $uiHash.ListView.Dispatcher.Invoke("Background",[action]{                    
                    $uiHash.Listview.Items.EditItem($Computer)
                    If ($Connection) {
                        $Computer.status = $status
                        $Computer.notes = $resout
                    } ElseIf (-Not $Connection) {
                        $Computer.Status = "Offline"
						$Computer.Notes = "Check Server"
                    } Else {
                        $Computer.Status = "Unknown"
						$Computer.Notes = "Check Server"
                    } 
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh()
                })
                $uiHash.ProgressBar.Dispatcher.Invoke("Normal",[action]{   
                    $uiHash.ProgressBar.value++  
                })
                $uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
                    #Check to see if find job
                    If ($uiHash.ProgressBar.value -eq $uiHash.ProgressBar.Maximum) { 
                        $End = New-Timespan $uihash.StartTime (Get-Date)                                             
                        $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)
                        $uiHash.RunButton.IsEnabled = $True
                        $uiHash.StartImage.Source = "$pwd\Images\Start.jpg"
                        $uiHash.CancelButton.IsEnabled = $False
                        $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.jpg"                    
                    }
                })  
                
            }

            Write-Verbose ("Creating runspace pool and session states")
            $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
            $runspaceHash.runspacepool = [runspacefactory]::CreateRunspacePool(1, $maxConcurrentJobs, $sessionstate, $Host)
            $runspaceHash.runspacepool.Open()   

            ForEach ($Computer in $selectedItems) {
                $uiHash.Listview.Items.EditItem($Computer)
                $computer.Status = "Pending..."
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh() 
                #Create the powershell instance and supply the scriptblock with the other parameters 
                $powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($computer).AddArgument($uiHash).AddArgument($Path)
           
                #Add the runspace into the powershell instance
                $powershell.RunspacePool = $runspaceHash.runspacepool
           
                #Create a temporary collection for each runspace
                $temp = "" | Select-Object PowerShell,Runspace,Computer
                $Temp.Computer = $Computer.computer
                $temp.PowerShell = $powershell
           
                #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
                $temp.Runspace = $powershell.BeginInvoke()
                Write-Verbose ("Adding {0} collection" -f $temp.Computer)
                $jobs.Add($temp) | Out-Null                
            }
		#Endregion CheckWinDefend
		#EndCheckWinDefend
        } 
             Else {
                $uiHash.StatusTextBox.Foreground = "Red"
                $uiHash.StatusTextBox.Text = "No server/s selected!"
        }
        
}

#Windows Defender Uninstall
Function WinDefendUninstall{
 Write-Debug ("ComboBox {0}" -f $uiHash.RunOptionComboBox.Text)
        $selectedItems = $uiHash.Listview.SelectedItems
        If ($selectedItems.Count -gt 0) {
        $uiHash.ProgressBar.Maximum = $selectedItems.count
        $uiHash.Listview.ItemsSource | ForEach {$_.Status = $Null; $_.notes = $Null} 
	#region WinDefendUninstall
            $uiHash.RunButton.IsEnabled = $False
            $uiHash.StartImage.Source = "$pwd\Images\Start_locked.jpg"
            $uiHash.CancelButton.IsEnabled = $True
            $uiHash.CancelImage.Source = "$pwd\Images\Stop.jpg"            
            $uiHash.StatusTextBox.Foreground = "Black"
            $uiHash.StatusTextBox.Text = "Uninstalling Windows Defender..."            
            $uiHash.StartTime = (Get-Date)
            
            [Float]$uiHash.ProgressBar.Value = 0
            $scriptBlock = {
                Param (
                    $Computer,
                    $uiHash,
                    $Path
                )               
                $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                    $uiHash.Listview.Items.EditItem($Computer)
                    $computer.Status = "Running"
					$computer.Notes = "Uninstalling Windows Defender... please wait."
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh() 
                })                
                Set-Location $Path
                $Connection = (Test-Connection -ComputerName $Computer.computer -Count 1 -Quiet)
                $server = $Computer.computer 
                    
				$status, $resout = Invoke-Command -ComputerName $server -ScriptBlock{
					Try{
					$OS = (Get-WmiObject Win32_OperatingSystem).Version
					if($OS -eq "6.3.9600")
					{
						$WindowsDefender = Get-WmiObject -Class Win32_Product | Where-Object{$_.Name -eq “Microsoft Defender for Endpoint”}
						if($WindowsDefender)
							{
							#Uninstall_Script.
							$WDUninstall = $WindowsDefender.uninstall()
							#$WDRes = $WDUninstall | Out-string
							return "Completed", "OS: Win2012 R2 `n`nWindows Defender was uninstalled. `nReboot needed."
							}
						else
							{return "Completed", "OS: Win2012 R2 `n`nWindows Defender is NOT Installed."}
					}
					elseif($OS -eq "10.0.14393")
						{
						#Uninstall_Script.
						$WDUninstall = Uninstall-WindowsFeature -Name Windows-Defender | Select-Object Success, RestartNeeded, ExitCode, FeatureResult | FL
						$WDRes = $WDUninstall | Out-String
						return "Completed", "OS: Win2016 `n`nWindows Defender Uninstall Result: $WDRes"
						}
					elseif($OS -eq "10.0.17763")
						{
						#Uninstall_Script.
						$WDUninstall = Uninstall-WindowsFeature -Name Windows-Defender | Select-Object Success, RestartNeeded, ExitCode, FeatureResult | FL
						$WDRes = $WDUninstall | Out-String
						return "Completed", "OS: Win2019 `n`nWindows Defender Uninstall Result: $WDRes"
						}
												
					Else{return "Failed", "WinDefend is NOT supported by OS!"}

				
						}Catch{return "Failed", $Error[0]}
						
					}

               $uiHash.ListView.Dispatcher.Invoke("Background",[action]{                    
                    $uiHash.Listview.Items.EditItem($Computer)
                    If ($Connection) {
                        $Computer.status = $status
                        $Computer.notes = $resout
                    } ElseIf (-Not $Connection) {
                        $Computer.Status = "Offline"
						$Computer.Notes = "Check Server"
                    } Else {
                        $Computer.Status = "Unknown"
						$Computer.Notes = "Check Server"
                    } 
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh()
                })
                $uiHash.ProgressBar.Dispatcher.Invoke("Normal",[action]{   
                    $uiHash.ProgressBar.value++  
                })
                $uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
                    #Check to see if find job
                    If ($uiHash.ProgressBar.value -eq $uiHash.ProgressBar.Maximum) { 
                        $End = New-Timespan $uihash.StartTime (Get-Date)                                             
                        $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)
                        $uiHash.RunButton.IsEnabled = $True
                        $uiHash.StartImage.Source = "$pwd\Images\Start.jpg"
                        $uiHash.CancelButton.IsEnabled = $False
                        $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.jpg"                    
                    }
                })  
                
            }

            Write-Verbose ("Creating runspace pool and session states")
            $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
            $runspaceHash.runspacepool = [runspacefactory]::CreateRunspacePool(1, $maxConcurrentJobs, $sessionstate, $Host)
            $runspaceHash.runspacepool.Open()   

            ForEach ($Computer in $selectedItems) {
                $uiHash.Listview.Items.EditItem($Computer)
                $computer.Status = "Pending..."
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh() 
                #Create the powershell instance and supply the scriptblock with the other parameters 
                $powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($computer).AddArgument($uiHash).AddArgument($Path)
           
                #Add the runspace into the powershell instance
                $powershell.RunspacePool = $runspaceHash.runspacepool
           
                #Create a temporary collection for each runspace
                $temp = "" | Select-Object PowerShell,Runspace,Computer
                $Temp.Computer = $Computer.computer
                $temp.PowerShell = $powershell
           
                #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
                $temp.Runspace = $powershell.BeginInvoke()
                Write-Verbose ("Adding {0} collection" -f $temp.Computer)
                $jobs.Add($temp) | Out-Null                
            }
		#endregion WinDefendUninstall
		#EndWinDefendUninstall
        } 
             Else {
                $uiHash.StatusTextBox.Foreground = "Red"
                $uiHash.StatusTextBox.Text = "No server/s selected!"
        }
        
}


#Build the GUI
[xml]$xaml = @"
<Window
    xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation'
    xmlns:x='http://schemas.microsoft.com/winfx/2006/xaml'
    x:Name='Window' Title='Windows Deployment Services PowerShell-GUI Tool v4.2' WindowStartupLocation = 'CenterScreen' 
    Width = '880' Height = '575' ShowInTaskbar = 'True'>
    <Window.Background>
        <LinearGradientBrush StartPoint='0,0' EndPoint='0,1'>
            <LinearGradientBrush.GradientStops> <GradientStop Color='#C4CBD8' Offset='0' /> <GradientStop Color='#E6EAF5' Offset='0.2' /> 
            <GradientStop Color='#CFD7E2' Offset='0.9' /> <GradientStop Color='#C4CBD8' Offset='1' /> </LinearGradientBrush.GradientStops>
        </LinearGradientBrush>
    </Window.Background> 
    <Window.Resources>        
        <DataTemplate x:Key="HeaderTemplate">
            <DockPanel>
                <TextBlock FontSize="10" Foreground="Green" FontWeight="Bold" >
                    <TextBlock.Text>
                        <Binding/>
                    </TextBlock.Text>
                </TextBlock>
            </DockPanel>
        </DataTemplate>            
    </Window.Resources>    
    <Grid x:Name = 'Grid' ShowGridLines = 'false'>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>
        <Grid.RowDefinitions>
            <RowDefinition Height = 'Auto'/>
            <RowDefinition Height = 'Auto'/>
            <RowDefinition Height = '*'/>
            <RowDefinition Height = 'Auto'/>
            <RowDefinition Height = 'Auto'/>
            <RowDefinition Height = 'Auto'/>
        </Grid.RowDefinitions>    
        <Menu Width = 'Auto' HorizontalAlignment = 'Stretch' Grid.Row = '0'>
        <Menu.Background>
            <LinearGradientBrush StartPoint='0,0' EndPoint='0,1'>
                <LinearGradientBrush.GradientStops> <GradientStop Color='#C4CBD8' Offset='0' /> <GradientStop Color='#E6EAF5' Offset='0.2' /> 
                <GradientStop Color='#CFD7E2' Offset='0.9' /> <GradientStop Color='#C4CBD8' Offset='1' /> </LinearGradientBrush.GradientStops>
            </LinearGradientBrush>
        </Menu.Background>
            <MenuItem x:Name = 'FileMenu' Header = '_File'>
                <MenuItem x:Name = 'RunMenu' Header = '_Run' ToolTip = 'Initiate Run operation' InputGestureText ='F5'> </MenuItem>
                <MenuItem x:Name = 'GenerateReportMenu' Header = 'Generate R_eport' ToolTip = 'Generate Report' InputGestureText ='F8'/>
                <Separator />            
                <MenuItem x:Name = 'OptionMenu' Header = '_Options' ToolTip = 'Open up options window.' InputGestureText ='Ctrl+O'/>
                <Separator />
                <MenuItem x:Name = 'ExitMenu' Header = 'E_xit' ToolTip = 'Exits the utility.' InputGestureText ='Ctrl+E'/>
            </MenuItem>
            <MenuItem x:Name = 'EditMenu' Header = '_Edit'>
                <MenuItem x:Name = 'SelectAllMenu' Header = 'Select _All' ToolTip = 'Selects all rows.' InputGestureText ='Ctrl+A'/>               
                <Separator />
                <MenuItem x:Name = 'ClearErrorMenu' Header = 'Clear ErrorLog' ToolTip = 'Clears error log.'> </MenuItem>                
                <MenuItem x:Name = 'ClearAllMenu' Header = 'Clear All' ToolTip = 'Clears everything on the WSUS utility.'/>
            </MenuItem>
            <MenuItem x:Name = 'ActionMenu' Header = '_View'>
                <MenuItem Header = 'Reports'>
                    <MenuItem x:Name = 'ClearPrepReportMenu' Header = 'Clear Pre-Patch Report' ToolTip = 'Clears the current pre-patch report.'/>
                    <MenuItem x:Name = 'ClearAuditReportMenu' Header = 'Clear Audit Report' ToolTip = 'Clears the current audit report.'/>
                    <MenuItem x:Name = 'ClearPostPatchReportMenu' Header = 'Clear Post-Patch Report' ToolTip = 'Clears the current post-patch report.'/>
                    <MenuItem x:Name = 'ClearInstallReportMenu' Header = 'Clear Install Report' ToolTip = 'Clears the current report.'/>                   
                    <MenuItem x:Name = 'ClearInstalledUpdateMenu' Header = 'Clear Installed Update Report' ToolTip = 'Clears the installed update report.'/>
                    <MenuItem x:Name = 'ClearServicesReportMenu' Header = 'Clear Services Report' ToolTip = 'Clears "Stopped Services" report.'/>
                    <MenuItem x:Name = 'ClearESUReportMenu' Header = 'Clear ESU License Report' ToolTip = 'Clears 2k8 ESU License Activation report.'/>
                </MenuItem>
                <MenuItem Header = 'Server List'>
                    <MenuItem x:Name = 'ClearServerListMenu' Header = 'Clear Server List' ToolTip = 'Clears the server list.'/>
                    <MenuItem x:Name = 'OfflineHostsMenu' Header = 'Remove Offline Servers' ToolTip = 'Removes all offline hosts from Server List'/>  
                    <MenuItem x:Name = 'CompletedHostsMenu' Header = 'Remove "Patch Completed" Servers' ToolTip = 'Removes all patch completed servers from Dashboard'/>
                    <MenuItem x:Name = 'SEPMenu' Header = 'Remove SEP version "14.2.5323"' ToolTip = 'Removes servers with latest SEP version from Dashboard'/> 
                    <MenuItem x:Name = 'ESUMenu' Header = 'Remove "Licensed" 2k8' ToolTip = 'Removes servers with Licensed ESU from Dashboard'/>                 
                    <MenuItem x:Name = 'ResetDataMenu' Header = 'Reset Computer List Data' ToolTip = 'Resets the audit and patch data on Server List'/>
                </MenuItem> 
            <Separator />                           
            <MenuItem x:Name = 'HostListMenu' Header = 'Create Host List' ToolTip = 'Creates a list of all servers and saves to a text file.'/>
                <MenuItem x:Name = 'ServerListReportMenu' Header = 'Create Server List Report' 
                ToolTip = 'Creates a CSV file listing the current Server List.'/>
                <Separator/>
                <MenuItem x:Name = 'ViewErrorMenu' Header = 'View ErrorLog' ToolTip = 'View error log.'/>            
            </MenuItem>            
            <MenuItem x:Name = 'HelpMenu' Header = '_Help'>
                <MenuItem x:Name = 'AboutMenu' Header = '_About' ToolTip = 'Show the current version and other information.'> </MenuItem>
                <MenuItem x:Name = 'HelpFileMenu' Header = 'WSUS Utility _Help' 
                ToolTip = 'Displays a help file to use the WSUS Utility.' InputGestureText ='F1'> </MenuItem>
            </MenuItem>            
        </Menu>
        <ToolBarTray Grid.Row = '1' Grid.Column = '0'>
        <ToolBarTray.Background>
            <LinearGradientBrush StartPoint='0,0' EndPoint='0,1'>
                <LinearGradientBrush.GradientStops> <GradientStop Color='#C4CBD8' Offset='0' /> <GradientStop Color='#E6EAF5' Offset='0.2' /> 
                <GradientStop Color='#CFD7E2' Offset='0.9' /> <GradientStop Color='#C4CBD8' Offset='1' /> </LinearGradientBrush.GradientStops>
            </LinearGradientBrush>        
        </ToolBarTray.Background>
            <ToolBar Background = 'Transparent' Band = '1' BandIndex = '1'>
                <Button x:Name = 'RunButton' Width = 'Auto' ToolTip = 'Performs action against all servers in the server list based on checked radio button.'>
                    <Image x:Name = 'StartImage' Source = '$Pwd\Images\Start.jpg'/>
                </Button>         
                <Separator Background = 'Black'/>   
                <Button x:Name = 'CancelButton' Width = 'Auto' ToolTip = 'Cancels currently running operations.' IsEnabled = 'False'>
                    <Image x:Name = 'CancelImage' Source = '$pwd\Images\Stop_locked.jpg' />
                </Button>
                <Separator Background = 'Black'/>
                <ComboBox x:Name = 'RunOptionComboBox' Width = 'Auto' IsReadOnly = 'True'
                SelectedIndex = '0'>
                    <TextBlock> Pre-Patch </TextBlock>
                    <TextBlock> Audit Patches </TextBlock>
                    <TextBlock> Install Patches </TextBlock>
                    <TextBlock> Post-Patch </TextBlock>
                    <TextBlock> Check Pending Reboot </TextBlock>
                    <TextBlock> PING/BootUpTime Check </TextBlock>
                    <TextBlock> Services Check </TextBlock>
                    <TextBlock> Reboot Systems </TextBlock>
                    <TextBlock> Check C-drive </TextBlock>
                    <TextBlock> Check CHEF Version </TextBlock>
                </ComboBox>                
            </ToolBar>
            <ToolBar Background = 'Transparent' Band = '1' BandIndex = '1'>
                <Button x:Name = 'GenerateReportButton' Width = 'Auto' ToolTip = 'Generates a report based on user selection.'>
                    <Image Source = '$pwd\Images\Gen_Report.gif' />
                </Button>            
                <ComboBox x:Name = 'ReportComboBox' Width = 'Auto' IsReadOnly = 'True' SelectedIndex = '0'>
                    <TextBlock> Prep CSV Report </TextBlock>
                    <TextBlock> Prep UI Report </TextBlock>
                    <TextBlock> PostPatch CSV Report </TextBlock>
                    <TextBlock> PostPatch UI Report </TextBlock>
                    <TextBlock> Audit CSV Report </TextBlock>
                    <TextBlock> Audit UI Report </TextBlock>
                    <TextBlock> Install CSV Report </TextBlock>
                    <TextBlock> Install UI Report </TextBlock>
                    <TextBlock> Installed Updates CSV Report </TextBlock>
                    <TextBlock> Installed Updates UI Report </TextBlock>
                    <TextBlock> Services CSV Report </TextBlock> 
                    <TextBlock> Services UI Report </TextBlock>                                                           
                    <TextBlock> Host File List </TextBlock>
                    <TextBlock> Computer List Report </TextBlock>
                    <TextBlock> Error UI Report </TextBlock>
                    <TextBlock> ESU License UI Report </TextBlock>
                    <TextBlock> ESU License CSV Report </TextBlock>
                </ComboBox>              
                <Separator Background = 'Black'/>
            </ToolBar>
            <ToolBar Background = 'Transparent' Band = '1' BandIndex = '1'>
                <Button x:Name = 'BrowseFileButton' Width = 'Auto' 
                ToolTip = 'Open a file dialog to select a host file. Upon selection, the contents will be loaded into Server list.'>
                    <Image Source = '$pwd\Images\BrowseFile.gif' />
                </Button>    
                <Button x:Name = 'LoadADButton' Width = 'Auto' 
                ToolTip = 'Creates a list of computers from Active Directory to use in Server List.'>
                    <Image Source = '$pwd\Images\ActiveDirectory.gif' />
                </Button>                                      
                <Separator Background = 'Black'/>
            </ToolBar>             
        </ToolBarTray>
        <Grid Grid.Row = '2' Grid.Column = '0' ShowGridLines = 'false'>  
            <Grid.Resources>
                <Style x:Key="AlternatingRowStyle" TargetType="{x:Type Control}" >
                    <Setter Property="Background" Value="LightGray"/>
                    <Setter Property="Foreground" Value="Black"/>
                    <Style.Triggers>
                        <Trigger Property="ItemsControl.AlternationIndex" Value="1">                            
                            <Setter Property="Background" Value="White"/>
                            <Setter Property="Foreground" Value="Black"/>                                
                        </Trigger>                            
                    </Style.Triggers>
                </Style>                    
            </Grid.Resources>                  
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="10"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="10"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="10"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <Grid.RowDefinitions>
                <RowDefinition Height = 'Auto'/>
                <RowDefinition Height = 'Auto'/>
                <RowDefinition Height = '*'/>
                <RowDefinition Height = '*'/>
                <RowDefinition Height = 'Auto'/>
                <RowDefinition Height = 'Auto'/>
                <RowDefinition Height = 'Auto'/>
            </Grid.RowDefinitions> 
            <GroupBox Header = "Computer List" Grid.Column = '0' Grid.Row = '2' Grid.ColumnSpan = '11' Grid.RowSpan = '3'>
                <Grid Width = 'Auto' Height = 'Auto' ShowGridLines = 'false'>
                <ListView x:Name = 'Listview' AllowDrop = 'True' AlternationCount="2" ItemContainerStyle="{StaticResource AlternatingRowStyle}"
                ToolTip = 'Server List that displays all information regarding statuses of servers and patches.'>
                    <ListView.View>
                        <GridView x:Name = 'GridView' AllowsColumnReorder = 'True' ColumnHeaderTemplate="{StaticResource HeaderTemplate}">
                            <GridViewColumn x:Name = 'ComputerColumn' Width = '110' DisplayMemberBinding = '{Binding Path = Computer}' Header='Computer'/>
                            <GridViewColumn x:Name = 'AuditedColumn' Width = '110' DisplayMemberBinding = '{Binding Path = Audited}' Header='Audited'/>                    
                            <GridViewColumn x:Name = 'InstalledColumn' Width = '110' DisplayMemberBinding = '{Binding Path = Installed}' Header='Installed' />                    
                            <GridViewColumn x:Name = 'InstallErrorColumn' Width = '110' DisplayMemberBinding = '{Binding Path = InstallErrors}' Header='InstallErrors'/>  
                            <GridViewColumn x:Name = 'ServicesColumn' Width = '115' DisplayMemberBinding = '{Binding Path = Services}' Header='NonRunningServices'/>   
                            <GridViewColumn x:Name = 'StatusColumn' Width = '200' DisplayMemberBinding = '{Binding Path = Status}' Header='Status'/>                                              
                            <GridViewColumn x:Name = 'NotesColumn' Width = '300' DisplayMemberBinding = '{Binding Path = Notes}' Header='Notes'/>                    
                        </GridView>
                    </ListView.View>
                    <ListView.ContextMenu>
                        <ContextMenu x:Name = 'ListViewContextMenu'>
                            <MenuItem x:Name = 'AddServerMenu' Header = 'Add Server' InputGestureText ='Ctrl+S'/>               
                            <MenuItem x:Name = 'RemoveServerMenu' Header = 'Remove Server' InputGestureText ='Ctrl+D'/>
                            <Separator />
                             <MenuItem x:Name = 'GetMenu' Header = 'Run' > 
                                <MenuItem x:Name = 'PrePatchMenu' Header = 'Pre-Patch Check'/>
                                <MenuItem x:Name = 'PostPatchMenu' Header = 'Post-Patch Check'/>
                                <Separator />
                                <MenuItem x:Name = 'SpectreMenu' Header = 'Check Spectre Meltdown'/>
                                <MenuItem x:Name = 'ExchangeMenu' Header = 'Check Exchange Services'/>
                                <MenuItem x:Name = 'TSMMenu' Header = 'Check TSM service'/>
								<MenuItem x:Name = 'MsOfficeMenu' Header = 'Check MS Office version'/>
                                <MenuItem x:Name = 'EDGECheckMenu' Header = 'Check EDGE'/>
                                <MenuItem x:Name = 'IISCheckMenu' Header = 'Check IIS'/>
                                <Separator />
								<MenuItem x:Name = 'IISInstallMenu' Header = 'Install IIS'/>
                            </MenuItem>

                            <MenuItem x:Name = 'RemediateMenu' Header = 'Remediate' > 
                                <MenuItem x:Name = 'ChefPatchingMenu' Header = 'Chef Patching'>
                                    <MenuItem x:Name = 'ChefRunMenu' Header = 'Run Chef'/>
                                    <MenuItem x:Name = 'ChefRunNRMenu' Header = 'Run Chef(No Reboot)'/>
                                </MenuItem>
                                <MenuItem x:Name = 'WUPatchingMenu' Header = 'WU Patching'>
                                    <MenuItem x:Name = 'WindowsUpdateMenu' Header = 'Run WindowsUpdate'/>
                                    <MenuItem x:Name = 'WindowsUpdateNRMenu' Header = 'Run WU(No Reboot)'/>
                                    <MenuItem x:Name = 'WindowsUpdateSRMenu' Header = 'Run WU(Scheduled Reboot)'/>
                                </MenuItem>
                            </MenuItem>
                            <MenuItem x:Name = 'RunScriptMenu' Header = 'Ad hoc' >
                                <MenuItem x:Name = 'RunPSMenu' Header = 'Powershell'/>
                                <MenuItem x:Name = 'RunBatchMenu' Header = 'Batch'/>
                            </MenuItem>
                            <MenuItem x:Name = 'FailoverClusterMenu' Header = 'Failover Cluster'>
                                <MenuItem x:Name = 'ChkNameNodeMenu' Header = 'Check Clustername/Node(s)'/>
                                <MenuItem x:Name = 'RunMoveResourceMenu' Header = 'Move Resource(s)'/>
                            </MenuItem>
                            <MenuItem x:Name = 'RobocopyMenu' Header = 'RoboCopy'/>
                            <MenuItem x:Name = 'RebootMenu' Header = 'Reboot Server'/>
							<MenuItem x:Name = 'CHEFCommandsMenu' Header = 'CHEF Patching Scripts' >
                                <MenuItem x:Name = 'ChefRunAbortMenu' Header = 'Abort Chef'/>
								<MenuItem x:Name = 'ChefLogsMenu' Header = 'Check Chef Logs'/>
                                <MenuItem x:Name = 'CheckPatchLogMenu' Header = 'Check Patch Log Tail 25'/>
                               	<MenuItem x:Name = 'RunChefClientMenu' Header = 'Run Chef-Client'/>
                                <Separator />
                                <MenuItem x:Name = 'CheckPatchingRebFlagMenu' Header = 'Check RebootRequested.TXT'/>
								<MenuItem x:Name = 'DelRebootFlagMenu' Header = 'Delete RebootRequested.TXT'/>
                                <Separator />
                                <MenuItem x:Name = 'CheckRubyProcessMenu' Header = 'Check Ruby Interpreter Process'/>
                                <MenuItem x:Name = 'ENDRubyProcessMenu' Header = 'End Ruby Interpreter Process'/>
                                <Separator />
                                <MenuItem x:Name = 'CheckCHEFenvVarMenu' Header = 'Check CHEF Environment Variables'/>
                                <MenuItem x:Name = 'AddMissingCHEFEnvVarMenu' Header = 'Add Missing CHEF Environment Variables'/>
                                </MenuItem>
                            <MenuItem x:Name = 'MicrosoftDotNetMenu' Header = 'Microsoft .NET' >
                               	<MenuItem x:Name = 'DotNetRuntimeOnOSMenu' Header = 'Check Latest .NET Version'/>
								<MenuItem x:Name = 'DotnetchkAllmenu' Header = 'Check All .NET Versions'/>								
                                <MenuItem x:Name = 'ASPVerMenu' Header = 'Check ASP.NET Core Version'/>
                                <MenuItem x:Name = 'DotNetAddOnsMenu' Header = 'Check .NET AddOns'/>
							    </MenuItem>
                            <MenuItem x:Name = 'SymantecEndpointProtectionMenu' Header = 'Symantec Endpoint Protection' >
                                <MenuItem x:Name = 'SEPHealthCheckMenu' Header = 'SEP Health Check'/>
                                <MenuItem x:Name = 'SEPverMenu' Header = 'Check SEP Version'/>
                                <MenuItem x:Name = 'SEPVirusDefFoldersCheckMenu' Header = 'Check SEP Virus Def Folder(s)'/>
                                <Separator />
                                <MenuItem x:Name = 'RepairSEPMenu' Header = 'Repair SEP Client'/>
                                <MenuItem x:Name = 'EndSepScanServiceMenu' Header = 'End SEP Scan Service'/>
                                <MenuItem x:Name = 'SEPOLDVirusDefDELETEMenu' Header = 'Delete OLD SEP Virus Def Folder(s)'/>
                                <MenuItem x:Name = 'StartSEPMasterServiceMenu' Header = 'Start SEP Master Service'/>
                                </MenuItem>
                            <MenuItem x:Name = 'TrendMicroMenu' Header = 'Trend Micro' >
                                <MenuItem x:Name = 'TrendMicroDSAHealthCheckMenu' Header = 'Trend Micro DSA Health Check'/>
                                <Separator />
                                <MenuItem x:Name = 'TMEDRProdAgentDeployentMenu' Header = 'Trend Micro EDR-PROD Agent Deployment'/>
                                </MenuItem>
                            <MenuItem x:Name = 'Win2012ESUYear1Menu' Header = 'Win2012 ESU-Year1' >
                               	<MenuItem x:Name = 'Win2012ESUYear1LicenseCheckMenu' Header = 'Check Win2012 ESU-Year1 License'/>
                                <Separator />
                                <MenuItem x:Name = 'Win2012ESUYear1ProductKeyInstallMenu' Header = 'Install Win2012 ESU-Year1 Product Key'/>
                                </MenuItem>
                            <MenuItem x:Name = 'WindowsDefenderMenu' Header = 'Windows Defender' >
                               	<MenuItem x:Name = 'WinDefendCheckMenu' Header = 'Check Windows Defender'/>
                                <Separator />
								<MenuItem x:Name = 'WinDefendOffboardMenu' Header = 'Offboard Windows Defender'/>
								<MenuItem x:Name = 'WinDefendUninstallMenu' Header = 'Uninstall Windows Defender'/>
                                </MenuItem>
                            <MenuItem x:Name = 'WUTroubleshootingMenu' Header = 'Windows Update Troubleshooting' >
                               	<MenuItem x:Name = 'WinUpdateFixMenu' Header = 'Windows Update Fix (Single Server with Output)'/>
								<MenuItem x:Name = 'WinUpdateFixMultiServersMenu' Header = 'Windows Update Fix (Multiple Servers)'/>
								<MenuItem x:Name = 'WUAUCLTMenu' Header = 'WU Automatic Update Client (WUAUCLT)' >
                               	<MenuItem x:Name = 'DetectNowMenu' Header = 'Run Detect Now'/> 
                               	<MenuItem x:Name = 'ResetAuthorizationMenu' Header = 'Run Reset Authorization'/>
                                </MenuItem>
                                <MenuItem x:Name = 'WindowsUpdateServiceMenu' Header = 'Windows Update Service' > 
                               		<MenuItem x:Name = 'WUStopServiceMenu' Header = 'Stop Service' />
                               		<MenuItem x:Name = 'WUStartServiceMenu' Header = 'Start Service' />
                               		<MenuItem x:Name = 'WURestartServiceMenu' Header = 'Restart Service' />
                            	    </MenuItem> 
                                </MenuItem>
                            <MenuItem x:Name = 'WUInformationMenu' Header = 'Windows Update Information' >
								<MenuItem x:Name = 'InstalledUpdatesMenu' Header = 'Installed Updates' >
                                <MenuItem x:Name = 'GUIInstalledUpdatesMenu' Header = 'Get Installed Updates'/>
                            	</MenuItem>
								<MenuItem x:Name = 'WindowsUpdateLogMenu' Header = 'WindowsUpdateLog' > 
                                	<MenuItem x:Name = 'EntireLogMenu' Header = 'View Entire Log'/>
                                	<MenuItem x:Name = 'Last25LogMenu' Header = 'View Last 25' />
                                	<MenuItem x:Name = 'Last50LogMenu' Header = 'View Last 50'/>
                                	<MenuItem x:Name = 'Last100LogMenu' Header = 'View Last 100'/>
                            	    </MenuItem>
								<MenuItem x:Name = 'CBSlogsMenu' Header = 'Check CBS logs'/>
								<MenuItem x:Name = 'WSUSRegMenu' Header = 'Check WSUS Registry Settings'/>
							    </MenuItem>
                        </ContextMenu>
                    </ListView.ContextMenu>            
                </ListView>                
                </Grid>
            </GroupBox>                                    
        </Grid>        
        <ProgressBar x:Name = 'ProgressBar' Grid.Row = '3' Height = '20' ToolTip = 'Displays progress of current action via a graphical progress bar.'/>   
        <TextBox x:Name = 'StatusTextBox' Grid.Row = '4' ToolTip = 'Displays current status of operation'> Waiting for Action... </TextBox>                           
    </Grid>   
</Window>
"@ 

#region Load XAML into PowerShell
$reader=(New-Object System.Xml.XmlNodeReader $xaml)
$uiHash.Window=[Windows.Markup.XamlReader]::Load( $reader )
#endregion
 
#region Background runspace to clean up jobs
$jobCleanup.Flag = $True
$newRunspace =[runspacefactory]::CreateRunspace()
$newRunspace.ApartmentState = "STA"
$newRunspace.ThreadOptions = "ReuseThread"          
$newRunspace.Open()
$newRunspace.SessionStateProxy.SetVariable("uiHash",$uiHash)          
$newRunspace.SessionStateProxy.SetVariable("jobCleanup",$jobCleanup)     
$newRunspace.SessionStateProxy.SetVariable("jobs",$jobs) 
$jobCleanup.PowerShell = [PowerShell]::Create().AddScript({
    #Routine to handle completed runspaces
    Do {    
        Foreach($runspace in $jobs) {
            If ($runspace.Runspace.isCompleted) {
                $runspace.powershell.EndInvoke($runspace.Runspace) | Out-Null
                $runspace.powershell.dispose()
                $runspace.Runspace = $null
                $runspace.powershell = $null               
            } 
        }
        #Clean out unused runspace jobs
        $temphash = $jobs.clone()
        $temphash | Where {
            $_.runspace -eq $Null
        } | ForEach {
            Write-Verbose ("Removing {0}" -f $_.computer)
            $jobs.remove($_)
        }        
        Start-Sleep -Seconds 1     
    } while ($jobCleanup.Flag)
})
$jobCleanup.PowerShell.Runspace = $newRunspace
$jobCleanup.Thread = $jobCleanup.PowerShell.BeginInvoke()  
#endregion

#region Connect to all controls
$uiHash.GenerateReportMenu = $uiHash.Window.FindName("GenerateReportMenu")
$uiHash.ClearAuditReportMenu = $uiHash.Window.FindName("ClearAuditReportMenu")
$uiHash.ClearPostPatchReportMenu = $uiHash.Window.FindName("ClearPostPatchReportMenu")
$uiHash.ClearPrepReportMenu = $uiHash.Window.FindName("ClearPrepReportMenu")
$uiHash.ClearESUReportMenu = $uiHash.Window.FindName("ClearESUReportMenu")
$uiHash.ClearInstallReportMenu = $uiHash.Window.FindName("ClearInstallReportMenu")
$uiHash.SelectAllMenu = $uiHash.Window.FindName("SelectAllMenu")
$uiHash.OptionMenu = $uiHash.Window.FindName("OptionMenu")
$uiHash.WUStopServiceMenu = $uiHash.Window.FindName("WUStopServiceMenu")
$uiHash.WUStartServiceMenu = $uiHash.Window.FindName("WUStartServiceMenu")
$uiHash.WURestartServiceMenu = $uiHash.Window.FindName("WURestartServiceMenu")
$uiHash.WindowsUpdateServiceMenu = $uiHash.Window.FindName("WindowsUpdateServiceMenu")
$uiHash.GenerateReportButton = $uiHash.Window.FindName("GenerateReportButton")
$uiHash.ReportComboBox = $uiHash.Window.FindName("ReportComboBox")
$uiHash.StartImage = $uiHash.Window.FindName("StartImage")
$uiHash.CancelImage = $uiHash.Window.FindName("CancelImage")
$uiHash.RunOptionComboBox = $uiHash.Window.FindName("RunOptionComboBox")
$uiHash.ClearErrorMenu = $uiHash.Window.FindName("ClearErrorMenu")
$uiHash.ViewErrorMenu = $uiHash.Window.FindName("ViewErrorMenu")
$uiHash.EntireLogMenu = $uiHash.Window.FindName("EntireLogMenu")
$uiHash.Last25LogMenu = $uiHash.Window.FindName("Last25LogMenu")
$uiHash.Last50LogMenu = $uiHash.Window.FindName("Last50LogMenu")
$uiHash.Last100LogMenu = $uiHash.Window.FindName("Last100LogMenu")
$uiHash.ResetDataMenu = $uiHash.Window.FindName("ResetDataMenu")
$uiHash.ResetAuthorizationMenu = $uiHash.Window.FindName("ResetAuthorizationMenu")
$uiHash.ServerListReportMenu = $uiHash.Window.FindName("ServerListReportMenu")
$uiHash.OfflineHostsMenu = $uiHash.Window.FindName("OfflineHostsMenu")
$uiHash.CompletedHostsMenu = $uiHash.Window.FindName("CompletedHostsMenu")
$uiHash.SEPMenu = $uiHash.Window.FindName("SEPMenu")
$uiHash.ESUMenu = $uiHash.Window.FindName("ESUMenu")
$uiHash.HostListMenu = $uiHash.Window.FindName("HostListMenu")
$uiHash.InstalledUpdatesMenu = $uiHash.Window.FindName("InstalledUpdatesMenu")
$uiHash.DetectNowMenu = $uiHash.Window.FindName("DetectNowMenu")
$uiHash.WindowsUpdateLogMenu = $uiHash.Window.FindName("WindowsUpdateLogMenu")
$uiHash.WUAUCLTMenu = $uiHash.Window.FindName("WUAUCLTMenu")
$uiHash.GUIInstalledUpdatesMenu = $uiHash.Window.FindName("GUIInstalledUpdatesMenu")
$uiHash.AddServerMenu = $uiHash.Window.FindName("AddServerMenu")
$uiHash.RemoveServerMenu = $uiHash.Window.FindName("RemoveServerMenu")
$uiHash.GetMenu = $uiHash.Window.FindName("GetMenu")
$uiHash.PrePatchMenu = $uiHash.Window.FindName("PrePatchMenu")
$uiHash.PostPatchMenu = $uiHash.Window.FindName("PostPatchMenu")
$uiHash.RebootMenu = $uiHash.Window.FindName("RebootMenu")
$uiHash.RemediateMenu = $uiHash.Window.FindName("RemediateMenu")
$uiHash.ChefPatchingMenu = $uiHash.Window.FindName("ChefPatchingMenu")
$uiHash.ChefRunMenu = $uiHash.Window.FindName("ChefRunMenu")
$uiHash.ChefRunNRMenu = $uiHash.Window.FindName("ChefRunNRMenu")
$uiHash.WUPatchingMenu = $uiHash.Window.FindName("WUPatchingMenu")
$uiHash.WindowsUpdateMenu = $uiHash.Window.FindName("WindowsUpdateMenu")
$uiHash.WindowsUpdateNRMenu = $uiHash.Window.FindName("WindowsUpdateNRMenu")
$uiHash.WindowsUpdateSRMenu = $uiHash.Window.FindName("WindowsUpdateSRMenu")
$uiHash.ChefRunAbortMenu = $uiHash.Window.FindName("ChefRunAbortMenu")
$uiHash.RunScriptMenu = $uiHash.Window.FindName("RunScriptMenu")
$uiHash.RunPSMenu = $uiHash.Window.FindName("RunPSMenu")
$uiHash.RunBatchMenu = $uiHash.Window.FindName("RunBatchMenu")
$uiHash.FailoverClusterMenu = $uiHash.Window.FindName("FailoverClusterMenu")
$uiHash.ChkNameNodeMenu = $uiHash.Window.FindName("ChkNameNodeMenu")
$uiHash.RunMoveResourceMenu = $uiHash.Window.FindName("RunMoveResourceMenu")
$uiHash.ChefLogsMenu = $uiHash.Window.FindName("ChefLogsMenu")
$uiHash.RobocopyMenu = $uiHash.Window.FindName("RobocopyMenu")
$uiHash.SEPverMenu = $uiHash.Window.FindName("SEPverMenu")
$uiHash.SpectreMenu = $uiHash.Window.FindName("SpectreMenu")
$uiHash.ExchangeMenu = $uiHash.Window.FindName("ExchangeMenu")
$uiHash.Run2k8ESUMenu = $uiHash.Window.FindName("Run2k8ESUMenu")
$uiHash.TSMMenu = $uiHash.Window.FindName("TSMMenu")
$uiHash.ListviewContextMenu = $uiHash.Window.FindName("ListViewContextMenu")
$uiHash.ExitMenu = $uiHash.Window.FindName("ExitMenu")
$uiHash.ClearInstalledUpdateMenu = $uiHash.Window.FindName("ClearInstalledUpdateMenu")
$uiHash.ClearServicesReportMenu = $uiHash.Window.FindName("ClearServicesReportMenu")
$uiHash.RunMenu = $uiHash.Window.FindName('RunMenu')
$uiHash.ClearAllMenu = $uiHash.Window.FindName("ClearAllMenu")
$uiHash.ClearServerListMenu = $uiHash.Window.FindName("ClearServerListMenu")
$uiHash.AboutMenu = $uiHash.Window.FindName("AboutMenu")
$uiHash.HelpFileMenu = $uiHash.Window.FindName("HelpFileMenu")
$uiHash.Listview = $uiHash.Window.FindName("Listview")
$uiHash.LoadFileButton = $uiHash.Window.FindName("LoadFileButton")
$uiHash.BrowseFileButton = $uiHash.Window.FindName("BrowseFileButton")
$uiHash.LoadADButton = $uiHash.Window.FindName("LoadADButton")
$uiHash.StatusTextBox = $uiHash.Window.FindName("StatusTextBox")
$uiHash.ProgressBar = $uiHash.Window.FindName("ProgressBar")
$uiHash.RunButton = $uiHash.Window.FindName("RunButton")
$uiHash.CancelButton = $uiHash.Window.FindName("CancelButton")
$uiHash.GridView = $uiHash.Window.FindName("GridView")
$uiHash.SymantecEndpointProtectionMenu = $uiHash.Window.FindName("SymantecEndpointProtectionMenu")
$uiHash.StartSEPMasterServiceMenu = $uiHash.Window.FindName("StartSEPMasterServiceMenu")
$uiHash.RepairSEPMenu = $uiHash.Window.FindName("RepairSEPMenu")
$uiHash.EndSepScanServiceMenu = $uiHash.Window.FindName("EndSepScanServiceMenu")
$uiHash.SEPHealthCheckMenu = $uiHash.Window.FindName("SEPHealthCheckMenu")
$uiHash.SEPVirusDefFoldersCheckMenu = $uiHash.Window.FindName("SEPVirusDefFoldersCheckMenu")
$uiHash.SEPOLDVirusDefDELETEMenu = $uiHash.Window.FindName("SEPOLDVirusDefDELETEMenu")
$uiHash.TrendMicroDSAHealthCheckMenu = $uiHash.Window.FindName("TrendMicroDSAHealthCheckMenu")
$uiHash.TMEDRProdAgentDeployentMenu = $uiHash.Window.FindName("TMEDRProdAgentDeployentMenu")
$uiHash.CheckPatchLogMenu = $uiHash.Window.FindName("CheckPatchLogMenu")
$uiHash.RunChefClientMenu = $uiHash.Window.FindName("RunChefClientMenu")
$uiHash.WSUSRegMenu = $uiHash.Window.FindName("WSUSRegMenu")
$uiHash.MsOfficeMenu = $uiHash.Window.FindName("MsOfficeMenu")
$uiHash.CBSlogsMenu = $uiHash.Window.FindName("CBSlogsMenu")
$uiHash.CheckPatchingRebFlagMenu = $uiHash.Window.FindName("CheckPatchingRebFlagMenu")
$uiHash.DelRebootFlagMenu = $uiHash.Window.FindName("DelRebootFlagMenu")
$uiHash.CheckRubyProcessMenu = $uiHash.Window.FindName("CheckRubyProcessMenu")
$uiHash.ENDRubyProcessMenu = $uiHash.Window.FindName("ENDRubyProcessMenu")
$uiHash.CheckCHEFenvVarMenu = $uiHash.Window.FindName("CheckCHEFenvVarMenu")
$uiHash.AddMissingCHEFEnvVarMenu = $uiHash.Window.FindName("AddMissingCHEFEnvVarMenu")
$uiHash.WinUpdateFixMenu = $uiHash.Window.FindName("WinUpdateFixMenu")
$uiHash.WinUpdateFixMultiServersMenu = $uiHash.Window.FindName("WinUpdateFixMultiServersMenu")
$uiHash.CHEFCommandsMenu = $uiHash.Window.FindName("CHEFCommandsMenu")
$uiHash.WUTroubleshootingMenu = $uiHash.Window.FindName("WUTroubleshootingMenu")
$uiHash.WUInformationMenu = $uiHash.Window.FindName("WUInformationMenu")
$uiHash.MicrosoftDotNetMenu = $uiHash.Window.FindName("MicrosoftDotNetMenu")
$uiHash.DotNetRuntimeOnOSMenu = $uiHash.Window.FindName("DotNetRuntimeOnOSMenu")
$uiHash.ASPVerMenu = $uiHash.Window.FindName("ASPVerMenu")
$uiHash.DotnetchkAllmenu = $uiHash.Window.FindName("DotnetchkAllmenu")
$uiHash.DotNetAddOnsMenu = $uiHash.Window.FindName("DotNetAddOnsMenu")
$uiHash.EDGECheckMenu = $uiHash.Window.FindName("EDGECheckMenu")
$uiHash.IISCheckMenu = $uiHash.Window.FindName("IISCheckMenu")
$uiHash.IISInstallMenu = $uiHash.Window.FindName("IISInstallMenu")
$uiHash.Win2012ESUYear1LicenseCheckMenu = $uiHash.Window.FindName("Win2012ESUYear1LicenseCheckMenu")
$uiHash.Win2012ESUYear1ProductKeyInstallMenu = $uiHash.Window.FindName("Win2012ESUYear1ProductKeyInstallMenu")
$uiHash.WindowsDefenderMenu = $uiHash.Window.FindName("WindowsDefenderMenu")
$uiHash.WinDefendCheckMenu = $uiHash.Window.FindName("WinDefendCheckMenu")
$uiHash.WinDefendOffboardMenu = $uiHash.Window.FindName("WinDefendOffboardMenu")
$uiHash.WinDefendUninstallMenu = $uiHash.Window.FindName("WinDefendUninstallMenu")
#endregion

#region Event Handlers

#Window Load Events
$uiHash.Window.Add_SourceInitialized({
    #Configure Options
    Write-Verbose 'Updating configuration based on options'
    Set-PoshPAIGOption 
    Write-Debug ("maxConcurrentJobs: {0}" -f $maxConcurrentJobs)
    Write-Debug ("MaxRebootJobs: {0}" -f $MaxRebootJobs)
    Write-Debug ("reportpath: {0}" -f $reportpath)
    
    #Define hashtable of settings
    $Script:SortHash = @{}
    
    #Sort event handler
    [System.Windows.RoutedEventHandler]$Global:ColumnSortHandler = {
        If ($_.OriginalSource -is [System.Windows.Controls.GridViewColumnHeader]) {
            Write-Verbose ("{0}" -f $_.Originalsource.getType().FullName)
            If ($_.OriginalSource -AND $_.OriginalSource.Role -ne 'Padding') {
                $Column = $_.Originalsource.Column.DisplayMemberBinding.Path.Path
                Write-Debug ("Sort: {0}" -f $Column)
                If ($SortHash[$Column] -eq 'Ascending') {
                    Write-Debug "Descending"
                    $SortHash[$Column]  = 'Descending'
                } Else {
                    Write-Debug "Ascending"
                    $SortHash[$Column]  = 'Ascending'
                }
                Write-Verbose ("Direction: {0}" -f $SortHash[$Column])
                $lastColumnsort = $Column
                Write-Verbose "Clearing sort descriptions"
                $uiHash.Listview.Items.SortDescriptions.clear()
                Write-Verbose ("Sorting {0} by {1}" -f $Column, $SortHash[$Column])
                $uiHash.Listview.Items.SortDescriptions.Add((New-Object System.ComponentModel.SortDescription $Column, $SortHash[$Column]))
                Write-Verbose "Refreshing View"
                $uiHash.Listview.Items.Refresh()   
            }             
        }
    }
    $uiHash.Listview.AddHandler([System.Windows.Controls.GridViewColumnHeader]::ClickEvent, $ColumnSortHandler)
    
    #Create and bind the observable collection to the GridView
    $Script:clientObservable = New-Object System.Collections.ObjectModel.ObservableCollection[object]    
    $uiHash.ListView.ItemsSource = $clientObservable
    $Global:Clients = $clientObservable | Select -Expand Computer
})    

#Window Close Events
$uiHash.Window.Add_Closed({
    #Halt job processing
    $jobCleanup.Flag = $False

    #Stop all runspaces
    $jobCleanup.PowerShell.Dispose()
    
    [gc]::Collect()
    [gc]::WaitForPendingFinalizers()    
})

#Cancel Button Event
$uiHash.CancelButton.Add_Click({
    $runspaceHash.runspacepool.Dispose()
    $uiHash.StatusTextBox.Foreground = "Black"
    $uiHash.StatusTextBox.Text = "Action cancelled" 
    [Float]$uiHash.ProgressBar.Value = 0
    $uiHash.RunButton.IsEnabled = $True
    $uiHash.StartImage.Source = "$pwd\Images\Start.jpg"
    $uiHash.CancelButton.IsEnabled = $False
    $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.jpg"    
         
})

#EntireUpdateLog Event
$uiHash.EntireLogMenu.Add_Click({
    If ($uiHash.Listview.Items.count -eq 1) {
        $selectedItem = $uiHash.Listview.SelectedItem
        [Float]$uiHash.ProgressBar.Maximum = $servers.count
        $uiHash.Listview.ItemsSource | ForEach {$_.Status = $Null}
        $uiHash.RunButton.IsEnabled = $False
        $uiHash.StartImage.Source = "$pwd\Images\Start_locked.jpg"
        $uiHash.CancelButton.IsEnabled = $True
        $uiHash.CancelImage.Source = "$pwd\Images\Stop.jpg"         
        $uiHash.StatusTextBox.Foreground = "Black"
        $uiHash.StatusTextBox.Text = "Retrieving Windows Update log from Server..."            
        $uiHash.StartTime = (Get-Date)        
        [Float]$uiHash.ProgressBar.Value = 0

        $uiHash.Listview.Items.EditItem($selectedItem)
        $selectedItem.Status = "Retrieving Update Log"
        $uiHash.Listview.Items.CommitEdit()
        $uiHash.Listview.Items.Refresh() 
        . .\Scripts\Get-UpdateLog.ps1
        Try {
            $log = Get-UpdateLog -Computername $selectedItem.computer 
            If ($log) {
                $log | Out-GridView -Title ("{0} Update Log" -f $selectedItem.computer)
                $selectedItem.Status = "Completed"
            }
        } Catch {
            $selectedItem.Status = $_.Exception.Message
        }
        $uiHash.ProgressBar.value++ 
        $End = New-Timespan $uihash.StartTime (Get-Date) 
        $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)
        $uiHash.RunButton.IsEnabled = $True
        $uiHash.StartImage.Source = "$pwd\Images\Start.jpg"
        $uiHash.CancelButton.IsEnabled = $False
        $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.jpg"
    }  
}) 

#Last100UpdateLog Event
$uiHash.Last100LogMenu.Add_Click({
    If ($uiHash.Listview.Items.count -eq 1) {
        $selectedItem = $uiHash.Listview.SelectedItem
        [Float]$uiHash.ProgressBar.Maximum = $servers.count
        $uiHash.Listview.ItemsSource | ForEach {$_.Status = $Null}
        $uiHash.RunButton.IsEnabled = $False
        $uiHash.StartImage.Source = "$pwd\Images\Start_locked.jpg"
        $uiHash.CancelButton.IsEnabled = $True
        $uiHash.CancelImage.Source = "$pwd\Images\Stop.jpg"         
        $uiHash.StatusTextBox.Foreground = "Black"
        $uiHash.StatusTextBox.Text = "Retrieving Windows Update log from Server..."            
        $uiHash.StartTime = (Get-Date)        
        [Float]$uiHash.ProgressBar.Value = 0

        $uiHash.Listview.Items.EditItem($selectedItem)
        $selectedItem.Status = "Retrieving Update Log"
        $uiHash.Listview.Items.CommitEdit()
        $uiHash.Listview.Items.Refresh() 
        . .\Scripts\Get-UpdateLog.ps1
        Try {
            $log = Get-UpdateLog -Last 100 -Computername $selectedItem.computer 
            If ($log) {
                $log | Out-GridView -Title ("{0} Update Log" -f $selectedItem.computer)
                $selectedItem.Status = "Completed"
            }
        } Catch {
            $selectedItem.Status = $_.Exception.Message
        }
        $uiHash.ProgressBar.value++ 
        $End = New-Timespan $uihash.StartTime (Get-Date) 
        $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)
        $uiHash.RunButton.IsEnabled = $True
        $uiHash.StartImage.Source = "$pwd\Images\Start.jpg"
        $uiHash.CancelButton.IsEnabled = $False
        $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.jpg"
    }       
})

#Last50UpdateLog Event
$uiHash.Last50LogMenu.Add_Click({
    If ($uiHash.Listview.Items.count -eq 1) {
        $selectedItem = $uiHash.Listview.SelectedItem
        [Float]$uiHash.ProgressBar.Maximum = $servers.count
        $uiHash.Listview.ItemsSource | ForEach {$_.Status = $Null}
        $uiHash.RunButton.IsEnabled = $False
        $uiHash.StartImage.Source = "$pwd\Images\Start_locked.jpg"
        $uiHash.CancelButton.IsEnabled = $True
        $uiHash.CancelImage.Source = "$pwd\Images\Stop.jpg"         
        $uiHash.StatusTextBox.Foreground = "Black"
        $uiHash.StatusTextBox.Text = "Retrieving Windows Update log from Server..."            
        $uiHash.StartTime = (Get-Date)        
        [Float]$uiHash.ProgressBar.Value = 0

        $uiHash.Listview.Items.EditItem($selectedItem)
        $selectedItem.Status = "Retrieving Update Log"
        $uiHash.Listview.Items.CommitEdit()
        $uiHash.Listview.Items.Refresh() 
        . .\Scripts\Get-UpdateLog.ps1
        Try {
            $log = Get-UpdateLog -Last 50 -Computername $selectedItem.computer 
            If ($log) {
                $log | Out-GridView -Title ("{0} Update Log" -f $selectedItem.computer)
                $selectedItem.Status = "Completed"
            }
        } Catch {
            $selectedItem.Status = $_.Exception.Message
        }
        $uiHash.ProgressBar.value++ 
        $End = New-Timespan $uihash.StartTime (Get-Date) 
        $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)
        $uiHash.RunButton.IsEnabled = $True
        $uiHash.StartImage.Source = "$pwd\Images\Start.jpg"
        $uiHash.CancelButton.IsEnabled = $False
        $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.jpg"
    }                  
})

#Last25UpdateLog Event
$uiHash.Last25LogMenu.Add_Click({
    If ($uiHash.Listview.Items.count -eq 1) {
        $selectedItem = $uiHash.Listview.SelectedItem
        [Float]$uiHash.ProgressBar.Maximum = $servers.count
        $uiHash.Listview.ItemsSource | ForEach {$_.Status = $Null}
        $uiHash.RunButton.IsEnabled = $False
        $uiHash.StartImage.Source = "$pwd\Images\Start_locked.jpg"
        $uiHash.CancelButton.IsEnabled = $True
        $uiHash.CancelImage.Source = "$pwd\Images\Stop.jpg"         
        $uiHash.StatusTextBox.Foreground = "Black"
        $uiHash.StatusTextBox.Text = "Retrieving Windows Update log from Server..."            
        $uiHash.StartTime = (Get-Date)        
        [Float]$uiHash.ProgressBar.Value = 0

        $uiHash.Listview.Items.EditItem($selectedItem)
        $selectedItem.Status = "Retrieving Update Log"
        $uiHash.Listview.Items.CommitEdit()
        $uiHash.Listview.Items.Refresh() 
        . .\Scripts\Get-UpdateLog.ps1
        Try {
            $log = Get-UpdateLog -Last 25 -Computername $selectedItem.computer 
            If ($log) {
                $log | Out-GridView -Title ("{0} Update Log" -f $selectedItem.computer)
                $selectedItem.Status = "Completed"
            }
        } Catch {
            $selectedItem.Status = $_.Exception.Message
        }
        $uiHash.ProgressBar.value++ 
        $End = New-Timespan $uihash.StartTime (Get-Date) 
        $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)
        $uiHash.RunButton.IsEnabled = $True
        $uiHash.StartImage.Source = "$pwd\Images\Start.jpg"
        $uiHash.CancelButton.IsEnabled = $False
        $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.jpg"             
    }    
})




#Offline server removal
$uiHash.OfflineHostsMenu.Add_Click({
    Write-Verbose "Removing any server that is shown as offline"
    $Offline = @($uiHash.Listview.Items | Where {$_.Status -eq "Offline"})
    $Offline | ForEach {
        Write-Verbose ("Removing {0}" -f $_.Computer)
        $clientObservable.Remove($_)
        }
    $uiHash.ProgressBar.Maximum = $uiHash.Listview.ItemsSource.count
})

#Completed patch server removal
$uiHash.CompletedHostsMenu.Add_Click({
    Write-Verbose "Removing any server that is shown as patch completed"
    $complete = @($uiHash.Listview.Items | Where {$_.Notes -eq "Patch Completed"})
    $complete | ForEach {
        Write-Verbose ("Removing {0}" -f $_.Computer)
        $clientObservable.Remove($_)
        }
    #$uiHash.ProgressBar.Maximum = $uiHash.Listview.ItemsSource.count
})

#Latest SEP version removal
$uiHash.SEPMenu.Add_Click({
    Write-Verbose "Remove server having SEP version 14.2.5323.2000 from the list."
    $latest = @($uiHash.Listview.Items | Where {$_.Notes -like '*14.2.5323.2000*'})
    $latest | ForEach {
        Write-Verbose ("Removing {0}" -f $_.Computer)
        $clientObservable.Remove($_)
        }
    #$uiHash.ProgressBar.Maximum = $uiHash.Listview.ItemsSource.count
})

#2k8 Licensed ESU removal
$uiHash.ESUMenu.Add_Click({
    Write-Verbose "Remove server having Licensed ESU from the list."
    $ESULic = @($uiHash.Listview.Items | Where {$_.Notes -eq 'Licensed'})
    $ESULic | ForEach {
        Write-Verbose ("Removing {0}" -f $_.Computer)
        $clientObservable.Remove($_)
        }
    #$uiHash.ProgressBar.Maximum = $uiHash.Listview.ItemsSource.count
})

#ResetAuthorization Event
$uiHash.ResetAuthorizationMenu.Add_Click({
    If ($uiHash.Listview.Items.count -gt 0) {
        $Servers = $uiHash.Listview.SelectedItems | Select -ExpandProperty Computer
        [Float]$uiHash.ProgressBar.Maximum = $servers.count
        $uiHash.Listview.ItemsSource | ForEach {$_.Status = $Null}
        $uiHash.RunButton.IsEnabled = $False
        $uiHash.StartImage.Source = "$pwd\Images\Start_locked.jpg"
        $uiHash.CancelButton.IsEnabled = $True
        $uiHash.CancelImage.Source = "$pwd\Images\Stop.jpg"
        $uiHash.StatusTextBox.Foreground = "Black"
        $uiHash.StatusTextBox.Text = "Forcing Reset Authorization on Servers"           
        $uiHash.StartTime = (Get-Date)        
        [Float]$uiHash.ProgressBar.Value = 0
        $scriptBlock = {
            Param (
                $Computer,
                $uiHash,
                $Path
            )               
            $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                $uiHash.Listview.Items.EditItem($Computer)
                $computer.Status = "Reset Authorization on Update Client"
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh() 
            })                
            Set-Location $Path
            $wmi = @{
                Computername = $computer.computer
                Class = "Win32_Process"
                Name = "Create"
                ErrorAction = "Stop"
                ArgumentList = "wuauclt /resetauthorization"
            }
            Try {
                If ((Invoke-WmiMethod @wmi).ReturnValue -eq 0) {
                    $result = $True
                } Else {
                    $result = $False
                }
            } Catch {
                $result = $False
                $returnMessage = $_.Exception.Message
            }
            $uiHash.ListView.Dispatcher.Invoke("Background",[action]{                    
                $uiHash.Listview.Items.EditItem($Computer)
                If ($result) {
                    $Computer.Status = "Completed"
                } Else {
                    $computer.Status = ("Issue Occurred: {0}" -f $returnMessage)
                } 
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh()
                $uiHash.ProgressBar.value++  
                    
                #Check to see if find job
                If ($uiHash.ProgressBar.value -eq $uiHash.ProgressBar.Maximum) { 
                    $End = New-Timespan $uihash.StartTime (Get-Date)                                             
                    $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)
                    $uiHash.RunButton.IsEnabled = $True
                    $uiHash.StartImage.Source = "$pwd\Images\Start.jpg"
                    $uiHash.CancelButton.IsEnabled = $False
                    $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.jpg"                    
                }
            })  
                
        }

        Write-Verbose ("Creating runspace pool and session states")
        $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
        $runspaceHash.runspacepool = [runspacefactory]::CreateRunspacePool(1, $maxConcurrentJobs, $sessionstate, $Host)
        $runspaceHash.runspacepool.Open()   

        ForEach ($Computer in $selectedItems) {
            $uiHash.Listview.Items.EditItem($Computer)
            $computer.Status = "Pending ResetAuthorization"
            $uiHash.Listview.Items.CommitEdit()
            $uiHash.Listview.Items.Refresh() 
            #Create the powershell instance and supply the scriptblock with the other parameters 
            $powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($computer).AddArgument($uiHash).AddArgument($Path)
           
            #Add the runspace into the powershell instance
            $powershell.RunspacePool = $runspaceHash.runspacepool
           
            #Create a temporary collection for each runspace
            $temp = "" | Select-Object PowerShell,Runspace,Computer
            $Temp.Computer = $Computer.computer
            $temp.PowerShell = $powershell
           
            #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
            $temp.Runspace = $powershell.BeginInvoke()
            Write-Verbose ("Adding {0} collection" -f $temp.Computer)
            $jobs.Add($temp) | Out-Null
        } 
    
    }    
})  
    
#DetectNow Event
$uiHash.DetectNowMenu.Add_Click({
    If ($uiHash.Listview.Items.count -gt 0) {
        $Servers = $uiHash.Listview.SelectedItems | Select -ExpandProperty Computer
        [Float]$uiHash.ProgressBar.Maximum = $servers.count
        $uiHash.Listview.ItemsSource | ForEach {$_.Status = $Null}
        $uiHash.RunButton.IsEnabled = $False
        $uiHash.StartImage.Source = "$pwd\Images\Start_locked.jpg"
        $uiHash.CancelButton.IsEnabled = $True
        $uiHash.CancelImage.Source = "$pwd\Images\Stop.jpg"
        $uiHash.StatusTextBox.Foreground = "Black"
        $uiHash.StatusTextBox.Text = "Forcing Re-Detection of Update Client on Servers"          
        $uiHash.StartTime = (Get-Date)        
        [Float]$uiHash.ProgressBar.Value = 0
        $scriptBlock = {
            Param (
                $Computer,
                $uiHash,
                $Path
            )               
            $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                $uiHash.Listview.Items.EditItem($Computer)
                $computer.Status = "Re-Detect on Update Client"
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh() 
            })                
            Set-Location $Path
            $wmi = @{
                Computername = $computer.computer
                Class = "Win32_Process"
                Name = "Create"
                ErrorAction = "Stop"
                ArgumentList = "wuauclt /detectnow"
            }
            Try {
                If ((Invoke-WmiMethod @wmi).ReturnValue -eq 0) {
                    $result = $True
                } Else {
                    $result = $False
                }
            } Catch {
                $result = $False
                $returnMessage = $_.Exception.Message
            }
            $uiHash.ListView.Dispatcher.Invoke("Background",[action]{                    
                $uiHash.Listview.Items.EditItem($Computer)
                If ($result) {
                    $Computer.Status = "Completed"
                } Else {
                    $computer.Status = ("Issue Occurred: {0}" -f $returnMessage)
                } 
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh()
                $uiHash.ProgressBar.value++  
                    
                #Check to see if find job
                If ($uiHash.ProgressBar.value -eq $uiHash.ProgressBar.Maximum) { 
                    $End = New-Timespan $uihash.StartTime (Get-Date)                                             
                    $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)
                    $uiHash.RunButton.IsEnabled = $True
                    $uiHash.StartImage.Source = "$pwd\Images\Start.jpg"
                    $uiHash.CancelButton.IsEnabled = $False
                    $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.jpg"                    
                }
            })  
                
        }

        Write-Verbose ("Creating runspace pool and session states")
        $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
        $runspaceHash.runspacepool = [runspacefactory]::CreateRunspacePool(1, $maxConcurrentJobs, $sessionstate, $Host)
        $runspaceHash.runspacepool.Open()   

        ForEach ($Computer in $selectedItems) {
            $uiHash.Listview.Items.EditItem($Computer)
            $computer.Status = "Pending DetectNow"
            $uiHash.Listview.Items.CommitEdit()
            $uiHash.Listview.Items.Refresh() 
            #Create the powershell instance and supply the scriptblock with the other parameters 
            $powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($computer).AddArgument($uiHash).AddArgument($Path)
           
            #Add the runspace into the powershell instance
            $powershell.RunspacePool = $runspaceHash.runspacepool
           
            #Create a temporary collection for each runspace
            $temp = "" | Select-Object PowerShell,Runspace,Computer
            $Temp.Computer = $Computer.computer
            $temp.PowerShell = $powershell
           
            #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
            $temp.Runspace = $powershell.BeginInvoke()
            Write-Verbose ("Adding {0} collection" -f $temp.Computer)
            $jobs.Add($temp) | Out-Null
        } 
    
    }    
})        

#Rightclick Event
$uiHash.Listview.Add_MouseRightButtonUp({
    $listcount = $uiHash.Listview.SelectedItems.count
    Write-Debug "$($This.SelectedItem.Row.Computer)"
    If ($uiHash.Listview.SelectedItems.count -eq 0) {
        $uiHash.RemoveServerMenu.IsEnabled = $False
        $uiHash.GetMenu.IsEnabled = $False
		$uiHash.ChefRunAbortMenu.IsEnabled = $False
		$uiHash.RunScriptMenu.IsEnabled = $False
        $uiHash.RebootMenu.IsEnabled = $False
        $uihash.RobocopyMenu.isenabled = $false
        $uiHash.RemediateMenu.IsEnabled = $False
        $uiHash.InstalledUpdatesMenu.IsEnabled = $False
        $uiHash.WindowsUpdateLogMenu.IsEnabled = $False
        $uiHash.WindowsUpdateServiceMenu.IsEnabled = $False
        $uiHash.WUAUCLTMenu.IsEnabled = $False
		$uiHash.WinUpdateFixMenu.IsEnabled = $False
		$uihash.WindowsDefenderMenu.isenabled = $False
        } ElseIf ($uiHash.Listview.SelectedItems.count -eq 1) {
        $uiHash.RemoveServerMenu.IsEnabled = $True
        $uiHash.GetMenu.IsEnabled = $True
		$uiHash.ChefRunAbortMenu.IsEnabled = $true
		$uiHash.RunScriptMenu.IsEnabled = $true
        $uiHash.RebootMenu.IsEnabled = $True
        $uihash.RobocopyMenu.isenabled = $true
        $uiHash.RemediateMenu.IsEnabled = $True
        $uiHash.InstalledUpdatesMenu.IsEnabled = $True
        $uiHash.WindowsUpdateLogMenu.IsEnabled = $True
        $uiHash.WindowsUpdateServiceMenu.IsEnabled = $True
        $uiHash.WUAUCLTMenu.IsEnabled = $True
		$uiHash.WinUpdateFixMenu.IsEnabled = $True
		$uihash.WindowsDefenderMenu.isenabled = $True
        $uiHash.StatusTextBox.Text = "Total selected server(s): $listcount"
        } Else {
        $uiHash.RemoveServerMenu.IsEnabled = $True
        $uiHash.GetMenu.IsEnabled = $True
		$uiHash.ChefRunAbortMenu.IsEnabled = $true
		$uiHash.RunScriptMenu.IsEnabled = $true
        $uiHash.RebootMenu.IsEnabled = $True
        $uihash.RobocopyMenu.isenabled = $True
        $uiHash.RemediateMenu.IsEnabled = $True
        $uiHash.InstalledUpdatesMenu.IsEnabled = $True
        $uiHash.WindowsUpdateLogMenu.IsEnabled = $False
        $uiHash.WUAUCLTMenu.IsEnabled = $True
		$uiHash.WinUpdateFixMenu.IsEnabled = $False
		$uihash.WindowsDefenderMenu.IsEnabled = $True
        $uiHash.StatusTextBox.Text = "Total selected server(s): $listcount"   
    }    
})

#ListView drop file Event
$uiHash.Listview.add_Drop({
    $content = Get-Content $_.Data.GetFileDropList()
    $content | ForEach {
        $clientObservable.Add((
            New-Object PSObject -Property @{
                Computer = $_
                Audited = 0 -as [int]
                Installed = 0 -as [int]
                InstallErrors = 0 -as [int]
                Services = 0 -as [int]
                Status = $Null
                Notes = $Null
            }
        ))      
    }
    Show-DebugState
})

#FindFile Button
$uiHash.BrowseFileButton.Add_Click({
    $File = Open-FileDialog
    $hostCount = 0 -as [int]
    If (-Not ([system.string]::IsNullOrEmpty($File))) {
        Get-Content $File | Where {$_ -ne ""} | ForEach {
            $clientObservable.Add((
                New-Object PSObject -Property @{
                    Computer = $_
                    Audited = 0 -as [int]
                    Installed = 0 -as [int]
                    InstallErrors = 0 -as [int]
                    Services = 0 -as [int]
                    Status = $Null
                    Notes = $Null
                }
            )) 
            $hostCount = $hostCount + 1 
        }
        $res += $hostCount
        $uiHash.StatusTextBox.Foreground = "Black"
        $uiHash.StatusTextBox.Text = "Total added server(s): $res"
        Show-DebugState     
    }        
})

#LoadADButton Events    
$uiHash.LoadADButton.Add_Click({
    $domain = Open-DomainDialog
    $uiHash.StatusTextBox.Foreground = "Black"
    $uiHash.StatusTextBox.Text = "Querying Active Directory for Computers..."
    $Searcher = [adsisearcher]""  
    $Searcher.SearchRoot= [adsi]"LDAP://$domain"
    $Searcher.Filter = ("(&(objectCategory=computer)(OperatingSystem=*server*))")
    $Searcher.PropertiesToLoad.Add('name') | Out-Null
    Write-Verbose "Checking for exempt list"
    If (Test-Path Exempt.txt) {
        Write-Verbose "Collecting systems from exempt list"
        [string[]]$exempt = Get-Content Exempt.txt
    }
    $Results = $Searcher.FindAll()
    foreach ($result in $Results) {
        [string]$computer = $result.Properties.name
        If ($Exempt -notcontains $computer -AND -NOT $ComputerCache.contains($Computer)) {
            [void]$ComputerCache.Add($Computer)
            $clientObservable.Add((
                New-Object PSObject -Property @{
                    Computer = $computer
                    Audited = 0 -as [int]
                    Installed = 0 -as [int]
                    InstallErrors = 0 -as [int]
                    Services = 0 -as [int]
                    Status = $Null
                    Notes = $Null
                }
            ))     
        } Else {
            Write-Verbose "Excluding $computer"
        }
    }
    $uiHash.ProgressBar.Maximum = $uiHash.Listview.ItemsSource.count   
    $Global:clients = $clientObservable | Select -Expand Computer
    Show-DebugState                      
})

#RunButton Events    
$uiHash.RunButton.add_Click({
    Start-RunJob      
})

#region Client WSUS Service Action
#Stop Service
$uiHash.WUStopServiceMenu.Add_Click({
    If ($uiHash.Listview.Items.count -gt 0) {
        $Servers = $uiHash.Listview.SelectedItems | Select -ExpandProperty Computer
        [Float]$uiHash.ProgressBar.Maximum = $servers.count
        $uiHash.Listview.ItemsSource | ForEach {$_.Status = $Null}
        $uiHash.RunButton.IsEnabled = $False
        $uiHash.StartImage.Source = "$pwd\Images\Start_locked.jpg"
        $uiHash.CancelButton.IsEnabled = $True
        $uiHash.CancelImage.Source = "$pwd\Images\Stop.jpg"
        $uiHash.StatusTextBox.Foreground = "Black"
        $uiHash.StatusTextBox.Text = "Stopping WSUS Client service on selected servers"            
        $uiHash.StartTime = (Get-Date)        
        [Float]$uiHash.ProgressBar.Value = 0
        $scriptBlock = {
            Param (
                $Computer,
                $uiHash,
                $Path
            )               
            $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                $uiHash.Listview.Items.EditItem($Computer)
                $computer.Status = "Stopping Update Client Service"
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh() 
            })                
            Set-Location $Path
            Try {
                $updateClient = Get-Service -ComputerName $computer.computer -Name wuauserv -ErrorAction Stop
                Stop-Service -inputObject $updateClient -ErrorAction Stop
                $result = $True
            } Catch {
                $updateClient = $_.Exception.Message
            }
            $uiHash.ListView.Dispatcher.Invoke("Background",[action]{                    
                $uiHash.Listview.Items.EditItem($Computer)
                If ($result) {
                    $Computer.Status = "Service Stopped"
                } Else {
                    $computer.Status = ("Issue Occurred: {0}" -f $updateClient)
                } 
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh()
            })
            $uiHash.ProgressBar.Dispatcher.Invoke("Background",[action]{   
                $uiHash.ProgressBar.value++  
            })
            $uiHash.Window.Dispatcher.Invoke("Background",[action]{                       
                #Check to see if find job
                If ($uiHash.ProgressBar.value -eq $uiHash.ProgressBar.Maximum) { 
                    $End = New-Timespan $uihash.StartTime (Get-Date)                                             
                    $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)
                    $uiHash.RunButton.IsEnabled = $True
                    $uiHash.StartImage.Source = "$pwd\Images\Start.jpg"
                    $uiHash.CancelButton.IsEnabled = $False
                    $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.jpg"                    
                }
            })  
                
        }

        Write-Verbose ("Creating runspace pool and session states")
        $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
        $runspaceHash.runspacepool = [runspacefactory]::CreateRunspacePool(1, $maxConcurrentJobs, $sessionstate, $Host)
        $runspaceHash.runspacepool.Open()   

        ForEach ($Computer in $selectedItems) {
            $uiHash.Listview.Items.EditItem($Computer)
            $computer.Status = "Pending Stop Service"
            $uiHash.Listview.Items.CommitEdit()
            $uiHash.Listview.Items.Refresh() 
            #Create the powershell instance and supply the scriptblock with the other parameters 
            $powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($computer).AddArgument($uiHash).AddArgument($Path)
           
            #Add the runspace into the powershell instance
            $powershell.RunspacePool = $runspaceHash.runspacepool
           
            #Create a temporary collection for each runspace
            $temp = "" | Select-Object PowerShell,Runspace,Computer
            $Temp.Computer = $Computer.computer
            $temp.PowerShell = $powershell
           
            #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
            $temp.Runspace = $powershell.BeginInvoke()
            Write-Verbose ("Adding {0} collection" -f $temp.Computer)
            $jobs.Add($temp) | Out-Null
        }     
    }    
})

#Start Service
$uiHash.WUStartServiceMenu.Add_Click({
    If ($uiHash.Listview.Items.count -gt 0) {
        $Servers = $uiHash.Listview.SelectedItems | Select -ExpandProperty Computer
        [Float]$uiHash.ProgressBar.Maximum = $servers.count
        $uiHash.Listview.ItemsSource | ForEach {$_.Status = $Null}
        $uiHash.RunButton.IsEnabled = $False
        $uiHash.StartImage.Source = "$pwd\Images\Start_locked.jpg"
        $uiHash.CancelButton.IsEnabled = $True
        $uiHash.CancelImage.Source = "$pwd\Images\Stop.jpg"
        $uiHash.StatusTextBox.Foreground = "Black"
        $uiHash.StatusTextBox.Text = "Starting WSUS Client service on selected servers"            
        $uiHash.StartTime = (Get-Date)        
        [Float]$uiHash.ProgressBar.Value = 0
        $scriptBlock = {
            Param (
                $Computer,
                $uiHash,
                $Path
            )               
            $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                $uiHash.Listview.Items.EditItem($Computer)
                $computer.Status = "Starting Update Client Service"
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh() 
            })                
            Set-Location $Path
            Try {
                $updateClient = Get-Service -ComputerName $computer.computer -Name wuauserv -ErrorAction Stop
                Start-Service -inputObject $updateClient -ErrorAction Stop
                $result = $True
            } Catch {
                $updateClient = $_.Exception.Message
            }
            $uiHash.ListView.Dispatcher.Invoke("Background",[action]{                    
                $uiHash.Listview.Items.EditItem($Computer)
                If ($result) {
                    $Computer.Status = "Service Started"
                } Else {
                    $computer.Status = ("Issue Occurred: {0}" -f $updateClient)
                } 
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh()
            })
            $uiHash.ProgressBar.Dispatcher.Invoke("Background",[action]{   
                $uiHash.ProgressBar.value++  
            })
            $uiHash.Window.Dispatcher.Invoke("Background",[action]{           
                #Check to see if find job
                If ($uiHash.ProgressBar.value -eq $uiHash.ProgressBar.Maximum) { 
                    $End = New-Timespan $uihash.StartTime (Get-Date)                                             
                    $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)
                    $uiHash.RunButton.IsEnabled = $True
                    $uiHash.StartImage.Source = "$pwd\Images\Start.jpg"
                    $uiHash.CancelButton.IsEnabled = $False
                    $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.jpg"                    
                }
            })  
                
        }

        Write-Verbose ("Creating runspace pool and session states")
        $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
        $runspaceHash.runspacepool = [runspacefactory]::CreateRunspacePool(1, $maxConcurrentJobs, $sessionstate, $Host)
        $runspaceHash.runspacepool.Open()   

        ForEach ($Computer in $selectedItems) {
            $uiHash.Listview.Items.EditItem($Computer)
            $computer.Status = "Pending Start Service"
            $uiHash.Listview.Items.CommitEdit()
            $uiHash.Listview.Items.Refresh()
            #Create the powershell instance and supply the scriptblock with the other parameters 
            $powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($computer).AddArgument($uiHash).AddArgument($Path)
           
            #Add the runspace into the powershell instance
            $powershell.RunspacePool = $runspaceHash.runspacepool
           
            #Create a temporary collection for each runspace
            $temp = "" | Select-Object PowerShell,Runspace,Computer
            $Temp.Computer = $Computer.computer
            $temp.PowerShell = $powershell
           
            #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
            $temp.Runspace = $powershell.BeginInvoke()
            Write-Verbose ("Adding {0} collection" -f $temp.Computer)
            $jobs.Add($temp) | Out-Null
        } 
    
    }    
})

#Restart Service
$uiHash.WURestartServiceMenu.Add_Click({
    If ($uiHash.Listview.Items.count -gt 0) {
        $Servers = $uiHash.Listview.SelectedItems | Select -ExpandProperty Computer
        [Float]$uiHash.ProgressBar.Maximum = $servers.count
        $uiHash.Listview.ItemsSource | ForEach {$_.Status = $Null}
        $uiHash.RunButton.IsEnabled = $False
        $uiHash.StartImage.Source = "$pwd\Images\Start_locked.jpg"
        $uiHash.CancelButton.IsEnabled = $True
        $uiHash.CancelImage.Source = "$pwd\Images\Stop.jpg"
        $uiHash.StatusTextBox.Foreground = "Black"
        $uiHash.StatusTextBox.Text = "Restarting WSUS Client service on selected servers"            
        $uiHash.StartTime = (Get-Date)        
        [Float]$uiHash.ProgressBar.Value = 0
        $scriptBlock = {
            Param (
                $Computer,
                $uiHash,
                $Path
            )               
            $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                $uiHash.Listview.Items.EditItem($Computer)
                $computer.Status = "Restarting Update Client Service"
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh() 
            })                
            Set-Location $Path
            Try {
                $updateClient = Get-Service -ComputerName $computer.computer -Name wuauserv -ErrorAction Stop
                Restart-Service -inputObject $updateClient -ErrorAction Stop
                $result = $True
            } Catch {
                $updateClient = $_.Exception.Message
            }
            $uiHash.ListView.Dispatcher.Invoke("Background",[action]{                    
                $uiHash.Listview.Items.EditItem($Computer)
                If ($result) {
                    $Computer.Status = "Service Restarted"
                } Else {
                    $computer.Status = ("Issue Occurred: {0}" -f $updateClient)
                } 
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh()
            })
            $uiHash.ProgressBar.Dispatcher.Invoke("Background",[action]{   
                $uiHash.ProgressBar.value++  
            })
            $uiHash.Window.Dispatcher.Invoke("Background",[action]{           
                #Check to see if find job
                If ($uiHash.ProgressBar.value -eq $uiHash.ProgressBar.Maximum) { 
                    $End = New-Timespan $uihash.StartTime (Get-Date)                                             
                    $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)
                    $uiHash.RunButton.IsEnabled = $True
                    $uiHash.StartImage.Source = "$pwd\Images\Start.jpg"
                    $uiHash.CancelButton.IsEnabled = $False
                    $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.jpg"                    
                }
            })  
                
        }

        Write-Verbose ("Creating runspace pool and session states")
        $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
        $runspaceHash.runspacepool = [runspacefactory]::CreateRunspacePool(1, $maxConcurrentJobs, $sessionstate, $Host)
        $runspaceHash.runspacepool.Open()   

        ForEach ($Computer in $selectedItems) {
            $uiHash.Listview.Items.EditItem($Computer)
            $computer.Status = "Pending Restart Service"
            $uiHash.Listview.Items.CommitEdit()
            $uiHash.Listview.Items.Refresh()
            #Create the powershell instance and supply the scriptblock with the other parameters 
            $powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($computer).AddArgument($uiHash).AddArgument($Path)
           
            #Add the runspace into the powershell instance
            $powershell.RunspacePool = $runspaceHash.runspacepool
           
            #Create a temporary collection for each runspace
            $temp = "" | Select-Object PowerShell,Runspace,Computer
            $Temp.Computer = $Computer.computer
            $temp.PowerShell = $powershell
           
            #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
            $temp.Runspace = $powershell.BeginInvoke()
            Write-Verbose ("Adding {0} collection" -f $temp.Computer)
            $jobs.Add($temp) | Out-Null
        } 
    
    }    
})
#endregion

#View Installed Update Event
$uiHash.GUIInstalledUpdatesMenu.Add_Click({
    If ($uiHash.Listview.Items.count -gt 0) {
        $selectedItems = $uiHash.Listview.SelectedItems
        [Float]$uiHash.ProgressBar.Maximum = $servers.count
        $uiHash.Listview.ItemsSource | ForEach {$_.Status = $Null}
        $uiHash.RunButton.IsEnabled = $False
        $uiHash.StartImage.Source = "$pwd\Images\Start_locked.jpg"
        $uiHash.CancelButton.IsEnabled = $True
        $uiHash.CancelImage.Source = "$pwd\Images\Stop.jpg"        
        $uiHash.StatusTextBox.Foreground = "Black"
        $uiHash.StatusTextBox.Text = "Gathering all installed patches on Servers"            
        $uiHash.StartTime = (Get-Date)        
        [Float]$uiHash.ProgressBar.Value = 0             

        $scriptBlock = {
            Param (
                $Computer,
                $uiHash,
                $Path,
                $installedUpdates
            )               
            $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                $uiHash.Listview.Items.EditItem($Computer)
                $computer.Status = "Querying Installed Updates"
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh() 
            })                
            Set-Location $Path
            Try {
                $updates = Get-HotFix -ComputerName $computer.computer -ErrorAction Stop | Where {$_.Description -ne ""}
                If ($updates) {
                    $installedUpdates.AddRange($updates) | Out-Null

                    If ($installedUpdates.count -gt 0) {
                        $installedUpdates #| Out-GridView -Title 'Installed Updates Report'
                    } Else {
                        $uiHash.StatusTextBox.Foreground = "Red"
                        $uiHash.StatusTextBox.Text = "No report to create!"         
                    } 

                }
            } Catch {
                $result = $_.exception.Message
            }
            $uiHash.ListView.Dispatcher.Invoke("Background",[action]{                    
                $uiHash.Listview.Items.EditItem($Computer)
                If ($result) {
                    $computer.Status = $result
                } Else {
                    $computer.Status = "Completed"
                }
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh()
            })
            $uiHash.ProgressBar.Dispatcher.Invoke("Background",[action]{   
                $uiHash.ProgressBar.value++  
            })
            $uiHash.Window.Dispatcher.Invoke("Background",[action]{           
                #Check to see if find job
                If ($uiHash.ProgressBar.value -eq $uiHash.ProgressBar.Maximum) { 
                    $End = New-Timespan $uihash.StartTime (Get-Date)                                             
                    $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)
                    $uiHash.RunButton.IsEnabled = $True
                    $uiHash.StartImage.Source = "$pwd\Images\Start.jpg"
                    $uiHash.CancelButton.IsEnabled = $False
                    $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.jpg"                 
                }
            })  
                
        }
        Write-Verbose ("Creating runspace pool and session states")
        $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
        $runspaceHash.runspacepool = [runspacefactory]::CreateRunspacePool(1, $maxConcurrentJobs, $sessionstate, $Host)
        $runspaceHash.runspacepool.Open()   

        ForEach ($Computer in $selectedItems) {
            #Create the powershell instance and supply the scriptblock with the other parameters 
            $powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($computer).AddArgument($uiHash).AddArgument($Path).AddArgument($installedUpdates)
           
            #Add the runspace into the powershell instance
            $powershell.RunspacePool = $runspaceHash.runspacepool
           
            #Create a temporary collection for each runspace
            $temp = "" | Select-Object PowerShell,Runspace,Computer
            $Temp.Computer = $Computer.computer
            $temp.PowerShell = $powershell
           
            #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
            $temp.Runspace = $powershell.BeginInvoke()
            Write-Verbose ("Adding {0} collection" -f $temp.Computer)
            $jobs.Add($temp) | Out-Null                
        }
    }
})

#ClearAuditReportMenu Events    
$uiHash.ClearAuditReportMenu.Add_Click({
    $global:updateAudit.Clear()
    $uiHash.StatusTextBox.Foreground = "Black"
    $uiHash.StatusTextBox.Text = "Audit Report Cleared!"  
})

#ClearPrepReportMenu Events    
$uiHash.ClearPrepReportMenu.Add_Click({
    $global:prep.Clear()
    $uiHash.StatusTextBox.Foreground = "Black"
    $uiHash.StatusTextBox.Text = "Prep Report Cleared!"  
})

#ClearESUReportMenu Events    
$uiHash.ClearESUReportMenu.Add_Click({
    $global:Esureport.Clear()
    $uiHash.StatusTextBox.Foreground = "Black"
    $uiHash.StatusTextBox.Text = "ESU Report Cleared!"  
})

#ClearPostPatchReportMenu Events    
$uiHash.ClearPostPatchReportMenu.Add_Click({
    $global:postpatch.Clear()
    $uiHash.StatusTextBox.Foreground = "Black"
    $uiHash.StatusTextBox.Text = "Post-Patch Report Cleared!"  
})

#ClearInstallReportMenu Events    
$uiHash.ClearInstallReportMenu.Add_Click({
    Remove-Variable InstallPatchReport -scope Global -force -ea 'silentlycontinue'
    $uiHash.StatusTextBox.Foreground = "Black"
    $uiHash.StatusTextBox.Text = "Install Report Cleared!"  
})

#ClearInstalledUpdateMenu
$uiHash.ClearInstalledUpdateMenu.Add_Click({
    Remove-Variable InstalledPatches -scope Global -force -ea 'silentlycontinue'
    $uiHash.StatusTextBox.Foreground = "Black"
    $uiHash.StatusTextBox.Text = "Installed Updates Report Cleared!"    
})

#ClearServicesReportMenu
$uiHash.ClearServicesReportMenu.Add_Click({
    $global:servicesAudit.clear()
    $uiHash.StatusTextBox.Foreground = "Black"
    $uiHash.StatusTextBox.Text = "Stopped Services Report Cleared!"    
})
    
#ClearServerListMenu Events    
$uiHash.ClearServerListMenu.Add_Click({
    $clientObservable.Clear()
    $ComputerCache.Clear()
    $uiHash.StatusTextBox.Foreground = "Black"
    $uiHash.StatusTextBox.Text = "Server List Cleared!"  
})    

#AboutMenu Event
$uiHash.AboutMenu.Add_Click({
    Open-PoshPAIGAbout
})

#RunPrep Event
$uiHash.PrePatchMenu.Add_Click({
RunPrePatch
})

#RunPostPatch Event
$uiHash.PostPatchMenu.Add_Click({
RunPostPatch
})

#RunChefPatch Event
$uiHash.ChefRunMenu.Add_Click({
$invoke.Clear()
$invoke.Add("Chefinvoke")
Show-ChefRunWarning "Chef Process"
})

#RunChefPatch Event
$uiHash.ChefRunNRMenu.Add_Click({
$invoke.Clear()
$invoke.Add("ChefinvokeNR")
Show-ChefRunWarning "Suppressed Chef Process"
})

#RunChefPatch Event
$uiHash.ChefRunAbortMenu.Add_Click({
$invoke.Clear()
$invoke.Add("ChefinvokeAbort")
Show-ChefRunWarning "Abort Chef Process"
})

#RunWindowsUpdate Event
$uiHash.WindowsUpdateMenu.Add_Click({
$WUinvoke.clear()
$WUinvoke.add("WU")
Show-ChefRunWarning "Windows Update"
})

#RunWindowsUpdateNR Event
$uiHash.WindowsUpdateNRMenu.Add_Click({
$WUinvoke.clear()
$WUinvoke.add("NR")
Show-ChefRunWarning "Windows Update"
})

#RunWindowsUpdateSR Event
$uiHash.WindowsUpdateSRMenu.Add_Click({
$WUinvoke.clear()
$WUinvoke.add("SR")
Show-ChefRunWarning "Windows Update"
})

#RunPSScript Event
$uiHash.RunPSMenu.Add_Click({
showInputbox "PS"
})

#RunBATScript Event
$uiHash.RunBatchMenu.Add_Click({
showInputbox "BAT"
})

#SystemRestart Event
$uiHash.RebootMenu.Add_Click({
#System-reboot
Show-Warning "ServerReboot"
})

#ChefLogs Event
$uiHash.ChefLogsMenu.Add_Click({
Check-ChefLogs
})

#ChkNodeName Event
$uiHash.ChkNameNodeMenu.Add_Click({
$invkClstr.Clear()
$invkClstr.Add("chk")
Chk-ClstrNameNode
})

#ChkNodeName Event
$uiHash.RunMoveResourceMenu.Add_Click({
$invkClstr.Clear()
$invkClstr.Add("move")
Chk-ClstrNameNode
})

#RoboCopy Event
$uiHash.RobocopyMenu.Add_Click({
RobocopyPrompt
})

#Check SEP version Event
$uiHash.SEPverMenu.Add_Click({
Check-SEPver
})

#CheckSpectre Event
$uiHash.SpectreMenu.Add_Click({
Check-Spectre
})

#CheckExchangeServices Event
$uiHash.ExchangeMenu.Add_Click({
Check-Exchange
})

#CheckTSM Event
$uiHash.TSMMenu.Add_Click({
Check-TSM
})


#Run2k8ESU Event
#$uiHash.Run2k8ESUMenu.Add_Click({
#Show-Warning "2k8ESULicense"
#})

#Options Menu
$uiHash.OptionMenu.Add_Click({
    #Launch options window
    Write-Verbose "Launching Options Menu"
    .\Options.ps1
    #Process Updates Options
    Set-PoshPAIGOption    
})

#Select All
$uiHash.SelectAllMenu.Add_Click({
    $uiHash.Listview.SelectAll()
    Get-SelectedItemCount
})

#HelpFileMenu Event
$uiHash.HelpFileMenu.Add_Click({
    Open-PoshPAIGHelp
})

#KeyDown Event
$uiHash.Window.Add_Keydown({ 
    $key = $_.Key 
    If ([System.Windows.Input.Keyboard]::IsKeydown("RightCtrl") -OR [System.Windows.Input.Keyboard]::IsKeydown("LeftCtrl")) {
        Switch ($Key) {
        "E" {$This.Close()}
        "O" {
            .\Options.ps1
            #Process Updates Options
            Set-PoshPAIGOption
        }
        "S" {Add-Server}
        "D" {Remove-Server}
        Default {$Null}
        }Get-SelectedItemCount
    } ElseIf ([System.Windows.Input.Keyboard]::IsKeydown("LeftShift") -OR [System.Windows.Input.Keyboard]::IsKeydown("RightShift")) {
        Switch ($Key) {
            "RETURN" {Write-Host "Hit Shift+Return"}
        }
    }   
})

#Key Up Event
$uiHash.Window.Add_KeyUp({
    $Global:Test = $_
    Write-Debug ("Key Pressed: {0}" -f $_.Key)
    Switch ($_.Key) {
        "F1" {Open-PoshPAIGHelp}
        "F5" {Start-RunJob}
        "F8" {Start-Report}
        Default {$Null}
    }

    $key = $_.Key
    If ([System.Windows.Input.Keyboard]::IsKeyUp("RightCtrl") -OR [System.Windows.Input.Keyboard]::IsKeyUp("LeftCtrl")) {
        Switch ($Key) {
        "A" {$uiHash.Listview.SelectAll()
            Get-SelectedItemCount}
        Default {$Null}
        }
    } ElseIf ([System.Windows.Input.Keyboard]::IsKeyUp("LeftShift") -OR [System.Windows.Input.Keyboard]::IsKeyUp("RightShift")) {
        Switch ($Key) {
            "RETURN" {Write-Host "Hit Shift+Return"}
        }
    }
})

#AddServer Menu
$uiHash.AddServerMenu.Add_Click({
    Add-Server   
})

#RemoveServer Menu
$uiHash.RemoveServerMenu.Add_Click({
    Remove-Server 
})  

#Run Menu
$uiHash.RunMenu.Add_Click({
    Start-RunJob
})      
      
#Report Menu
$uiHash.GenerateReportMenu.Add_Click({
    Start-Report
    $computer.audited = 0
})       
      
#Exit Menu
$uiHash.ExitMenu.Add_Click({
    $uiHash.Window.Close()
})

#ClearAll Menu
$uiHash.ClearAllMenu.Add_Click({
    $clientObservable.Clear()
    $ComputerCache.Clear()
    $content = $Null
    [Float]$uiHash.ProgressBar.value = 0
    $uiHash.StatusTextBox.Foreground = "Black"
    $Global:updateAudit.Clear()
    $Global:prep.Clear()
    $Global:Esureport.Clear()
    $Global:postpatch.clear()
    $Global:servicesAudit.clear()
    $Global:installAudit.clear()
    $Global:installedUpdates.clear()
    $uiHash.StatusTextBox.Text = "Waiting for action..."    
})



#Save Server List
$uiHash.ServerListReportMenu.Add_Click({
    If ($uiHash.Listview.Items.count -gt 0) {
        $uiHash.StatusTextBox.Foreground = "Black"
        $savedreport = Join-Path (Join-Path $home Desktop) "serverlist.csv"
        $uiHash.Listview.ItemsSource | Export-Csv -NoTypeInformation $savedreport
        $uiHash.StatusTextBox.Text = "Report saved to $savedreport"        
    } Else {
        $uiHash.StatusTextBox.Foreground = "Red"
        $uiHash.StatusTextBox.Text = "No report to create!"         
    }         
})
     
#HostListMenu
$uiHash.HostListMenu.Add_Click({
    If ($uiHash.Listview.Items.count -gt 0) { 
        $uiHash.StatusTextBox.Foreground = "Black"
        $savedreport = Join-Path $reportpath "hosts.txt"
        $uiHash.Listview.DataContext | Select -Expand Computer | Out-File $savedreport
        $uiHash.StatusTextBox.Text = "Report saved to $savedreport"
        } Else {
        $uiHash.StatusTextBox.Foreground = "Red"
        $uiHash.StatusTextBox.Text = "No report to create!"         
    }         
})     

#Report Generation
$uiHash.GenerateReportButton.Add_Click({
    Start-Report
})

#Clear Error log
$uiHash.ClearErrorMenu.Add_Click({
    Write-Verbose "Clearing error log"
    $Error.Clear()
})

#View Error Event
$uiHash.ViewErrorMenu.Add_Click({
    Get-Error | Out-GridView -Title 'Error Report'
})



#ResetServerListData Event
$uiHash.ResetDataMenu.Add_Click({
    Write-Verbose "Resetting Server List data"
    $uiHash.Listview.ItemsSource | ForEach {$_.Notes = $null; $_.Status = $null; $_.Audited = 0; $_.Installed = 0; $_.InstallErrors = 0; $_.Services = 0}
    $uiHash.Listview.Items.Refresh()
})

#RMS
#StartSEPMasterService Event
$uiHash.StartSEPMasterServiceMenu.Add_Click({
StartSEPMasterService
})

#RepairSEP Event
$uiHash.RepairSEPMenu.Add_Click({
Show-ExecWarning "RepairSEP"
})

#EndSepScanService Event
$uiHash.EndSepScanServiceMenu.Add_Click({
Show-ExecWarning "EndSepScanService"
})


#SEPHealthCheck Event
$uiHash.SEPHealthCheckMenu.Add_Click({
SEPHealthCheck
})

#SEPVirusDefFoldersCheck Event
$uiHash.SEPVirusDefFoldersCheckMenu.Add_Click({
SEPVirusDefFoldersCheck
})

#SEPOLDVirusDefDELETE Event
$uiHash.SEPOLDVirusDefDELETEMenu.Add_Click({
Show-ExecWarning "SEPOLDVirusDefDELETE"
})

#TM EDR Prod Agent Deployment Event
$uiHash. TMEDRProdAgentDeployentMenu.Add_Click({
Show-ExecWarning "TMEDRProdAgentDeployment"
})

#TrendMicroDSAHealthCheck Event
$uiHash.TrendMicroDSAHealthCheckMenu.Add_Click({
TrendMicroDSAHealthCheck
}) 

#Check Patch Log Event
$uiHash.CheckPatchLogMenu.Add_Click({
Check-PatchLogTwentyFive
})

#RunChefClient Event
$uiHash.RunChefClientMenu.Add_Click({
RunChefClient
})

#Check ASP version Event
$uiHash.ASPVerMenu.Add_Click({
Check-ASPVer
})

#CheckWSUS Event
$uiHash.WSUSRegMenu.Add_Click({
Check-wsusregistry
})

#CheckDotNetRuntimeOnOS Event
$uiHash.DotNetRuntimeOnOSMenu.Add_Click({
Check-DotNetRuntimeOnOS
})

#CheckAllDotNetVersions Event
$uiHash.DotnetchkAllmenu.Add_Click({
Check-AllDotNet
})

#CheckDotNetAddOns Event
$uiHash.DotNetAddOnsMenu.Add_Click({
Check-DotNetAddOns
})

#Check MsOfficeVersion Event
$uiHash.MsOfficeMenu.Add_Click({
Check-MsOfficeVersion
})

#Check CBSLogs Event
$uiHash.CBSlogsMenu.Add_Click({
Check-CBSLogs
})

#Check PatchingRebootFlag Event
$uiHash.CheckPatchingRebFlagMenu.Add_Click({
Check-PatchingRebootFlag
})

#Delete Reboot Flag Event
$uiHash.DelRebootFlagMenu.Add_Click({
Show-ExecWarning "Delete-RebootFlag"
})

#Check Ruby Process Event
$uiHash.CheckRubyProcessMenu.Add_Click({
Check-RubyProcess
})

#END Ruby Process Event
$uiHash.ENDRubyProcessMenu.Add_Click({
Show-ExecWarning "END-RubyProcess"
})


#Check CHEF Env Variable Event
$uiHash.CheckCHEFenvVarMenu.Add_Click({
Check-CHEFenvVar
})

#ADD Missing CHEF Env Var Event
$uiHash.AddMissingCHEFEnvVarMenu.Add_Click({
Show-ExecWarning "AddMissingCHEFEnvVar"
})


#RunWinUpdateFix on Multiple Servers Event
$uiHash.WinUpdateFixMultiServersMenu.Add_Click({
Show-ExecWarning "WinUpdateFix-MultipleServers"
})

#RunWinUpdateFix Event
$uiHash.WinUpdateFixMenu.Add_Click({
Show-ExecWarning "WinUpdateFix"
})

#CheckEDGE Event
$uiHash.EDGECheckMenu.Add_Click({
Check-EDGE
})

#CheckIIS Event
$uiHash.IISCheckMenu.Add_Click({
Check-IIS
})

#IIS Install Event
$uiHash.IISInstallMenu.Add_Click({
Show-ExecWarning "Install-IIS"
})

#Win2012ESUYear1LicenseCheck Event
$uiHash.Win2012ESUYear1LicenseCheckMenu.Add_Click({
Check-Win2012ESUYear1License
})

#Win2012ESUYear1ProductKeyInstall Event
$uiHash.Win2012ESUYear1ProductKeyInstallMenu.Add_Click({
Show-ExecWarning "Install-Win2012ESUYear1ProductKey"
})

#CheckWinDefend Event
$uiHash.WinDefendCheckMenu.Add_Click({
Check-WinDefend
})

#WinDefendOffboardMenu Event
$uiHash.WinDefendOffboardMenu.Add_Click({
Show-ExecWarning "WinDefendOffboard"
})

#WinDefendUninstall Event
$uiHash.WinDefendUninstallMenu.Add_Click({
Show-ExecWarning "WinDefendUninstall"
})
#endregion        

#Start the GUI
$uiHash.Window.ShowDialog() | Out-Null