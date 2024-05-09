Function 2k8_ESU_License {
[cmdletbinding()]
param($Computer)    
Begin {
    If (!$PSBoundParameters['computer']) {
        Write-Verbose "No computer name given, using local computername"
        [string]$computer = $Env:Computername
        }
    #Create container for Report
    Write-Verbose "Creating report collection"
    $reports = @() 
    $errorcount = 0  
    }
Process {
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    ForEach ($ServerName in $Computer) {
        Write-Verbose "Computer: $($ServerName)"
        If (Test-Connection -ComputerName $ServerName -Count 1 -Quiet) 
        {
                $remotesession = new-pssession -cn $ServerName
                $OS = Invoke-Command -Session $remotesession -ScriptBlock{Get-WmiObject -Class Win32_OperatingSystem}
                $remotelastexitcode = invoke-command -ScriptBlock {return $?} -Session $remotesession
                    $OSVersion = $OS.Version #Get the OS Version and Build Number
                    $OSArchitecture = $OS.OSArchitecture #32-bit vs 64-bit
                    $OSCaption = $OS.Caption #Get if Standard or Enterprise
            
            If($remotelastexitcode)
            {
              if($OSVersion -like "6.1*") #Windows Server 2008 R2
              {
                $chkCulprit = Get-HotFix -Id KB4512506 -cn $ServerName -ea SilentlyContinue
                If($?)
                {
                  $ClprtKB = "Yes"
                  Try{
                      $ClprtKBUnins = Invoke-command -cn $ServerName -ScriptBlock{
                        Start-Process wusa.exe -ArgumentList '/KB:4512506 /uninstall /quiet /norestart' -Wait
                        return "Culprit Uninstalled! reboot server at specified window, then re-run ESU"
                      }
                  }Catch{
                      $ClprtKBErr = $_.Exception.Message | Out-String
                      $errorcount += 1
                      $ClprtKBErr | Select-String -pattern "Error"
                      $ClprtKBUnins = "Culprit KB Uninstallation Failed: " + $ClprtKBErr | ForEach-Object {$_.line.Substring($_.Line.LastIndexOf(':')+2)}
                  }
                  $errorcount += 1
                  $temp = "" | Select  ComputerName, OS_Caption, OS_Architecture, Culprit_KB, Patch_Download, Patch_Install, PK_Install, PK_Activate, ESU_License, Notes, Error
                         $temp.ComputerName = $ServerName
                         $temp.OS_Caption = $OSCaption
                         $temp.OS_Architecture = $OSArchitecture
                         $temp.Culprit_KB = $ClprtKB
                         $temp.Patch_Download = "Skipped"
                         $temp.Patch_Install = "Skipped"
                         $temp.PK_Install = "Skipped"
                         $temp.PK_Activate = "Skipped"
                         $temp.ESU_License = "Skipped"
                         $temp.Notes = $ClprtKBUnins
                         $temp.Error = $errorcount
                         $reports += $temp
                }
                Else{ #No Culprit KB, proceed to ESU process
              
                    $ClprtKB = "No"
                    $preLicChck = try{ 
                        $prechckLic = Invoke-Command -cn $ServerName -ScriptBlock{cmd.exe /c "cscript /Nologo c:\Windows\System32\slmgr.vbs /dli 553673ed-6ddf-419c-a153-b760283472fd"}
                        $prechckLic | Select-String -pattern "License"
                        
                    }Catch{
                        $errorcount += 1
                        $PreLicNote = "Pre-License Check:" + $_.Exception.Message | Out-String 
                    }
                    $preLicStatus = $preLicChck | ForEach-Object {$_.line.Substring($_.Line.LastIndexOf(':')+2)}
               
                    if($preLicStatus -ne "Licensed")
                    {

                       $getHF = Get-HotFix -Id kb4536952, kb4474419, kb4490628, kb4538483 -ComputerName $ServerName
                       
                       if($getHF.hotfixid.count -lt 4)
                       {
                           $chkUnzip = test-path -path "\\$ServerName\C$\Temp\617600\*"
                           if($chkUnzip){
                                $patchDL =  "Skipped"
                           }Else
                           {
                                try{
                                   $source = "https://s3api-dmz.uhc.com/thirdpartysw/ESU_2k8R2.zip"
                                   $destination = "\\" + $ServerName + "\C$\Temp\ESU_2k8R2.zip"
                                   $progressPreference = 'silentlyContinue'    
                                   Invoke-WebRequest $source -OutFile $destination

                                   $ZipSource = "\\" + $ServerName + "\C$\Temp\ESU_2k8R2.zip"
                                   $ZipDestination = "\\" + $ServerName + "\C$\Temp\617600"
                                   Expand-Archive -Path $ZipSource -DestinationPath $ZipDestination
                                   $patchDL =  "Completed"
                                 }Catch{
                                   $errorcount += 1
                                   $patchDL = "Failed"
                                   $patchDLNote = "Patch_DL:" + $_.Exception.Message | Out-String }
                           }
                           $chkUnzip = test-path -path "\\$ServerName\C$\Temp\617600\*"  
                           if($chkUnzip){
                                $scrptblck = {
                                            Dism /online /add-package /packagepath:C:\Temp\617600\Windows6.1-KB4474419-v3-x64.cab /NoRestart /quiet
		    	                            Dism /online /add-package /packagepath:C:\Temp\617600\Windows6.1-KB4490628-x64.cab /NoRestart /quiet
		    	                            Dism /online /add-package /packagepath:C:\Temp\617600\Windows6.1-KB4536952-x64.cab /NoRestart /quiet
		    	                            Dism /online /add-package /packagepath:C:\Temp\617600\Windows6.1-KB4538483-x64.cab /NoRestart /quiet

                                            $getHF = Get-HotFix -Id kb4536952, kb4474419, kb4490628, kb4538483
                                                if($getHF.hotfixid.count -lt 4){
                                                    return "Installation completed! Reboot required"
                                                }Else{
                                                    return "Installation completed!"
                                                }
                                       }
                                $remotesession = new-pssession -computername $ServerName
                                $patchInstall = Invoke-Command -ScriptBlock $scrptblck -Session $remotesession
                                $remotelastexitcode = invoke-command -ScriptBlock { return $?} -Session $remotesession
                                 if(!$remotelastexitcode){
                                        $patchInstall =  "Failed"
                                        $patchInstallNote = "Patch_Install Failed"
                                 }
                                 if($patchInstall -match "Reboot required"){
                                        $patchInstallNote = "Pre-req Installed! reboot server at specified window, then re-run ESU."
                                 }
                                 
                           }Else{
                            $patchInstall = "Failed"
                            $patchInstallNote = "Unable to find Pre-required Patches installers."
                            }
                           if($patchInstall -match "Reboot required" -or $patchInstall -match "Failed"){
                           $errorcount += 1
                              $temp = "" | Select  ComputerName, OS_Caption, OS_Architecture, Culprit_KB, Patch_Download, Patch_Install, PK_Install, PK_Activate, ESU_License, Notes, Error
                              $temp.ComputerName = $ServerName
                              $temp.OS_Caption = $OSCaption
                              $temp.OS_Architecture = $OSArchitecture
                              $temp.Culprit_KB = $ClprtKB
                              $temp.Patch_Download = $patchDL
                              $temp.Patch_Install = $patchInstall
                              $temp.PK_Install = "Skipped"
                              $temp.PK_Activate = "Skipped"
                              $temp.ESU_License = "Skipped"
                              $temp.Notes = $patchDLNote + $patchInstallNote
                              $temp.Error = $errorcount
                              $reports += $temp
                           }Else{ 
                                    $runIPK = try{
                                    Invoke-Command -cn $ServerName -ScriptBlock{cmd.exe /c "cscript /Nologo c:\Windows\System32\slmgr.vbs /ipk TC3DD-7MHCC-D9FYP-GWK84-X8J9C"}
                                    
                                    }Catch{
                                    $errorcount += 1
                                    $runIPKNote = "PKey Install: " + $_.Exception.Message | Out-String}

                                    if($runIPK -match "success"){
                                        $pkInstall = "Product Key successfully applied" 
                                        $runActvte = try{
                                            Invoke-Command -cn $ServerName -ScriptBlock{cmd.exe /c "cscript /Nologo c:\Windows\System32\slmgr.vbs /ato 553673ed-6ddf-419c-a153-b760283472fd"}
                                            }Catch{
                                            $errorcount += 1
                                            $runATONote = "AKey Activate: " + $_.Exception.Message | Out-String
                                            }
                                            
                                            if($runActvte -match "success"){
                                                 $pkActivate = "Activation Key successfully applied"
                                            }else{
                                            $errorcount += 1
                                            $ATO = $runActvte | Select-String -pattern "Error:"
                                            $runATONote = $ATO | ForEach-Object {$_.line.Substring($_.Line.LastIndexOf(':')+2)}
                                            $pkActivate = "Failed"}
                                    }else{
                                    $errorcount += 1
                                    $IPK = $runIPK | Select-String -pattern "Error:"
                                    $runIPKNote = $IPK | ForEach-Object {$_.line.Substring($_.Line.LastIndexOf(':')+2)}
                                    $pkInstall = "Failed"
                                    }
                                    $LicChck = try{
                                        $chckLic = Invoke-Command -cn $ServerName -ScriptBlock{cmd.exe /c "cscript /Nologo c:\Windows\System32\slmgr.vbs /dli 553673ed-6ddf-419c-a153-b760283472fd"}
                                        $chckLic | Select-String -pattern "License"
                                        }Catch{
                                        $errorcount += 1
                                        $PostLicNote = "Post License Check: " + $_.Exception.Message | Out-String }
                                    $LicStatus = $LicChck | ForEach-Object {$_.line.Substring($_.Line.LastIndexOf(':')+2)}

                                    $temp = "" | Select  ComputerName, OS_Caption, OS_Architecture, Culprit_KB, Patch_Download, Patch_Install, PK_Install, PK_Activate, ESU_License, Notes, Error
                                    $temp.ComputerName = $ServerName
                                    $temp.OS_Caption = $OSCaption
                                    $temp.OS_Architecture = $OSArchitecture
                                    $temp.Culprit_KB = $ClprtKB
                                    $temp.Patch_Download = $patchDL
                                    $temp.Patch_Install = $patchInstall
                                    $temp.PK_Install = $pkInstall
                                    $temp.PK_Activate = $pkActivate
                                    $temp.ESU_License = $LicStatus
                                    $temp.Notes = $PreLicNote + " " + $runIPKNote + " " + $runATONote + " " + $PostLicNote +" "+ $patchInstallNote +" "+ $patchDLNote
                                    $temp.Error = $errorcount
                                    $reports += $temp
                           }
                       }else{
                            $runIPK = try{
                            Invoke-Command -cn $ServerName -ScriptBlock{cmd.exe /c "cscript /Nologo c:\Windows\System32\slmgr.vbs /ipk TC3DD-7MHCC-D9FYP-GWK84-X8J9C"}
                            }Catch{
                            $errorcount += 1
                            $runIPKNote = "PKey Install: " + $_.Exception.Message | Out-String}

                            if($runIPK -match "success"){
                                   $pkInstall = "Product Key successfully applied" 
                                   $runActvte = try{
                                       Invoke-Command -cn $ServerName -ScriptBlock{cmd.exe /c "cscript /Nologo c:\Windows\System32\slmgr.vbs /ato 553673ed-6ddf-419c-a153-b760283472fd"}
                                       }Catch{
                                       $errorcount += 1
                                       $runATONote = "AKey Activate: " + $_.Exception.Message | Out-String
                                       }
                                       
                                       if($runActvte -match "success"){
                                            $pkActivate = "Activation Key successfully applied"
                                       }else{
                                       $errorcount += 1
                                       $ATO = $runActvte | Select-String -pattern "Error:"
                                       $runATONote = $ATO | ForEach-Object {$_.line.Substring($_.Line.LastIndexOf(':')+2)}
                                       $pkActivate = "Failed"}
                               }else{
                               $errorcount += 1
                               $IPK = $runIPK | Select-String -pattern "Error:"
                               $runIPKNote = $IPK | ForEach-Object {$_.line.Substring($_.Line.LastIndexOf(':')+2)}
                               $pkInstall = "Failed"
                               }
                               $LicChck = try{
                                   $chckLic = Invoke-Command -cn $ServerName -ScriptBlock{cmd.exe /c "cscript /Nologo c:\Windows\System32\slmgr.vbs /dli 553673ed-6ddf-419c-a153-b760283472fd"}
                                   $chckLic | Select-String -pattern "License"
                                   }Catch{
                                   $errorcount += 1
                                   $PostLicNote = "Post License Check: " + $_.Exception.Message | Out-String }
                               $LicStatus = $LicChck | ForEach-Object {$_.line.Substring($_.Line.LastIndexOf(':')+2)}

                       $temp = "" | Select  ComputerName, OS_Caption, OS_Architecture, Culprit_KB, Patch_Download, Patch_Install, PK_Install, PK_Activate, ESU_License, Notes, Error
                       $temp.ComputerName = $ServerName
                       $temp.OS_Caption = $OSCaption
                       $temp.OS_Architecture = $OSArchitecture
                       $temp.Culprit_KB = $ClprtKB
                       $temp.Patch_Download = "Skipped"
                       $temp.Patch_Install = "Skipped"
                       $temp.PK_Install = $pkInstall
                       $temp.PK_Activate = $pkActivate
                       $temp.ESU_License = $LicStatus
                       $temp.Notes = $PreLicNote + " " + $runIPKNote + " " + $runATONote + " " + $PostLicNote +" "+ $patchInstallNote +" "+ $patchDLNote
                       $temp.Error = $errorcount
                       $reports += $temp
                       }
                          
                  }Else{
                        $temp = "" | Select  ComputerName, OS_Caption, OS_Architecture, Culprit_KB, Patch_Download, Patch_Install, PK_Install, PK_Activate, ESU_License, Notes, Error
                        $temp.ComputerName = $ServerName
                        $temp.OS_Caption = $OSCaption
                        $temp.OS_Architecture = $OSArchitecture
                        $temp.Culprit_KB = $ClprtKB
                        $temp.Patch_Download = "Skipped"
                        $temp.Patch_Install = "Skipped"
                        $temp.PK_Install = "Product Key previously applied"
                        $temp.PK_Activate = "Activation Key previously applied"
                        $temp.ESU_License = $preLicStatus
                        $temp.Notes = "Ok to proceed with Patching."
                        $temp.Error = $errorcount
                        $reports += $temp
                    }
                }
              }Else{
                $errorcount += 1
                $temp = "" | Select  ComputerName, OS_Caption, OS_Architecture, Culprit_KB, Patch_Download, Patch_Install, PK_Install, PK_Activate, ESU_License, Notes, Error
                $temp.ComputerName = $ServerName
                $temp.OS_Caption = $OSCaption
                $temp.OS_Architecture = $OSArchitecture
                $temp.Culprit_KB = $ClprtKB
                $temp.Patch_Download = "Not Applicable"
                $temp.Patch_Install = "Not Applicable"
                $temp.PK_Install = "Not Applicable"
                $temp.PK_Activate = "Not Applicable"
                $temp.ESU_License = "Not Applicable"
                $temp.Notes = "This will not work with $OSCaption $OSArchitecture"
                $temp.Error = $errorcount
                $reports += $temp
              }
            }Else{
            $errorcount += 1
            $temp = "" | Select  ComputerName, OS_Caption, OS_Architecture, Culprit_KB, Patch_Download, Patch_Install, PK_Install, PK_Activate, ESU_License, Notes, Error
                       $temp.ComputerName = $ServerName
                       $temp.OS_Caption = "Failed"
                       $temp.OS_Architecture = "Failed"
                       $temp.Culprit_KB = $ClprtKB
                       $temp.Patch_Download = "Skipped"
                       $temp.Patch_Install = "Skipped"
                       $temp.PK_Install = "Skipped"
                       $temp.PK_Activate = "Skipped"
                       $temp.ESU_License = "Skipped"
                       $temp.Notes =  "Error getting OS details"
                       $temp.Error = $errorcount
                       $reports += $temp
            
            }
        }Else{
        $temp = "" | Select  ComputerName, OS_Caption, OS_Architecture, Culprit_KB, Patch_Download, Patch_Install, PK_Install, PK_Activate, ESU_License, Notes, Error
                       $temp.ComputerName = $ServerName
                       $temp.OS_Caption = "Offline"
                       $temp.OS_Architecture = "Offline"
                       $temp.Culprit_KB = "Offline"
                       $temp.Patch_Download = "Offline"
                       $temp.Patch_Install = "Offline"
                       $temp.PK_Install = "Offline"
                       $temp.PK_Activate = "Offline"
                       $temp.ESU_License = "Offline"
                       $temp.Notes = "Offline"
                       $temp.Error = "Offline"
                       $reports += $temp
        }
    }
}
End{
Write-Output $reports}    
}