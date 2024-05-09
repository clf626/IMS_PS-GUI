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
              if($OSVersion -like "6.1*" -OR $OSVersion -like "6.0*") #Windows Server 2008 R2 or Non R2
              {
                
                    $preLicChck = try{ 
                        $prechckLic = Invoke-Command -cn $ServerName -ScriptBlock{cmd.exe /c "cscript /Nologo c:\Windows\System32\slmgr.vbs /dli 04fa0286-fa74-401e-bbe9-fbfbb158010d"}
                        $prechckLic | Select-String -pattern "License"
                        
                    }Catch{
                        $errorcount += 1
                        $PreLicNote = "Pre-License Check:" + $_.Exception.Message | Out-String 
                    }
                    $preLicStatus = $preLicChck | ForEach-Object {$_.line.Substring($_.Line.LastIndexOf(':')+2)}
               
                    if($preLicStatus -eq "Licensed")
                    {
                        $temp = "" | Select  ComputerName, OS_Caption, OS_Architecture, PK_Install, PK_Activate, ESU_License, Notes, Error
                                    $temp.ComputerName = $ServerName
                                    $temp.OS_Caption = $OSCaption
                                    $temp.OS_Architecture = $OSArchitecture
                                    $temp.PK_Install = "Skipped"
                                    $temp.PK_Activate = "Skipped"
                                    $temp.ESU_License = $preLicStatus
                                    $temp.Notes = ""
                                    $temp.Error = $errorcount
                                    $reports += $temp
                    
                    }Else{ 
                             $runIPK = try{
                             Invoke-Command -cn $ServerName -ScriptBlock{cmd.exe /c "cscript /Nologo c:\Windows\System32\slmgr.vbs /ipk 86P8P-2BP7W-8VKTD-93MDY-3MYQ4"}
                             
                             }Catch{
                             $errorcount += 1
                             $runIPKNote = "PKey Install: " + $_.Exception.Message | Out-String}

                             if($runIPK -match "success"){
                                 $pkInstall = "Product Key successfully applied" 
                                 $runActvte = try{
                                     Invoke-Command -cn $ServerName -ScriptBlock{cmd.exe /c "cscript /Nologo c:\Windows\System32\slmgr.vbs /ato 04fa0286-fa74-401e-bbe9-fbfbb158010d"}
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
                                 $chckLic = Invoke-Command -cn $ServerName -ScriptBlock{cmd.exe /c "cscript /Nologo c:\Windows\System32\slmgr.vbs /dli 04fa0286-fa74-401e-bbe9-fbfbb158010d"}
                                 $chckLic | Select-String -pattern "License"
                                 }Catch{
                                 $errorcount += 1
                                 $PostLicNote = "Post License Check: " + $_.Exception.Message | Out-String }
                             $LicStatus = $LicChck | ForEach-Object {$_.line.Substring($_.Line.LastIndexOf(':')+2)}

                             $temp = "" | Select  ComputerName, OS_Caption, OS_Architecture, PK_Install, PK_Activate, ESU_License, Notes, Error
                             $temp.ComputerName = $ServerName
                             $temp.OS_Caption = $OSCaption
                             $temp.OS_Architecture = $OSArchitecture
                             $temp.PK_Install = $pkInstall
                             $temp.PK_Activate = $pkActivate
                             $temp.ESU_License = $LicStatus
                             $temp.Notes = $PreLicNote + " " + $runIPKNote + " " + $runATONote + " " + $PostLicNote +" "+ $patchInstallNote +" "+ $patchDLNote
                             $temp.Error = $errorcount
                             $reports += $temp
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