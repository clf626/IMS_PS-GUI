Invoke-Command -ComputerName APSES3716 -ScriptBlock {Get-ItemProperty HKLM:\SOFTWARE\Wow6432node\Microsoft\Windows\CurrentVersion\uninstall\* | 
Select DisplayName, DisplayVersion} 
| Where-Object {$_.DisplayName -like 'Adobe*Reader*'})."DisplayName"} | Format-Table -AutoSize}

Invoke-Command -ComputerName APSES3716 -ScriptBlock {
$ReaderMUIname = (Get-ItemProperty HKLM:\SOFTWARE\Wow6432node\Microsoft\Windows\CurrentVersion\uninstall\* |  Where-Object {$_.DisplayName -like 'Adobe*Reader*'} )."DisplayVersion"
                        #$ReaderMUIVersion = (Get-ItemProperty HKLM:\SOFTWARE\Wow6432node\Microsoft\Windows\CurrentVersion\uninstall\* | 
                        #Select DisplayVersion | Where-Object {$_.DisplayVersion -like 'Adobe*Reader*'})
                        return $resout = "Adobe Reader Version: $ReaderMUIname"}


 Invoke-Command -ComputerName APSES3851 -ScriptBlock {
 $ReaderMUIname = (Get-WmiObject -Class win32_product -Filter "name LIKE '%Flash%'")."name"
 $ReaderMUIversion = (Get-WmiObject -Class win32_product -Filter "name LIKE '%Flash%'")."version"
 return $resout = "$ReaderMUIname Version: $ReaderMUIversion"}


 Invoke-Command -ComputerName APSES3716 -ScriptBlock {
 Get-ItemProperty HKLM:\SOFTWARE\Wow6432node\Microsoft\Windows\CurrentVersion\uninstall\*
 }

 Invoke-Command -ComputerName APSES3716 -ScriptBlock {
 wmic product where "Vendor like '%Adobe%Reader%'" get name "," version}
 
 (wmic /node:"APSES3716" product where "Vendor like '%Adobe%'" get name "," version)."Name"


 Get-WmiObject -Class win32_product -ComputerName APSES5195 

 Get-ItemProperty 'HKLM:\SOFTWARE\Macromedia\FlashPlayerPlugin' -Computer APSES5195