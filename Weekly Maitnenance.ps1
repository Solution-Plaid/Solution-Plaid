#Grab a Program Called "Shell Application"
$objShell=New-Object -ComObject Shell.Application
#Grab a Name Space of "Shell Application Object"
$objShellFolder=$objShell.namespace(0xA)
#Assign a Variable that Reflects File Path to Temporary Folder
$Temp=get-ChildItem "env:\TEMP"
#Assign a Variable that is Inside a TEMP Folder
$Temp2=$Temp.Value 
#Assign a Variable to the File Path of Window TEMP Directory
$WinTemp="c:\Windows\Temp\*" 
#Assign a Variable to Today's Date
$Todays_Date=Get-Date
#Enable C Drive to Create Restore Point
Enable-ComputerRestore -Drive "C:\"
#Creating a Restore Point
Checkpoint-Computer -Description "RestorePoint $Todays_Date " -RestorePointType "MODIFY_SETTINGS"
#To Remove Any TEMP Files Within Users TEMP Folder
Remove-Item -Recurse "$Temp2\*" -Force -Verbose 
#Empty the Recycle Bin
$objShellFolder.Items() | %{ remove-item $_.path -Recurse -Confirm:$false}
#Remove Windows TEMP Directory
Remove-Item -Recurse $WinTemp -Force
#Run Disk Cleanup Tool
Cleanmgr /sagerun:1 | out-null 