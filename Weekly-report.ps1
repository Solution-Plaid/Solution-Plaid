
$Path    = "$env:USERPROFILE\OneDrive\"

# create folder with week 

#Get week number 
if(((Get-Date).year)%4 -eq 1){ $week = (Get-Date -UFormat %V) -as [int] $week++ }else{ $week = (Get-Date -UFormat %V) } Write-Host $week


New-Item -ItemType "directory" -Path $Path\weekReport$week


# copy the empty template
#change the name of the file in to today name 
$today_report =$(Get-Date -Format yyyy-MM-dd)

Copy-Item $Path\template.xlxs -Destination "c:\Users\$username\OneDrive\weekReport$week\$today_report.xlxs"









