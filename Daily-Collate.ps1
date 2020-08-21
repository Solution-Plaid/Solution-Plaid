# Script Name: Project
# Author: Jin Kim
# Date of Last revision: 08/20/2020
# Description of purpose: To get all the sales report from excel file and arrange to name and total sales. Also outputs daily total.

## VARIABLES
# This is to create a object that has excel properties
$objExcel = New-Object -ComObject Excel.Application 
# Setting the visible false
$objExcel.Visible = $false
# Getting the week using first week of the year as 1 and save it to $week
if(((Get-Date).year)%4 -eq 1){ $week = (Get-Date -UFormat %V) -as [int] $week++ }else{ $week = (Get-Date -UFormat %V) }
# Getting user's one drive location
$Path = "$env:USERPROFILE\OneDrive"
# Getting todays date in format of yyyy-mm-dd
$today_report =$(Get-Date -Format yyyy-MM-dd)
# This is to absolute path for sales file
$ExcelFile = "$Path\WeekReport$week\$today_report.xlsx"
# This is to absolute path for DEMO file
$DemoFile = "$Path\WeekReport$week\Demo.xlsx"
# This is to create an object call workbook using the demo excel file
$WorkBook = $objExcel.Workbooks.Open($DemoFile);
# This is absolute path for textfile that will store the collate total.
$filePath = "C:\Users\WHS\Desktop\DailyReport\$today_report.txt";

## FUNCTIONS
<#
.Description
This is to get each persons's name and total of sales they made per day.
FUNCTIONs
#>
function Memo()
{
    foreach ($sheet in @($WorkBook.Worksheets))
    {
        [pscustomobject][ordered]@{
            Name = $sheet.Range("B2").Text | %{ $_.Split(':')[1]; }
            Total = $sheet.Range("H1").Text
        }
    }
}

<#
.Description
Summing all of the sales and append it to the file
#>
function TotalPrice()
{
    $textFile = Get-Content -Path $filePath
    [int]$total = 0;
    for($i = 3; $i -lt $textFile.Count; $i++)
    {
        $total += $textFile[$i].Split('$')[1];
    }
    Add-Content $filePath "All Total:    `$$total`.00";
}


# MAIN
Memo | Out-File $filePath;
TotalPrice;
# Releasing excel application from the memory.
Stop-Process -Name "Excel";

#END