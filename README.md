# Create-Sums
A PowerShell script that will create multiplication sums in an Excel spreadsheet that will auto check the answers.

## Why This Script

As of 23rd March 2020 schools in Scotland are closing due to the worldwide Covid-19 pandemic. So I wanted to create a PowerShell script that my young kids could run themselves to keep their mental arithmetic going. There are plenty of websites that have this type of thing but a lot of them are riddled with adverts and such like and requires the internet. This script once downloaded doesn't. 

## Script Info

Creates and Excel spreadsheet with a set amount of multiplication questions and tells the kids if the answer is correct or not.

## Sample Output

Sample_Multiplication.xlsx shows what the Excel spreadsheet looks like once generated and semi completed. 

## Prerequisites

1. Windows 10.
2. Microsoft Excel.

## Usage
1. Copy script to a location on your computer. 
2. Right click the script once downloaded click Properties.
3. You should already be on the General Tab, if not click the General tab.
4. At the bottom unselect Unblock and click OK.
5. If you don't see this setting then you are good to proceed.
6. Right click the script and select Run With PowerShell. (If you are prompted with a message click Y).
7. The Excel workbook should now be open.
8. If not click Start and type PowerShell.
9. Right click PowerShell and select Run As Administrator.
10. Type Get-ExecutionPolicy and press enter.
11. Take a note of what it says. (Something like restricted).
12. Type Set-ExecutionPolicy RemoteSigned -Force
13. Close PowerShell and retry step 6. 
14. To revert the PowerShell setting that was changed follow these steps.
15. Click Start and type PowerShell.
16. Right click PowerShell and select Run As Administrator.
17. Type Set-ExecutionPolicy x -Force (x will be the value you noted down in step 11, if in doubt replace x with Restricted).
18. Close PowerShell.

## Future Updates

- Add in Addition, Subtraction and Division. 

## Feedback

Please use GitHub Issues to report any, well.... issues with the script.
