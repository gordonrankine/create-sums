# Create-MultiplicationSums
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

## Script Customisation
By default the script will effectively only generate 4 to 8 multiplication sums (that's what my kids were learning at school). But it is easy to modify this. There are 3 settings that can be easily changed. 

1. The number in the first column. Currently can only be 4,5,6,7,8.
2. The number in the second column. Currently can only be 1,2,3,4,5,6,7,8,9,10,11,12.
3. The number of sums in each workbook. 

Instructions provided in the Usage section below.

## Usage

### Downloading and using the script

1. Copy Create-Sums.ps1 to a location on your computer. 
2. Right click Create-Sums.ps1 once downloaded click Properties.
3. You should already be on the General Tab, if not click the General tab.
4. At the bottom unselect Unblock and click OK.
5. If you don't see this setting then you are good to proceed.
6. Right click Create-Sums.ps1 and select Run With PowerShell. (If you are prompted with a message click Y).
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

### Customising the script

1. These steps assume you have downloaded the Create-Sums.ps1 to your computer.
2. Right click Create-Sums.ps1 and click Edit. This will open PowerShell ISE.
3. To change the first set of numbers look at line 2. $n1 = 4..8
  - Replace 4 and 8 with the new numbers you want in the first column.
  - For example, if you want these numbers to be 1 to 10 change the line to read $n1 = 1..10
  - That's it. Remember to have two periods between the numbers. 
  - Click File then click Save.
  - Give it a try. You can do this from the open PowerShell window by pressing F5. 
  - Otherwise close PowerShell ISE and follow the steps above in Downloading and using the script.
4. To change the second set of numbers look at line 4. $n2 = 1..12
  - Replace 1 and 12 with the new numbers you want in the second column.
  - For example, if you want these numbers to be 3 to 11 change the line to read $n2 = 3..11
  - That's it. Remember to have two periods between the numbers. 
  - Click File then click Save.
  - Give it a try. You can do this from the open PowerShell window by pressing F5. 
  - Otherwise close PowerShell ISE and follow the steps above in Downloading and using the script.
5. To change the amount of sums that are creating in the Excel spreadsheet look at line 6 $amt = 20
  - Replace 20 with the amount of sums you want generated.
  - For example, if you want 40 sums change the line to read $amt = 40
  - That's it.
  - Click File then click Save.
  - Give it a try. You can do this from the open PowerShell window by pressing F5. 
  - Otherwise close PowerShell ISE and follow the steps above in Downloading and using the script.
  
## Future Updates

- Add in Addition, Subtraction and Division. 

## Feedback

Please use GitHub Issues to report any, well.... issues with the script.
