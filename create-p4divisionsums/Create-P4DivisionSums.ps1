# This will be the number (denominator) in the first column. Eg 4..9 will generate numbers 4,5,6,7,8,9
$denominators = @(4..9)
# This will be the number (numerator) in the third column. Eg 1..9 will generate numbers 1,2,3,4,5,6,7,8,9. The number will be 2 digit if using 1..9
$numerators = @(1..9)
# This is the amount of multiplication sums to create
$amt = 20

### DON'T MODIFY ANYTHING BELOW THIS LINE ###

Clear-Host

    try{
    Write-Output "[INFO] Creating division sums."
    $excel = New-Object -ComObject Excel.Application -ErrorAction SilentlyContinue
    }
    catch{
    Write-Output "[ERROR] Unable to create the Excel document. Check that Excel is installed. Script terminated!"
    Write-Output "[ERROR] $($_.exception.message)"
    break
    }

$excel.Visible = $True
$workbook = $excel.Workbooks.Add()
$worksheet = $workbook.Worksheets.Item(1)
$worksheet.Cells.Item(1,1) = "Write answer in the format 5 r 2, or 5 r 0. Do not write 05 r 2 (drop the leading zero)."

$i = 1
$y = 3

    while($i -le $amt){

    # Get the numbers
    $num1 = $numerators | Get-Random
    $num2 = $numerators | Get-Random
    $num = "$num1$num2"
    $den1 = $denominators | Get-Random
    # Do the math. Get the whole number and the remainder
    $whole = [math]::Floor($num/$den1)
    $remainder = $num % $den1
    $answer = "$whole r $remainder"

    $worksheet.Cells.Item($y,1).ColumnWidth = 3
    #https://devblogs.microsoft.com/scripting/how-can-i-center-text-in-an-excel-cell/
    $worksheet.Cells.Item($y,1).HorizontalAlignment = -4108
    $worksheet.Cells.Item($y,1) = $den1
    $worksheet.Cells.Item($y,2).ColumnWidth = 3
    $worksheet.Cells.Item($y,2).HorizontalAlignment = -4108
    $worksheet.Cells.Item($y,2) = "\"
    $worksheet.Cells.Item($y,3).ColumnWidth = 3
    $worksheet.Cells.Item($y,3).HorizontalAlignment = -4108
    $worksheet.Cells.Item($y,3) = $num
    $worksheet.Cells.Item($y,4).ColumnWidth = 3
    $worksheet.Cells.Item($y,4).HorizontalAlignment = -4108
    $worksheet.Cells.Item($y,4) = "="
    $worksheet.Cells.Item($y,5).ColumnWidth = 13
    $worksheet.Cells.Item($y,5).HorizontalAlignment = -4108
    $worksheet.Cells.Item($y,6).ColumnWidth = 13
    $worksheet.Cells.Item($y,6).HorizontalAlignment = -4108
    $worksheet.Cells.Item($y,6) = "=IF(E$($i)=`"`",`"Please answer`",IF(E$($i)<>G$($i),`"Wrong`",`"Correct`"))"
    $worksheet.Cells.Item($y,7).HorizontalAlignment = -4108
    # https://analysistabs.com/excel-vba/colorindex/
    $worksheet.Cells.Item($y,7).Font.ColorIndex = 2
    $worksheet.Cells.Item($y,7) = $answer

    $i++
    $y++

    }

Write-Output "[INFO] Finshed creating multiplication sums. Please go to Excel to begin."
Start-Sleep -Seconds 5