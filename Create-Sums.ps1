# This will be the number in the first column. Eg 4..8 will generate numbers 4,5,6,7,8
$n1 = 4..8
# This will be the number in the third column. Eg 1..12 will generate numbers 1,2,3,4,5,6,7,8,9,10,11,12
$n2 = 1..12
# This is the amount of multiplication sums to create
$amt = 20

### DON'T MODIFY ANYTHING BELOW THIS LINE ###

Clear-Host

    try{
    Write-Output "[INFO] Creating multiplication sums."
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

$i = 1

    while($i -le $amt){

    $worksheet.Cells.Item($i,1).ColumnWidth = 3
    #https://devblogs.microsoft.com/scripting/how-can-i-center-text-in-an-excel-cell/
    $worksheet.Cells.Item($i,1).HorizontalAlignment = -4108
    $worksheet.Cells.Item($i,1) = Get-Random $n1
    $worksheet.Cells.Item($i,2).ColumnWidth = 3
    $worksheet.Cells.Item($i,2).HorizontalAlignment = -4108
    $worksheet.Cells.Item($i,2) = "x"
    $worksheet.Cells.Item($i,3).ColumnWidth = 3
    $worksheet.Cells.Item($i,3).HorizontalAlignment = -4108
    $worksheet.Cells.Item($i,3) = Get-Random $n2
    $worksheet.Cells.Item($i,4).ColumnWidth = 3
    $worksheet.Cells.Item($i,4).HorizontalAlignment = -4108
    $worksheet.Cells.Item($i,4) = "="
    $worksheet.Cells.Item($i,5).ColumnWidth = 13
    $worksheet.Cells.Item($i,5).HorizontalAlignment = -4108
    $worksheet.Cells.Item($i,6).ColumnWidth = 13
    $worksheet.Cells.Item($i,6).HorizontalAlignment = -4108
    $worksheet.Cells.Item($i,6) = "=IF(E$($i)=`"`",`"Please answer`",IF(E$($i)<>G$($i),`"Wrong`",`"Correct`"))"
    $worksheet.Cells.Item($i,7).HorizontalAlignment = -4108
    # https://analysistabs.com/excel-vba/colorindex/
    $worksheet.Cells.Item($i,7).Font.ColorIndex = 2
    $worksheet.Cells.Item($i,7) = "=A$($i)*C$($i)"

    $i++

    }

Write-Output "[INFO] Finshed creating multiplication sums. Please go to Excel to begin."
Start-Sleep -Seconds 5