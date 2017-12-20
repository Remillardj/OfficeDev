$file = "C:\path\to\excel\doc"
$sheetName = "SheetName"

$objExcel = New-Object -ComObject Excel.Application
$workbook = $objExcel.Workbooks.Open($file)

$sheet = $workbook.Worksheets.Item($sheetName)
$objExcel.Visible = $False

$rowMax = ($sheet.UsedRange.Rows).count


function getNetworkNames {
   Param()
   $ExcelSourceFile = $file
   $SheetName = $sheetName
   [System.Threading.Thread]::CurrentThread.CurrentCulture = [System.Globalization.CultureInfo]'en-US'
   $ExcelSourceObj = New-Object -comobject Excel.Application
   $ExcelSourceObj.Visible = $True
   $ExcelWorkbook = $ExcelSourceObj.Workbooks.Open($ExcelSourceFile, 2, $True)
   $ExcelWorkSheet = $ExcelWorkbook.Worksheets.Item($SheetName)

   $Row = 1
   $Column = 1
   $Found = $False
   while ($ExcelWorkSheet.Cells.Item($Row, $Column).Value() -ne $Null)
   {
        Write-Host $ExcelWorkSheet.Item($Row, $Column).Value()
        $Found = $True
   }
}

$xl = New-Object -COM "Excel.Application"
$xl.Visible = $true
$wb = $xl.Workbooks.Open($file)
$ws = $wb.Sheets.Item(1)

for ($i = 1; $i -le 3; $i++)
{
    if ($ws.Cells.Item($i, 1).Value -eq $num) {
        echo $ws.Cells.Item($i, 2).Value
        break
    }
}

$wb.Close()
$xl.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl)
