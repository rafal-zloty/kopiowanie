
$Excel = New-Object -ComObject Excel.Application
$Workbook = $Excel.Workbooks.Open('C:\do III stopnia\badania\przeliczenia\pomiar 001_cz01.xls')

#$excel.Visible = $false
$Worksheet = $Workbook.WorkSheets.item(“Arkusz1”)
$worksheet.activate()
$range = $WorkSheet.Range(“A1:G1”).EntireColumn
$range.Copy() | out-null


$Excel = New-Object -ComObject Excel.Application
$Workbook = $Excel.Workbooks.Open('C:\do III stopnia\badania\przeliczenia\przeliczenia pomiarow.xlsm')
$Worksheet = $Workbook.WorkSheets.item(“dane z pomiarow (V)”)
#$excel.Visible = $true
$worksheet.activate()
$Range = $Worksheet.Range(“A4”)
$Worksheet.Paste($range) 
$workbook.Save()
$workbook.Close()
$excel.Quit()

$Excel = New-Object -ComObject Excel.Application
$Workbook = $Excel.Workbooks.Open('C:\do III stopnia\badania\przeliczenia\pomiar 001_cz02.xls')

#$excel.Visible = $false
$Worksheet = $Workbook.WorkSheets.item(“Arkusz1”)
$worksheet.activate()
$range = $WorkSheet.Range(“A1:F1”).EntireColumn
$range.Copy() | out-null


$Excel = New-Object -ComObject Excel.Application
$Workbook = $Excel.Workbooks.Open('C:\do III stopnia\badania\przeliczenia\przeliczenia pomiarow.xlsm')
$Worksheet = $Workbook.WorkSheets.item(“dane z pomiarow (V)”)
#$excel.Visible = $true
$worksheet.activate()
$Range = $Worksheet.Range(“A64”)
$Worksheet.Paste($range) 
$workbook.Save()
$workbook.Close()
$excel.Quit()

$Excel = New-Object -ComObject Excel.Application
$Workbook = $Excel.Workbooks.Open('C:\do III stopnia\badania\przeliczenia\pomiar 001_cz02.xls')

#$excel.Visible = $false
$Worksheet = $Workbook.WorkSheets.item(“Arkusz1”)
$worksheet.activate()
$range = $WorkSheet.Range(“G1”).EntireColumn
$range.Copy() | out-null


$Excel = New-Object -ComObject Excel.Application
$Workbook = $Excel.Workbooks.Open('C:\do III stopnia\badania\przeliczenia\przeliczenia pomiarow.xlsm')
$Worksheet = $Workbook.WorkSheets.item(“dane z pomiarow (V)”)
#$excel.Visible = $true
$worksheet.activate()
$Range = $Worksheet.Range(“H4”)
$Worksheet.Paste($range) 
$workbook.Save()
$workbook.Close()
$excel.Quit()

$Excel = New-Object -ComObject Excel.Application
$Workbook = $Excel.Workbooks.Open('C:\do III stopnia\badania\przeliczenia\pomiar 001_cz03.xls')

#$excel.Visible = $false
$Worksheet = $Workbook.WorkSheets.item(“Arkusz1”)
$worksheet.activate()
$range = $WorkSheet.Range(“A1:F1”).EntireColumn
$range.Copy() | out-null


$Excel = New-Object -ComObject Excel.Application
$Workbook = $Excel.Workbooks.Open('C:\do III stopnia\badania\przeliczenia\przeliczenia pomiarow.xlsm')
$Worksheet = $Workbook.WorkSheets.item(“dane z pomiarow (V)”)
#$excel.Visible = $true
$worksheet.activate()
$Range = $Worksheet.Range(“A124”)
$Worksheet.Paste($range) 
$workbook.Save()
$workbook.Close()
$excel.Quit()

$Excel = New-Object -ComObject Excel.Application
$Workbook = $Excel.Workbooks.Open('C:\do III stopnia\badania\przeliczenia\pomiar 001_cz03.xls')

#$excel.Visible = $false
$Worksheet = $Workbook.WorkSheets.item(“Arkusz1”)
$worksheet.activate()
$range = $WorkSheet.Range(“G1”).EntireColumn
$range.Copy() | out-null


$Excel = New-Object -ComObject Excel.Application
$Workbook = $Excel.Workbooks.Open('C:\do III stopnia\badania\przeliczenia\przeliczenia pomiarow.xlsm')
$Worksheet = $Workbook.WorkSheets.item(“dane z pomiarow (V)”)
#$excel.Visible = $true
$worksheet.activate()
$Range = $Worksheet.Range(“I4”)
$Worksheet.Paste($range) 
$workbook.Save()
$workbook.Close()
$excel.Quit()