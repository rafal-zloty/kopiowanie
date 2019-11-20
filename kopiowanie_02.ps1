$baseDir = 'C:\rafal\workspaces\training-ground-2018-09\excel\'

$przeliczeniaExcel = New-Object -ComObject Excel.Application
$przeliczeniaWorkbook = $przeliczeniaExcel.Workbooks.Open($baseDir + 'przeliczenia.xlsx')
$przeliczeniaWorksheet = $przeliczeniaWorkbook.WorkSheets.item('dane z pomiarow (V)')
$przeliczeniaWorksheet.activate()

#(Nazwa pliku z pomiarami, nazwa arkusza z pomiarami, zakres kolumn w arkuszu z pomiarami, komórke docelowa z pliku z przeliczeniami )
$pomiaryList = @(@('pomiar1.xlsx','Arkusz1','A1:G1','A4'), @('pomiar2.xlsx','Arkusz1','A1:F1','A64'), @('pomiar2.xlsx','Arkusz1','G1','H4'), @('pomiar3.xlsx','Arkusz1','A1:F1','A124'), @('pomiar3.xlsx','Arkusz1','G1','I4'))

ForEach ($pomiar in $pomiaryList){
  $pomiarExcel = New-Object -ComObject Excel.Application
  $pomiarWorkbook = $pomiarExcel.Workbooks.Open($baseDir + $pomiar[0])
  $pomiarWorksheet = $pomiarWorkbook.WorkSheets.item($pomiar[1])
  $pomiarWorksheet.activate()
  $pomiarRange = $pomiarWorksheet.Range($pomiar[2]).EntireColumn
  $pomiarRange.Copy() | out-null

  $przeliczeniaRange = $przeliczeniaWorksheet.Range($pomiar[3])
  $przeliczeniaWorksheet.Paste($przeliczeniaRange)
  $przeliczeniaWorkbook.Save()
}

$przeliczeniaWorkbook.Close()
$przeliczeniaExcel.Quit()
