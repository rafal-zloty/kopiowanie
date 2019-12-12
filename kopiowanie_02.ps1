$baseDir = 'sciezka/do/plikow'

$przeliczeniaExcel = New-Object -ComObject Excel.Application
$przeliczeniaWorkbook = $przeliczeniaExcel.Workbooks.Open($baseDir + 'przeliczenia.xlsx')
$przeliczeniaWorksheet = $przeliczeniaWorkbook.WorkSheets.item('dane z pomiarow (V)')
$przeliczeniaWorksheet.activate()

# z kazdego pliku pobieramy pobieramy po dwie grupy danych
$pomiaryKomorki = @('A1:F1', 'G4')
# komorka docelowa w pliku z przeliczeniami
$przeliczeniaKomorki = @('A64', 'H4', 'A124', 'I4'))

$x = 1 #numer pliku z pomiarami
$cz = 1 #numer czujnika z pliku z pomiarami
$pomiar = 'pomiar '#nazwa pliku z pomiarami string - poczatek
$czN = '_cz' #nazwa pliku z pomiarami string - koncowka

#Write-Host ($baseDir + $pomiar + "$x".PadLeft(3, '0') + $czN + "$cz".PadLeft(2, '0')) #- tutaj sprawdzalem jak zapisuje mi się nazwa pliku - jest ok

$ind = 0 #indeks komorki w tablicy $przeliczeniaKomorki

for ($x = 1; $x -le 2 ; $x++) { # petla ktora otwiera kolejny pomiar czyli zmieni 001, 002 etc (na razie od 1 do 2)
  "$x = " + (1 + $x) 
  for ($cz = 1; $cz -le 3 ; $cz++) { # petla ktora otwiera kolejny czujnik w danym pomiarze czyli zmieni cz01, cz02 i cz03 - to zawsze jest od 1 do 3
    "$cz = " + (1 + $cz) 

    $Excel = New-Object -ComObject Excel.Application
    $Workbook = $Excel.Workbooks.Open($baseDir + $pomiar + "$x".PadLeft(3, '0') + $czN + "$cz".PadLeft(2, '0'))
    $Worksheet = $Workbook.WorkSheets.item(“Arkusz1”)
    $worksheet.activate()
    $excel.Visible = $true

	for ($i = 0; $i -le 1 ; $i++) { # petla po kolejnych grupach komorek w pliku z pomiarami
      $pomiarRange = $worksheet.Range($pomiaryKomorki[$i]).EntireColumn
      $pomiarRange.Copy() | out-null

      $przeliczeniaRange = $przeliczeniaWorksheet.Range($przeliczeniaKomorki[$ind++])
      $przeliczeniaWorksheet.Paste($przeliczeniaRange)
	}
	
    $przeliczeniaWorkbook.Save()
  }
}

$przeliczeniaWorkbook.Close()
$przeliczeniaExcel.Quit()
