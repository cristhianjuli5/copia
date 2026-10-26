$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$dir = "c:\Users\analista.tic\Downloads\automatizacion"
$wbPh = $excel.Workbooks.Open((Join-Path $dir "Archivo_Phidias.xls.xlsx"), 0, $true)
$sPh = $wbPh.Sheets.Item(1)

$f = $sPh.Cells.Find("Reyes Ramirez")
$sr = $f.Row
for ($r = $sr; $r -le $sr + 25; $r++) {
    $c = $sPh.Cells.Item($r, 1)
    Write-Host "Row $r [Indent=$($c.IndentLevel)]: '$($c.Text)' | M='$($sPh.Cells.Item($r, 13).Text)'"
}

$wbPh.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
