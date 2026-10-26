$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$dir = "c:\Users\analista.tic\Downloads\automatizacion"
$wbPh = $excel.Workbooks.Open((Join-Path $dir "Archivo_Phidias.xls.xlsx"), 0, $true)
$sPh = $wbPh.Sheets.Item(1)

$peFound = @()
$c = $sPh.Cells.Find("PHYSICAL")
if ($c) {
    $firstAddr = $c.Address
    do {
        $peFound += "Row $($c.Row): '$($c.Text)' | M='$($sPh.Cells.Item($c.Row, 13).Text)'"
        $c = $sPh.Cells.FindNext($c)
    } while ($c -and $c.Address -ne $firstAddr)
}

Write-Host "PE Occurrences in Phidias:"
$peFound | ForEach-Object { Write-Host $_ }

$wbPh.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
