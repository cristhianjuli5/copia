$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$dir = "c:\Users\analista.tic\Downloads\automatizacion"
$wbPh = $excel.Workbooks.Open((Join-Path $dir "Archivo_Phidias.xls.xlsx"), 0, $true)
$sPh = $wbPh.Sheets.Item(1)

$f = $sPh.Cells.Find("Rodriguez Martinez")
if ($f) {
    for ($r = $f.Row; $r -le $f.Row + 15; $r++) {
        Write-Host "R$($r): [$($sPh.Cells.Item($r, 1).IndentLevel)] '$($sPh.Cells.Item($r, 1).Text)' | M='$($sPh.Cells.Item($r, 13).Text)'"
    }
} else {
    Write-Host "Not found"
}

$wbPh.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
