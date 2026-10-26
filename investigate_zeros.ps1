$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$dir = "c:\Users\analista.tic\Downloads\automatizacion"
$wbPh = $excel.Workbooks.Open((Join-Path $dir "Archivo_Phidias.xls.xlsx"), 0, $true)
$sPh = $wbPh.Sheets.Item(1)

function Search-Phidias($term) {
    Write-Host "=== SEARCH PHIDIAS: '$term' ==="
    $matches = @()
    for ($r = 1; $r -le $sPh.UsedRange.Rows.Count; $r++) {
        $t = $sPh.Cells.Item($r, 1).Text
        if ($t -like "*$term*") {
            $matches += "Row $($r): Indent=$($sPh.Cells.Item($r, 1).IndentLevel), Text='$t', ColM='$($sPh.Cells.Item($r, 13).Text)'"
        }
    }
    $matches | ForEach-Object { Write-Host $_ }
}

Search-Phidias "Barrera Garcia"
Search-Phidias "Peña Diaz"
Search-Phidias "Pena Diaz"
Search-Phidias "Méndez Parra"
Search-Phidias "Mendez Parra"
Search-Phidias "Tobacia"
Search-Phidias "Vela Merchán"

$wbPh.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
