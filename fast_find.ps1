$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$dir = "c:\Users\analista.tic\Downloads\automatizacion"
$wbPh = $excel.Workbooks.Open((Join-Path $dir "Archivo_Phidias.xls.xlsx"), 0, $true)
$sPh = $wbPh.Sheets.Item(1)

foreach ($term in @("Barrera", "Mendez", "MÃ©ndez", "Tenth (10)", "Second (2)")) {
    $found = $sPh.Cells.Find($term)
    if ($found) {
        Write-Host "Found '$term' at row $($found.Row): '$($found.Text)'"
        $firstRow = $found.Row
        for ($r = [Math]::Max(1, $firstRow - 2); $r -le $firstRow + 10; $r++) {
            Write-Host "  R$($r): [$($sPh.Cells.Item($r, 1).IndentLevel)] '$($sPh.Cells.Item($r, 1).Text)' | M='$($sPh.Cells.Item($r, 13).Text)'"
        }
    } else {
        Write-Host "NOT FOUND: '$term'"
    }
}

$wbPh.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
