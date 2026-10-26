$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wbPl = $excel.Workbooks.Open('c:\Users\analista.tic\Downloads\automatizacion\PLANTILLA CA PROGRESO POR PERIODO.xlsx', 0, $true)

foreach ($shPl in $wbPl.Sheets) {
    $pName = $shPl.Name
    $cycleCol = 0
    $maxCols = [Math]::Min(35, $shPl.UsedRange.Columns.Count)
    for ($c = 1; $c -le $maxCols; $c++) {
        $h2 = $shPl.Cells.Item(2, $c).Text
        $h3 = $shPl.Cells.Item(3, $c).Text
        $comb = "$h2 $h3"
        if ($comb -match "C24\b" -or $comb -match "Ciclo 24\b" -or $comb -match "subido C24\b" -or $comb -match "Reporte Día 1 C24\b") {
            $subHNext = $shPl.Cells.Item(4, $c + 1).Text
            $subHCur = $shPl.Cells.Item(4, $c).Text
            if ($subHNext -like "*Acum*") { $cycleCol = $c + 1 }
            elseif ($subHCur -like "*Acum*") { $cycleCol = $c }
            else { $cycleCol = $c + 1 }
            break
        }
    }
    
    if ($cycleCol -gt 0) {
        Write-Host "Hoja: $pName | Columna: $cycleCol"
        $nRows = $shPl.UsedRange.Rows.Count
        for ($r = 4; $r -le $nRows; $r++) {
            $col1 = $shPl.Cells.Item($r, 1).Text
            $val = $shPl.Cells.Item($r, $cycleCol).Text
            if ($col1 -and $col1 -notmatch "Total") {
                Write-Host "  $col1 -> $val"
            }
        }
    }
}
$wbPl.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
