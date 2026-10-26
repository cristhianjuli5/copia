$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$dir = "c:\Users\analista.tic\Downloads\automatizacion"
$path = Join-Path $dir "PROGRESO_POR_PERIODO_ACTUALIZADO.xlsx"
$wb = $excel.Workbooks.Open($path, 0, $true)

foreach ($sh in $wb.Sheets) {
    $nRows = $sh.UsedRange.Rows.Count
    $lastCol = $sh.UsedRange.Columns.Count
    $lastCycleCol = 0
    for ($c = $lastCol; $c -ge 6; $c--) {
        if ($sh.Cells.Item(1, $c).Text -like "*CICLO 24*") {
            $lastCycleCol = $c
            break
        }
    }
    
    if ($lastCycleCol -eq 0) { continue }
    
    $zerosEsp = 0
    $zerosPhi = 0
    for ($r = 6; $r -le $nRows; $r++) {
        $curso = $sh.Cells.Item($r, 2).Text
        $docente = $sh.Cells.Item($r, 3).Text
        if (-not $curso -and -not $docente) { continue }
        
        $esp = $sh.Cells.Item($r, $lastCycleCol).Text
        $phi = $sh.Cells.Item($r, $lastCycleCol + 2).Text
        
        if ($esp -eq "0" -or $esp -eq "") { 
            $zerosEsp++
            Write-Host "[$($sh.Name)] ESPERADO es 0 -> Fila $r | Curso: $curso | Docente: $docente"
        }
        if ($phi -eq "0" -or $phi -eq "") { 
            $zerosPhi++
            Write-Host "[$($sh.Name)] PHIDIAS es 0 -> Fila $r | Curso: $curso | Docente: $docente"
        }
    }
}
$wb.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
