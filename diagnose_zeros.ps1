$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$dir = "c:\Users\analista.tic\Downloads\automatizacion"
$pathAct = Join-Path $dir "PROGRESO_POR_PERIODO_ACTUALIZADO.xlsx"
$wb = $excel.Workbooks.Open($pathAct, 0, $true)

$sheets = @("MATEMATICAS", "INGLES", "FRANCES", "LENGUA Y LITERATURA", "CIENCIAS", "SOCIALES", "TECNOLOGIA", "EDU_FISICA", "ARTES")

$report = @()

foreach ($shName in $sheets) {
    $s = $wb.Sheets.Item($shName)
    $lastCol = $s.UsedRange.Columns.Count
    # Find Ciclo 24 col
    $c24Col = 0
    $c23Col = 0
    for ($c = 1; $c -le $lastCol; $c++) {
        $t = $s.Cells.Item(1, $c).Text
        if ($t -like "*CICLO 24*") { $c24Col = $c }
        if ($t -like "*CICLO 23*") { $c23Col = $c }
    }
    
    $zeroEspCount = 0
    $zeroPhiCount = 0
    $totalRows = 0
    $zeroExamples = @()

    $nRows = $s.UsedRange.Rows.Count
    for ($r = 6; $r -le $nRows; $r++) {
        $curso = $s.Cells.Item($r, 2).Text
        $doc = $s.Cells.Item($r, 3).Text
        if (-not $curso -and -not $doc) { continue }
        $totalRows++

        $esp24 = $s.Cells.Item($r, $c24Col).Text
        $phi24 = $s.Cells.Item($r, $c24Col + 2).Text
        
        $esp23 = if ($c23Col -gt 0) { $s.Cells.Item($r, $c23Col).Text } else { "" }
        $phi23 = if ($c23Col -gt 0) { $s.Cells.Item($r, $c23Col + 2).Text } else { "" }

        $isZeroEsp = ($esp24 -eq "0" -or $esp24 -eq "")
        $isZeroPhi = ($phi24 -eq "0" -or $phi24 -eq "")

        if ($isZeroEsp) { $zeroEspCount++ }
        if ($isZeroPhi) { $zeroPhiCount++ }

        if (($isZeroEsp -and $esp23 -ne "0" -and $esp23 -ne "") -or ($isZeroPhi -and $phi23 -ne "0" -and $phi23 -ne "")) {
            $zeroExamples += "Row $($r): Curso='$curso', Docente='$doc' -> C23(Esp=$esp23, Phi=$phi23) vs C24(Esp=$esp24, Phi=$phi24)"
        }
    }

    Write-Host "================== SHEET: $shName (Total rows: $totalRows) =================="
    Write-Host "  C24 Col: $c24Col | C23 Col: $c23Col"
    Write-Host "  Filas con ESPERADO = 0: $zeroEspCount / $totalRows"
    Write-Host "  Filas con PHIDIAS = 0:  $zeroPhiCount / $totalRows"
    if ($zeroExamples.Count -gt 0) {
        Write-Host "  Ejemplos donde antes NO era 0 y ahora dio 0:"
        $zeroExamples[0..([Math]::Min(10, $zeroExamples.Count - 1))] | ForEach-Object { Write-Host "    $_" }
    }
}

$wb.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
