$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $excel.Workbooks.Open('c:\Users\analista.tic\Downloads\automatizacion\Archivo_Phidias.xls.xlsx', 0, $true)
$s = $wb.Sheets.Item(1)
$teacher = ''
for($r=1; $r-le $s.UsedRange.Rows.Count; $r++) { 
    $lvl = $s.Cells.Item($r, 1).IndentLevel
    $t = $s.Cells.Item($r, 1).Text
    if ($lvl -eq 0) { $teacher = $t } 
    elseif ($lvl -ge 1 -and $teacher -match 'Gil Duran') { 
        Write-Host ('lvl=' + $lvl + ' | ' + $t + ' | M=' + $s.Cells.Item($r, 13).Text) 
    } 
}
$wb.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
