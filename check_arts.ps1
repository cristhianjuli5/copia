$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$dir = "c:\Users\analista.tic\Downloads\automatizacion"
$wbPh = $excel.Workbooks.Open((Join-Path $dir "Archivo_Phidias.xls.xlsx"), 0, $true)
$sPh = $wbPh.Sheets.Item(1)

function Check-Teacher($name) {
    Write-Host "=== Check: $name ==="
    $f = $sPh.Cells.Find($name)
    if ($f) {
        for ($r = $f.Row; $r -le $f.Row + 20; $r++) {
            $t = $sPh.Cells.Item($r, 1).Text
            $lvl = $sPh.Cells.Item($r, 1).IndentLevel
            $m = $sPh.Cells.Item($r, 13).Text
            Write-Host "  R$($r) [$lvl]: '$t' | M='$m'"
            if ($lvl -eq 0 -and $r -gt $f.Row) { break }
        }
    } else {
        Write-Host "Not found: $name"
    }
}

Check-Teacher "Natalia"
Check-Teacher "Valero"
Check-Teacher "David Andres"

$wbPh.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
