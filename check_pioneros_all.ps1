$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$dir = "c:\Users\analista.tic\Downloads\automatizacion"
$wbPh = $excel.Workbooks.Open((Join-Path $dir "Archivo_Phidias.xls.xlsx"), 0, $true)
$sPh = $wbPh.Sheets.Item(1)

$curT = ""
$curS = ""
for ($r = 1; $r -le $sPh.UsedRange.Rows.Count; $r++) {
    $c = $sPh.Cells.Item($r, 1)
    $t = $c.Text
    if (-not $t) { continue }
    $lvl = $c.IndentLevel
    if ($lvl -eq 0) { $curT = $t }
    elseif ($lvl -eq 1) { $curS = $t }
    elseif ($lvl -eq 2) {
        if ($t -like "*Pioneros (Pn) A*") {
            $m = $sPh.Cells.Item($r, 13).Text
            Write-Host "Teacher: '$curT' | Subject: '$curS' | M='$m'"
        }
    }
}

$wbPh.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
