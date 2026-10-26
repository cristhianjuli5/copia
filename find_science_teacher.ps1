$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$dir = "c:\Users\analista.tic\Downloads\automatizacion"
$wbPh = $excel.Workbooks.Open((Join-Path $dir "Archivo_Phidias.xls.xlsx"), 0, $true)
$sPh = $wbPh.Sheets.Item(1)

$curTeacher = ""
$curSub = ""
for ($r = 1; $r -le $sPh.UsedRange.Rows.Count; $r++) {
    $c = $sPh.Cells.Item($r, 1)
    $t = $c.Text
    if (-not $t) { continue }
    $indent = $c.IndentLevel
    if ($indent -eq 0) { $curTeacher = $t }
    elseif ($indent -eq 1) { $curSub = $t }
    elseif ($indent -eq 2) {
        if ($t -like "*Second (2)*" -or $t -like "*Third (3)*") {
            $m = $sPh.Cells.Item($r, 13).Text
            Write-Host "Teacher: '$curTeacher' | Subject: '$curSub' | Course: '$t' | M='$m'"
        }
    }
}

$wbPh.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
