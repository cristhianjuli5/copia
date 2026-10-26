$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$dir = "c:\Users\analista.tic\Downloads\automatizacion"
$wbPh = $excel.Workbooks.Open((Join-Path $dir "Archivo_Phidias.xls.xlsx"), 0, $true)
$sPh = $wbPh.Sheets.Item(1)

$subjects = @{}
for ($r = 1; $r -le $sPh.UsedRange.Rows.Count; $r++) {
    $c = $sPh.Cells.Item($r, 1)
    if ($c.IndentLevel -eq 1) {
        $t = $c.Text.Trim()
        # strip (Grade) from end
        $baseName = [System.Text.RegularExpressions.Regex]::Replace($t, "\s*\(.*?\)$", "").Trim()
        if (-not $subjects.ContainsKey($baseName)) {
            $subjects[$baseName] = 0
        }
        $subjects[$baseName]++
    }
}

Write-Host "Distinct Phidias Subjects:"
$subjects.GetEnumerator() | Sort-Object Name | ForEach-Object { Write-Host "  '$($_.Key)' : $($_.Value) times" }

$wbPh.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
