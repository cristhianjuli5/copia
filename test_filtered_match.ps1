$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$dir = "c:\Users\analista.tic\Downloads\automatizacion"
$pathProgreso = Join-Path $dir "PROGRESO_POR_PERIODO.xlsx"
$pathPhidias = Join-Path $dir "Archivo_Phidias.xls.xlsx"
$pathPlantilla = Join-Path $dir "PLANTILLA CA PROGRESO POR PERIODO.xlsx"

function Clean-Text([string]$str) {
    if (-not $str) { return "" }
    $normalized = $str.Normalize([System.Text.NormalizationForm]::FormD)
    $sb = New-Object System.Text.StringBuilder
    foreach ($c in $normalized.ToCharArray()) {
        if ([System.Globalization.CharUnicodeInfo]::GetUnicodeCategory($c) -ne [System.Globalization.UnicodeCategory]::NonSpacingMark) {
            [void]$sb.Append($c)
        }
    }
    $s = $sb.ToString().Normalize([System.Text.NormalizationForm]::FormC)
    $s = $s.Replace([char]160, ' ').ToLower().Trim()
    $s = [System.Text.RegularExpressions.Regex]::Replace($s, "[^a-z0-9 ]", " ")
    $s = [System.Text.RegularExpressions.Regex]::Replace($s, "\s+", " ").Trim()
    return $s
}

function Get-GradeAndSection([string]$course) {
    $c = Clean-Text $course
    $grade = ""
    $letter = ""

    if ($c -match "\bpioneros\b|\bpn\b") { $grade = "pn" }
    elseif ($c -match "\bfirst garten\b|\bfg\b") { $grade = "fg" }
    elseif ($c -match "\bgarten\b|\bgr\b") { $grade = "gr" }
    elseif ($c -match "\btransition\b|\btr\b") { $grade = "tr" }
    elseif ($c -match "\bfirst\b|\b1\b") { $grade = "1" }
    elseif ($c -match "\bsecond\b|\b2\b") { $grade = "2" }
    elseif ($c -match "\bthird\b|\b3\b") { $grade = "3" }
    elseif ($c -match "\bfourth\b|\b4\b") { $grade = "4" }
    elseif ($c -match "\bfifth\b|\b5\b") { $grade = "5" }
    elseif ($c -match "\bsixth\b|\b6\b") { $grade = "6" }
    elseif ($c -match "\bseventh\b|\b7\b") { $grade = "7" }
    elseif ($c -match "\beighth\b|\b8\b") { $grade = "8" }
    elseif ($c -match "\bninth\b|\b9\b") { $grade = "9" }
    elseif ($c -match "\btenth\b|\b10\b") { $grade = "10" }
    elseif ($c -match "\beleventh\b|\b11\b") { $grade = "11" }

    if ($c -match "\b([a-g])\b") {
        $letter = $matches[1].ToLower()
    }

    return @{ Grade = $grade; Letter = $letter; ShortCode = "$grade$letter" }
}

# 1. Parse Phidias
$wbPh = $excel.Workbooks.Open($pathPhidias, 0, $true)
$sPh = $wbPh.Sheets.Item(1)
$nPh = $sPh.UsedRange.Rows.Count
$rangeA = $sPh.Range("A1:A$nPh").Value2
$rangeM = $sPh.Range("M1:M$nPh").Value2

$phEntries = @()
$curT = ""
$curS = ""
for ($r = 1; $r -le $nPh; $r++) {
    $t = [string]$rangeA[$r, 1]
    if (-not $t) { continue }
    $lvl = $sPh.Cells.Item($r, 1).IndentLevel
    $m = [string]$rangeM[$r, 1]
    $mInt = 0
    if ($m) { [int]::TryParse($m, [ref]$mInt) | Out-Null }

    if ($lvl -eq 0) {
        $curT = $t.Trim()
    } elseif ($lvl -eq 1) {
        $curS = $t.Trim()
    } elseif ($lvl -eq 2) {
        $phEntries += [PSCustomObject]@{
            Teacher = $curT
            TeacherClean = Clean-Text $curT
            Subject = $curS
            SubjectClean = Clean-Text $curS
            Course = $t.Trim()
            CourseClean = Clean-Text $t
            CourseInfo = Get-GradeAndSection $t
            Evaluados = $mInt
        }
    }
}
$wbPh.Close($false)
Write-Host "Phidias indexed: $($phEntries.Count) items"

# Map sheet names to allowed subject keywords
$sheetAllowedSubjects = @{
    "MATEMATICAS" = @("matemat", "mathem")
    "INGLES" = @("ingles", "english")
    "FRANCES" = @("franc")
    "LENGUA Y LITERATURA" = @("lengua", "literatura", "communication", "language", "comunicacion")
    "CIENCIAS" = @("ciencias", "science", "biolog", "quim", "fisic", "environmental", "project based")
    "SOCIALES" = @("social", "global", "enterprise", "ciudadan", "filosof", "investig", "project based")
    "TECNOLOGIA" = @("tecnolog", "technol")
    "EDU_FISICA" = @("fisica", "physical", "deport")
    "ARTES" = @("arte", "arts", "danza", "musica", "music", "teatro", "drama", "performing")
}

# Test matching on CIENCIAS Second (2) courses
$wbPr = $excel.Workbooks.Open($pathProgreso, 0, $true)
$sCien = $wbPr.Sheets.Item("CIENCIAS")

Write-Host "=== TEST CIENCIAS ROWS 28-34 (previously zero) ==="
for ($r = 28; $r -le 34; $r++) {
    $curso = $sCien.Cells.Item($r, 2).Text
    $doc = $sCien.Cells.Item($r, 3).Text
    $cClean = Clean-Text $curso
    $cInfo = Get-GradeAndSection $curso

    # Allowed subjects for CIENCIAS
    $allowed = $sheetAllowedSubjects["CIENCIAS"]

    $matched = $phEntries | Where-Object {
        $sub = $_.SubjectClean
        $isSubAllowed = $false
        foreach ($kw in $allowed) { if ($sub.Contains($kw)) { $isSubAllowed = $true; break } }
        
        $isCourseMatch = ($_.CourseClean -eq $cClean) -or 
                         ($cInfo.ShortCode -and $_.CourseInfo.ShortCode -eq $cInfo.ShortCode)

        $isSubAllowed -and $isCourseMatch
    }

    if ($matched) {
        Write-Host "Row $r [$curso]: MATCH FOUND! Teacher='$($matched[0].Teacher)', Subject='$($matched[0].Subject)', Evaluados=$($matched[0].Evaluados)"
    } else {
        Write-Host "Row $r [$curso]: STILL UNMATCHED"
    }
}

$wbPr.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
