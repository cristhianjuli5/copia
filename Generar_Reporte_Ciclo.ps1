param(
    [string]$TargetCycle = "24"
)

$ErrorActionPreference = "Stop"
$dir = $PSScriptRoot
if (-not $dir) { $dir = "c:\Users\analista.tic\Downloads\automatizacion" }

$pathProgreso = Join-Path $dir "PROGRESO_POR_PERIODO.xlsx"
$pathPhidias = Join-Path $dir "Archivo_Phidias.xls.xlsx"
$pathPlantilla = Join-Path $dir "PLANTILLA CA PROGRESO POR PERIODO.xlsx"
$pathSalida = Join-Path $dir "PROGRESO_POR_PERIODO_ACTUALIZADO.xlsx"

Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host "   AUTOMATIZACION DE REPORTE DE PROGRESO POR PERIODO     " -ForegroundColor Cyan
Write-Host "   Ciclo a procesar: CICLO $TargetCycle                  " -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan

function Format-Text([string]$str) {
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

function Get-ColLetter([int]$colNum) {
    $div = $colNum
    $colLetter = ""
    while ($div -gt 0) {
        $mod = ($div - 1) % 26
        $colLetter = [char](65 + $mod) + $colLetter
        $div = [int](($div - $mod) / 26)
    }
    return $colLetter
}

function Get-GradeAndSection([string]$course) {
    $c = Format-Text $course
    $grade = ""
    $letter = ""

    if ($c -match "\bpioneros\b|\bpn\b") { $grade = "pn" }
    elseif ($c -match "\bfirst garten\b|\bfg\b") { $grade = "fg" }
    elseif ($c -match "\bgarten\b|\bgr\b|^g\s") { $grade = "gr" }
    elseif ($c -match "\btransition\b|\btr\b|^t\s") { $grade = "tr" }
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

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.ScreenUpdating = $false

try {
    # ----------------------------------------------------
    # 1. PARSE ARCHIVO PHIDIAS
    # ----------------------------------------------------
    Write-Host "`n[1/3] Indexando Archivo_Phidias.xls.xlsx..." -ForegroundColor Yellow
    $wbPh = $excel.Workbooks.Open($pathPhidias, 0, $true)
    $sPh = $wbPh.Sheets.Item(1)
    $nPhRows = $sPh.UsedRange.Rows.Count

    $rangeA = $sPh.Range("A1:A$nPhRows").Value2
    $rangeM = $sPh.Range("M1:M$nPhRows").Value2

    $phidiasEntries = @()
    $curTeacher = ""
    $curSubject = ""

    for ($r = 1; $r -le $nPhRows; $r++) {
        $rawA = [string]$rangeA[$r, 1]
        if (-not $rawA) { continue }
        $lvl = $sPh.Cells.Item($r, 1).IndentLevel
        $rawM = [string]$rangeM[$r, 1]
        $intM = 0
        if ($rawM) { [int]::TryParse($rawM, [ref]$intM) | Out-Null }

        if ($lvl -eq 0) {
            $curTeacher = $rawA.Trim()
        } elseif ($lvl -eq 1) {
            $curSubject = $rawA.Trim()
        } elseif ($lvl -eq 2) {
            $curCourse = $rawA.Trim()
            $phidiasEntries += [PSCustomObject]@{
                TeacherRaw = $curTeacher
                TeacherClean = Format-Text $curTeacher
                SubjectRaw = $curSubject
                SubjectClean = Format-Text $curSubject
                CourseRaw = $curCourse
                CourseClean = Format-Text $curCourse
                CourseInfo = Get-GradeAndSection $curCourse
                Evaluados = $intM
            }
        }
    }
    $wbPh.Close($false)
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wbPh) | Out-Null
    Write-Host "   -> $($phidiasEntries.Count) cursos indexados de Phidias." -ForegroundColor Green

    # Map each sheet to allowed Phidias subject keywords
    $sheetAllowedSubjects = @{
        "MATEMATICAS" = @("matemat", "mathem", "calcul", "trigono")
        "INGLES" = @("ingles", "english")
        "FRANCES" = @("franc")
        "LENGUA Y LITERATURA" = @("lengua", "literatura", "communication", "language", "comunicacion", "filosof", "investig", "pensamiento global", "marketing", "diseño")
        "CIENCIAS" = @("ciencias", "science", "biolog", "quim", "fisic", "environmental", "project based")
        "SOCIALES" = @("social", "global", "enterprise", "ciudadan", "project based")
        "TECNOLOGIA" = @("tecnolog", "technol")
        "EDU_FISICA" = @("educacion fisica", "physical", "deport", "polimotor")
        "ARTES" = @("arte", "arts", "danza", "musica", "music", "teatro", "drama", "performing", "audiovisual", "efectos especiales")
    }

    # ----------------------------------------------------
    # 2. PARSE PLANTILLA CA PROGRESO POR PERIODO
    # ----------------------------------------------------
    Write-Host "`n[2/3] Leyendo PLANTILLA CA PROGRESO POR PERIODO..." -ForegroundColor Yellow
    $wbPl = $excel.Workbooks.Open($pathPlantilla, 0, $true)

    $plantillaExpected = @{}
    $llTable = @()

    foreach ($shPl in $wbPl.Sheets) {
        $pName = $shPl.Name
        $pClean = Format-Text $pName

        $cycleCol = 0
        $maxCols = [Math]::Min(35, $shPl.UsedRange.Columns.Count)
        for ($c = 1; $c -le $maxCols; $c++) {
            $h2 = $shPl.Cells.Item(2, $c).Text
            $h3 = $shPl.Cells.Item(3, $c).Text
            $comb = "$h2 $h3"
            
            $isMatch = $false
            if ($comb -match "\bC$TargetCycle\b" -or $comb -match "\bCiclo $TargetCycle\b" -or $comb -match "subido C$TargetCycle\b" -or $comb -match "Reporte.*C$TargetCycle\b") {
                $isMatch = $true
            }

            if ($isMatch) {
                $r3Cur = $shPl.Cells.Item(3, $c).Text
                $r4Cur = $shPl.Cells.Item(4, $c).Text
                $r3Next = $shPl.Cells.Item(3, $c + 1).Text
                $r4Next = $shPl.Cells.Item(4, $c + 1).Text

                # If current column already contains Acum (like in L&L row 3 or Ed Artistica row 4):
                if ($r3Cur -like "*Acum*" -or $r4Cur -like "*Acum*") {
                    $cycleCol = $c
                } elseif ($r4Next -like "*Acum*" -or $r3Next -like "*Acum*") {
                    $cycleCol = $c + 1
                } else {
                    $cycleCol = $c
                }
                break
            }
        }

        if ($cycleCol -gt 0) {
            Write-Host "   -> Hoja Plantilla '$pName' detectada en Columna $cycleCol." -ForegroundColor Green
            $nRows = $shPl.UsedRange.Rows.Count
            
            if ($pClean -eq "ciencias") {
                for ($r = 5; $r -le $nRows; $r++) {
                    $col1 = $shPl.Cells.Item($r, 1).Text
                    $val = $shPl.Cells.Item($r, $cycleCol).Text
                    $num = 0
                    if ($val) { [int]::TryParse($val, [ref]$num) | Out-Null }
                    
                    $c1Clean = Format-Text $col1
                    $gInfo = Get-GradeAndSection $col1
                    
                    if ($c1Clean -match "\bbio\b") {
                        $plantillaExpected["ciencias|bio|$($gInfo.Grade)"] = $num
                    } elseif ($c1Clean -match "\bem\b") {
                        $plantillaExpected["ciencias|em|$($gInfo.Grade)"] = $num
                    } elseif ($c1Clean -match "\bquim\b") {
                        $plantillaExpected["ciencias|quim|$($gInfo.Grade)"] = $num
                    } elseif ($c1Clean -match "\bfis\b|\bpreu 11\b") {
                        $plantillaExpected["ciencias|fis|$($gInfo.Grade)"] = $num
                    } else {
                        $plantillaExpected["ciencias|general|$($gInfo.Grade)"] = $num
                    }
                }
            } elseif ($pClean -eq "sociales") {
                $socSec = "general"
                for ($r = 5; $r -le $nRows; $r++) {
                    $col1 = $shPl.Cells.Item($r, 1).Text
                    $val = $shPl.Cells.Item($r, $cycleCol).Text
                    $num = 0
                    if ($val) { [int]::TryParse($val, [ref]$num) | Out-Null }
                    
                    $c1Clean = Format-Text $col1
                    if ($c1Clean.Contains("global")) { $socSec = "global" }
                    elseif ($c1Clean.Contains("enterprise")) { $socSec = "enterprise" }
                    elseif ($c1Clean.Contains("ciudadan")) { $socSec = "ciudadania" }
                    
                    $gInfo = Get-GradeAndSection $col1
                    if ($gInfo.Grade) {
                        $plantillaExpected["sociales|$socSec|$($gInfo.Grade)"] = $num
                    }
                }
            } elseif ($pClean -eq "l l") {
                $curSecLL = "ll"
                for ($r = 4; $r -le $nRows; $r++) {
                    $col1 = $shPl.Cells.Item($r, 1).Text
                    $col2 = $shPl.Cells.Item($r, 2).Text
                    $val = $shPl.Cells.Item($r, $cycleCol).Text
                    $num = 0
                    if ($val) { [int]::TryParse($val, [ref]$num) | Out-Null }
                    
                    $c1Clean = Format-Text $col1
                    if ($c1Clean.Contains("filosof")) { $curSecLL = "filosofia"; continue }
                    if ($c1Clean.Contains("investig")) { $curSecLL = "investigacion"; continue }
                    if ($c1Clean.Contains("comunicacion")) { $curSecLL = "comunicacion"; continue }
                    if ($c1Clean.Contains("pensamiento global")) { $curSecLL = "pensamiento_global"; continue }
                    if ($c1Clean.Contains("marketing") -or $c1Clean.Contains("diseno")) { $curSecLL = "diseno"; continue }
                    
                    $llTable += [PSCustomObject]@{
                        Row = $r
                        Section = $curSecLL
                        GradePattern = $c1Clean
                        GradeInfo = Get-GradeAndSection $col1
                        DocentePattern = Format-Text $col2
                        Expected = $num
                    }
                }
            } elseif ($pClean -eq "ed artistica") {
                for ($r = 5; $r -le $nRows; $r++) {
                    $col1 = $shPl.Cells.Item($r, 1).Text
                    $val = $shPl.Cells.Item($r, $cycleCol).Text
                    $num = 0
                    if ($val) { [int]::TryParse($val, [ref]$num) | Out-Null }
                    $gInfo = Get-GradeAndSection $col1
                    if ($gInfo.Grade -and -not $plantillaExpected.ContainsKey("artes|$($gInfo.Grade)")) {
                        $plantillaExpected["artes|$($gInfo.Grade)"] = $num
                    }
                }
            } else {
                # Standard sheets: math, eng, frances, tecnologia, pe
                for ($r = 4; $r -le $nRows; $r++) {
                    $col1 = $shPl.Cells.Item($r, 1).Text
                    $val = $shPl.Cells.Item($r, $cycleCol).Text
                    $num = 0
                    if ($val) { [int]::TryParse($val, [ref]$num) | Out-Null }
                    $gInfo = Get-GradeAndSection $col1
                    if ($gInfo.Grade) {
                        $plantillaExpected["$pClean|$($gInfo.Grade)"] = $num
                    }
                }
            }
        }
    }
    $wbPl.Close($false)
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wbPl) | Out-Null

    # ----------------------------------------------------
    # 3. ACTUALIZAR PROGRESO_POR_PERIODO
    # ----------------------------------------------------
    Write-Host "`n[3/3] Actualizando PROGRESO_POR_PERIODO con reglas inteligentes completas..." -ForegroundColor Yellow
    Copy-Item -Path $pathProgreso -Destination $pathSalida -Force
    $wbPr = $excel.Workbooks.Open($pathSalida)

    $progresoSheets = @("MATEMATICAS", "INGLES", "FRANCES", "LENGUA Y LITERATURA", "CIENCIAS", "SOCIALES", "TECNOLOGIA", "EDU_FISICA", "ARTES")

    $sheetToPlantillaMap = @{
        "MATEMATICAS" = "math"
        "INGLES" = "eng"
        "FRANCES" = "frances"
        "LENGUA Y LITERATURA" = "l l"
        "CIENCIAS" = "ciencias"
        "SOCIALES" = "sociales"
        "TECNOLOGIA" = "tecnologia"
        "EDU_FISICA" = "pe"
        "ARTES" = "ed artistica"
    }

    foreach ($sheetName in $progresoSheets) {
        $sPr = $wbPr.Sheets.Item($sheetName)
        $lastCol = $sPr.UsedRange.Columns.Count
        $lastCycleCol = 0
        for ($c = $lastCol; $c -ge 6; $c--) {
            if ($sPr.Cells.Item(1, $c).Text -like "*CICLO*") {
                $lastCycleCol = $c
                break
            }
        }
        if ($lastCycleCol -eq 0) { $lastCycleCol = 36 }

        $newStartCol = $lastCycleCol + 5
        $nRows = $sPr.UsedRange.Rows.Count
        Write-Host "   -> Procesando '$sheetName'..."

        # Headers
        $sPr.Cells.Item(1, $newStartCol).Value2 = "CICLO $TargetCycle"
        $sPr.Cells.Item(3, $newStartCol).Value2 = "ESPERADO"
        $sPr.Cells.Item(3, $newStartCol + 1).Value2 = "TOTAL"
        $sPr.Cells.Item(3, $newStartCol + 2).Value2 = "PHIDIAS"
        $sPr.Cells.Item(3, $newStartCol + 3).Value2 = "PORCENTAJE"

        # Bulk copy formats
        $sPr.Range($sPr.Cells.Item(1, $lastCycleCol), $sPr.Cells.Item(3, $lastCycleCol + 3)).Copy() | Out-Null
        $sPr.Range($sPr.Cells.Item(1, $newStartCol), $sPr.Cells.Item(3, $newStartCol + 3)).PasteSpecial(-4122) | Out-Null
        $sPr.Range($sPr.Cells.Item(6, $lastCycleCol), $sPr.Cells.Item($nRows, $lastCycleCol + 3)).Copy() | Out-Null
        $sPr.Range($sPr.Cells.Item(6, $newStartCol), $sPr.Cells.Item($nRows, $newStartCol + 3)).PasteSpecial(-4122) | Out-Null

        $pSheetClean = $sheetToPlantillaMap[$sheetName]
        $allowedKeywords = $sheetAllowedSubjects[$sheetName]

        $colLetEsp = Get-ColLetter $newStartCol
        $colLetTot = Get-ColLetter ($newStartCol + 1)
        $colLetPhi = Get-ColLetter ($newStartCol + 2)

        $curSecPr = "general"
        for ($r = 6; $r -le $nRows; $r++) {
            $col1 = $sPr.Cells.Item($r, 1).Text
            $curso = $sPr.Cells.Item($r, 2).Text
            $docente = $sPr.Cells.Item($r, 3).Text

            # Update section BEFORE filtering
            $col1Clean = Format-Text $col1
            if ($sheetName -eq "CIENCIAS") {
                if ($col1Clean.Contains("biologia")) { $curSecPr = "bio" }
                elseif ($col1Clean.Contains("quimica")) { $curSecPr = "quim" }
                elseif ($col1Clean.Contains("fisica")) { $curSecPr = "fis" }
                elseif ($col1Clean.Contains("environmental")) { $curSecPr = "em" }
            } elseif ($sheetName -eq "SOCIALES") {
                if ($col1Clean.Contains("global")) { $curSecPr = "global" }
                elseif ($col1Clean.Contains("enterprise")) { $curSecPr = "enterprise" }
                elseif ($col1Clean.Contains("ciudadan")) { $curSecPr = "ciudadania" }
            } elseif ($sheetName -eq "LENGUA Y LITERATURA") {
                if ($col1Clean.Contains("filosof")) { $curSecPr = "filosofia" }
                elseif ($col1Clean.Contains("comunicacion")) { $curSecPr = "comunicacion" }
            } elseif ($sheetName -eq "ARTES") {
                if ($col1Clean.Contains("musica")) { $curSecPr = "musica" }
                elseif ($col1Clean.Contains("teatro")) { $curSecPr = "teatro" }
                elseif ($col1Clean.Contains("audiovisual")) { $curSecPr = "audiovisual" }
                elseif ($col1Clean.Contains("danza")) { $curSecPr = "danza" }
                elseif ($col1Clean.Contains("arte y diseno") -or $col1Clean.Contains("arte")) { $curSecPr = "arte" }
            }

            if (-not $curso -and -not $docente) { continue }
            if ($curso -match "CURSOS" -or $docente -match "DOCENTE") { continue }
            if ($curso -match "TOTALES GENERALES") { continue }

            $cCurso = Format-Text $curso
            $cDocente = Format-Text $docente
            $cInfo = Get-GradeAndSection $curso

            # -----------------------------------------------
            # A. VALOR ESPERADO (PLANTILLA)
            # -----------------------------------------------
            $expVal = 0

            if ($sheetName -eq "CIENCIAS") {
                $key = "ciencias|$curSecPr|$($cInfo.Grade)"
                if ($plantillaExpected.ContainsKey($key)) {
                    $expVal = $plantillaExpected[$key]
                } elseif ($plantillaExpected.ContainsKey("ciencias|general|$($cInfo.Grade)")) {
                    $expVal = $plantillaExpected["ciencias|general|$($cInfo.Grade)"]
                }
            } elseif ($sheetName -eq "SOCIALES") {
                if ($cInfo.Grade -in @("pn", "fg", "gr", "tr")) {
                    if ($plantillaExpected.ContainsKey("ciencias|general|$($cInfo.Grade)")) {
                        $expVal = $plantillaExpected["ciencias|general|$($cInfo.Grade)"]
                    }
                } elseif ($curSecPr -eq "ciudadania" -and $cInfo.Grade -eq "11") {
                    $expVal = 0
                } else {
                    $key = "sociales|$curSecPr|$($cInfo.Grade)"
                    if ($plantillaExpected.ContainsKey($key)) {
                        $expVal = $plantillaExpected[$key]
                    } elseif ($plantillaExpected.ContainsKey("sociales|general|$($cInfo.Grade)")) {
                        $expVal = $plantillaExpected["sociales|general|$($cInfo.Grade)"]
                    }
                }
            } elseif ($sheetName -eq "LENGUA Y LITERATURA") {
                if ($curSecPr -eq "filosofia") {
                    if ($cInfo.Grade -eq "10") {
                        $row = $llTable | Where-Object { $_.Section -eq "filosofia" -and $_.GradeInfo.Grade -eq "10" }
                        if ($row) { $expVal = $row[0].Expected }
                    } elseif ($cInfo.Grade -eq "11") {
                        $row = $llTable | Where-Object { $_.Section -eq "investigacion" -and $_.GradeInfo.Grade -eq "11" }
                        if ($row) { $expVal = $row[0].Expected }
                    }
                } elseif ($curSecPr -eq "comunicacion") {
                    $docWords = $cDocente.Split(' ') | Where-Object { $_.Length -ge 4 }
                    foreach ($rowLL in ($llTable | Where-Object { $_.Section -in @("comunicacion", "pensamiento_global", "diseno") })) {
                        $matchW = 0
                        foreach ($w in $docWords) {
                            if ($rowLL.DocentePattern.Contains($w)) { $matchW++ }
                        }
                        if ($matchW -ge 1) {
                            $expVal = $rowLL.Expected
                            break
                        }
                    }
                    if ($expVal -eq 0) {
                        $row = $llTable | Where-Object { $_.Section -eq "comunicacion" -and $_.GradeInfo.Grade -eq "11" }
                        if ($row) { $expVal = $row[0].Expected }
                    }
                } else {
                    $matchedRow = $null
                    if ($cInfo.Grade -and $cInfo.Letter) {
                        $candidates = $llTable | Where-Object { 
                            $_.Section -eq "ll" -and 
                            $_.GradeInfo.Grade -eq $cInfo.Grade -and 
                            $_.GradePattern.Contains($cInfo.Letter) 
                        }
                        if ($candidates) { $matchedRow = $candidates[0] }
                    }
                    if (-not $matchedRow -and $cDocente) {
                        $docWords = $cDocente.Split(' ') | Where-Object { $_.Length -ge 4 }
                        foreach ($rowLL in ($llTable | Where-Object { $_.Section -eq "ll" })) {
                            if ($cInfo.Grade -and $rowLL.GradeInfo.Grade -ne $cInfo.Grade) { continue }
                            $mW = 0
                            foreach ($w in $docWords) {
                                if ($rowLL.DocentePattern.Contains($w)) { $mW++ }
                            }
                            if ($mW -ge 1) {
                                $matchedRow = $rowLL
                                break
                            }
                        }
                    }
                    if (-not $matchedRow -and $cInfo.Grade) {
                        $candidates = $llTable | Where-Object { $_.Section -eq "ll" -and $_.GradeInfo.Grade -eq $cInfo.Grade }
                        if ($candidates) { $matchedRow = $candidates[0] }
                    }
                    if ($matchedRow) { $expVal = $matchedRow.Expected }
                }
            } elseif ($sheetName -eq "ARTES") {
                if ($plantillaExpected.ContainsKey("artes|$($cInfo.Grade)")) {
                    $expVal = $plantillaExpected["artes|$($cInfo.Grade)"]
                } else {
                    $expVal = 4
                }
            } else {
                if ($cInfo.Grade -and $plantillaExpected.ContainsKey("$pSheetClean|$($cInfo.Grade)")) {
                    $expVal = $plantillaExpected["$pSheetClean|$($cInfo.Grade)"]
                }
            }

            # -----------------------------------------------
            # B. NOTAS SUBIDAS EN PHIDIAS
            # -----------------------------------------------
            $phiVal = 0
            $foundPhi = $false

            # Specific keyword narrowing by section:
            $allowedKeywords = $sheetAllowedSubjects[$sheetName]
            if ($sheetName -eq "CIENCIAS") {
                if ($curSecPr -eq "bio") { $allowedKeywords = @("biolog") }
                elseif ($curSecPr -eq "quim") { $allowedKeywords = @("quim") }
                elseif ($curSecPr -eq "fis") { $allowedKeywords = @("fisic") }
                elseif ($curSecPr -eq "em") { $allowedKeywords = @("environmental") }
                else { $allowedKeywords = @("ciencias", "science", "project based") }
            } elseif ($sheetName -eq "SOCIALES") {
                if ($curSecPr -eq "global") { $allowedKeywords = @("global") }
                elseif ($curSecPr -eq "enterprise") { $allowedKeywords = @("enterprise") }
                elseif ($curSecPr -eq "ciudadania") { $allowedKeywords = @("ciudadan") }
                else { $allowedKeywords = @("social", "project based") }
            } elseif ($sheetName -eq "LENGUA Y LITERATURA") {
                if ($curSecPr -eq "filosofia") { $allowedKeywords = @("filosof", "investig") }
                elseif ($curSecPr -eq "comunicacion") { $allowedKeywords = @("comunicacion", "communication", "pensamiento global", "marketing", "diseno", "diseño") }
                else { $allowedKeywords = @("lengua", "literatura", "language") }
            } elseif ($sheetName -eq "ARTES") {
                if ($curSecPr -eq "musica") { $allowedKeywords = @("musica", "music") }
                elseif ($curSecPr -eq "teatro") { $allowedKeywords = @("teatro", "drama") }
                elseif ($curSecPr -eq "audiovisual") { $allowedKeywords = @("audiovisual", "efectos especiales") }
                elseif ($curSecPr -eq "danza") { $allowedKeywords = @("danza") }
                elseif ($curSecPr -eq "arte") { $allowedKeywords = @("arte", "arts") }
                else { $allowedKeywords = @("performing", "arte", "arts", "danza", "musica") }
            }

            # Filter Phidias entries to the allowed subjects for this sheet
            $candidates = $phidiasEntries | Where-Object {
                $sub = $_.SubjectClean
                $ok = $false
                foreach ($kw in $allowedKeywords) {
                    if ($sub.Contains($kw)) { $ok = $true; break }
                }
                $ok
            }

            # 1. Match Docente + Curso exacto
            if ($cDocente) {
                $m = $candidates | Where-Object {
                    $isCourseMatch = ($_.CourseClean -eq $cCurso) -or 
                                     ($cInfo.ShortCode -and $_.CourseInfo.ShortCode -eq $cInfo.ShortCode)
                    $isTeacherMatch = ($_.TeacherClean -eq $cDocente) -or ($_.TeacherClean.Contains($cDocente)) -or ($cDocente.Contains($_.TeacherClean))
                    $isCourseMatch -and $isTeacherMatch
                }
                if ($m) {
                    $phiVal = $m[0].Evaluados
                    $foundPhi = $true
                }
            }

            # 2. Match Docente (apellidos) + Curso
            if (-not $foundPhi -and $cDocente) {
                $docWords = $cDocente.Split(' ') | Where-Object { $_.Length -ge 4 }
                foreach ($entry in $candidates) {
                    $isCourseMatch = ($entry.CourseClean -eq $cCurso) -or 
                                     ($cInfo.ShortCode -and $entry.CourseInfo.ShortCode -eq $cInfo.ShortCode)
                    if ($isCourseMatch) {
                        $mW = 0
                        foreach ($w in $docWords) {
                            if ($entry.TeacherClean.Contains($w)) { $mW++ }
                        }
                        if ($mW -ge 1) {
                            $phiVal = $entry.Evaluados
                            $foundPhi = $true
                            break
                        }
                    }
                }
            }

            # 3. Match Curso + Materia (cuando el docente cambió o la celda estaba vacía)
            if (-not $foundPhi) {
                $m = $candidates | Where-Object {
                    ($_.CourseClean -eq $cCurso) -or 
                    ($cInfo.ShortCode -and $_.CourseInfo.ShortCode -eq $cInfo.ShortCode)
                }
                if ($m) {
                    $phiVal = $m[0].Evaluados
                    $foundPhi = $true
                }
            }

            # 4. Fallback especial para cursos consolidados por grado (ej. Artes Fifth (5))
            if (-not $foundPhi -and $cDocente) {
                $docWords = $cDocente.Split(' ') | Where-Object { $_.Length -ge 4 }
                foreach ($entry in $candidates) {
                    if ($cInfo.Grade -and $entry.CourseInfo.Grade -eq $cInfo.Grade) {
                        $mW = 0
                        foreach ($w in $docWords) {
                            if ($entry.TeacherClean.Contains($w)) { $mW++ }
                        }
                        if ($mW -ge 1) {
                            $phiVal = $entry.Evaluados
                            $foundPhi = $true
                            break
                        }
                    }
                }
            }

            # Escribir en las celdas
            $sPr.Cells.Item($r, $newStartCol).Value2 = $expVal
            $sPr.Cells.Item($r, $newStartCol + 1).Formula = "=$colLetEsp$r*`$D$r"
            $sPr.Cells.Item($r, $newStartCol + 2).Value2 = $phiVal
            $sPr.Cells.Item($r, $newStartCol + 3).Formula = "=IFERROR($colLetPhi$r/$colLetTot$r, 0)"
            $sPr.Cells.Item($r, $newStartCol + 3).NumberFormat = "0%"
        }

        # Formato condicional
        $pctRange = $sPr.Range($sPr.Cells.Item(6, $newStartCol + 3), $sPr.Cells.Item($nRows, $newStartCol + 3))
        $pctRange.FormatConditions.Delete()
        $fcRed = $pctRange.FormatConditions.Add(1, 1, "0", "0,899999")
        $fcRed.Interior.Color = 13421823
        $fcGreen = $pctRange.FormatConditions.Add(1, 1, "0,9", "1")
        $fcGreen.Interior.Color = 13565902
        $fcYellow = $pctRange.FormatConditions.Add(1, 1, "1,000001", "10")
        $fcYellow.Interior.Color = 10092543
    }

    $wbPr.Save()
    $wbPr.Close($false)
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wbPr) | Out-Null

    Write-Host "`n==========================================================" -ForegroundColor Green
    Write-Host " [EXITO] REPORTE ACTUALIZADO CON COINCIDENCIAS EXACTAS!   " -ForegroundColor Green
    Write-Host " Archivo generado: $pathSalida" -ForegroundColor Green
    Write-Host "==========================================================" -ForegroundColor Green
} catch {
    Write-Error "Error: $_"
} finally {
    $excel.ScreenUpdating = $true
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}
