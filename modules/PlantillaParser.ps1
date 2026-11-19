<#
.SYNOPSIS
    Módulo para parsear la plantilla esperada de progreso y calcular el valor esperado por curso/docente.
#>

function Get-PlantillaData {
    <#
    .SYNOPSIS
        Lee el archivo de Plantilla de Progreso y extrae los valores esperados correspondientes al Ciclo objetivo.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [object]$ExcelApp,

        [Parameter(Mandatory = $true)]
        [string]$PathPlantilla,

        [Parameter(Mandatory = $true)]
        [string]$TargetCycle
    )

    if (-not (Test-Path $PathPlantilla)) {
        throw "No se encontró el archivo de Plantilla en: $PathPlantilla"
    }

    $wbPl = $ExcelApp.Workbooks.Open($PathPlantilla, 0, $true)
    try {
        $plantillaExpected = @{}
        $llTable = @()

        foreach ($shPl in $wbPl.Sheets) {
            $pName = $shPl.Name
            $pClean = Format-Text $pName

            $cycleCol = 0
            $maxCols = [Math]::Min(35, $shPl.UsedRange.Columns.Count)

            # Para el cierre de periodo (Ciclo 28 / FINAL / TOTAL CIERRE), priorizar la última columna de cada tabla ("FINAL" o "TOTAL CIERRE")
            if ($TargetCycle -match "^(28|final|cierre|total)" ) {
                for ($c = $maxCols; $c -ge 1; $c--) {
                    $h2 = $shPl.Cells.Item(2, $c).Text
                    $h3 = $shPl.Cells.Item(3, $c).Text
                    $comb = "$h2 $h3"

                    if ($comb -match "\bFINAL\b" -or $comb -match "\bTOTAL CIERRE\b") {
                        $r3Cur = $shPl.Cells.Item(3, $c).Text
                        $r4Cur = $shPl.Cells.Item(4, $c).Text
                        $r3Next = $shPl.Cells.Item(3, $c + 1).Text
                        $r4Next = $shPl.Cells.Item(4, $c + 1).Text

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
            }

            if ($cycleCol -eq 0) {
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
            }

            if ($cycleCol -gt 0) {
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
                            Row            = $r
                            Section        = $curSecLL
                            GradePattern   = $c1Clean
                            GradeInfo      = Get-GradeAndSection $col1
                            DocentePattern = Format-Text $col2
                            Expected       = $num
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
                    # Hojas estándar: math, eng, frances, tecnologia, pe
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

        return @{
            PlantillaExpected = $plantillaExpected
            LLTable           = $llTable
        }
    } finally {
        $wbPl.Close($false)
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wbPl) | Out-Null
    }
}

function Get-ExpectedValue {
    <#
    .SYNOPSIS
        Calcula el valor esperado para una fila específica del reporte.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$SheetName,

        [Parameter(Mandatory = $false)]
        [string]$CurSecPr = "general",

        [Parameter(Mandatory = $true)]
        [hashtable]$CourseInfo,

        [Parameter(Mandatory = $false)]
        [string]$DocenteClean = "",

        [Parameter(Mandatory = $true)]
        [hashtable]$PlantillaExpected,

        [Parameter(Mandatory = $true)]
        [array]$LLTable,

        [Parameter(Mandatory = $false)]
        [string]$PSheetClean = ""
    )

    $expVal = 0

    if ($SheetName -eq "CIENCIAS") {
        $key = "ciencias|$CurSecPr|$($CourseInfo.Grade)"
        if ($PlantillaExpected.ContainsKey($key)) {
            $expVal = $PlantillaExpected[$key]
        } elseif ($PlantillaExpected.ContainsKey("ciencias|general|$($CourseInfo.Grade)")) {
            $expVal = $PlantillaExpected["ciencias|general|$($CourseInfo.Grade)"]
        }
    } elseif ($SheetName -eq "SOCIALES") {
        if ($CourseInfo.Grade -in @("pn", "fg", "gr", "tr")) {
            if ($PlantillaExpected.ContainsKey("ciencias|general|$($CourseInfo.Grade)")) {
                $expVal = $PlantillaExpected["ciencias|general|$($CourseInfo.Grade)"]
            }
        } elseif ($CurSecPr -eq "ciudadania" -and $CourseInfo.Grade -eq "11") {
            $expVal = 0
        } else {
            $key = "sociales|$CurSecPr|$($CourseInfo.Grade)"
            if ($PlantillaExpected.ContainsKey($key)) {
                $expVal = $PlantillaExpected[$key]
            } elseif ($PlantillaExpected.ContainsKey("sociales|general|$($CourseInfo.Grade)")) {
                $expVal = $PlantillaExpected["sociales|general|$($CourseInfo.Grade)"]
            }
        }
    } elseif ($SheetName -eq "LENGUA Y LITERATURA") {
        if ($CurSecPr -eq "filosofia") {
            if ($CourseInfo.Grade -eq "10") {
                $row = $LLTable | Where-Object { $_.Section -eq "filosofia" -and $_.GradeInfo.Grade -eq "10" }
                if ($row) { $expVal = $row[0].Expected }
            } elseif ($CourseInfo.Grade -eq "11") {
                $row = $LLTable | Where-Object { $_.Section -eq "investigacion" -and $_.GradeInfo.Grade -eq "11" }
                if ($row) { $expVal = $row[0].Expected }
            }
        } elseif ($CurSecPr -eq "comunicacion") {
            $docWords = $DocenteClean.Split(' ') | Where-Object { $_.Length -ge 4 }
            foreach ($rowLL in ($LLTable | Where-Object { $_.Section -in @("comunicacion", "pensamiento_global", "diseno") })) {
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
                $row = $LLTable | Where-Object { $_.Section -eq "comunicacion" -and $_.GradeInfo.Grade -eq "11" }
                if ($row) { $expVal = $row[0].Expected }
            }
        } else {
            $matchedRow = $null
            if ($CourseInfo.Grade -and $CourseInfo.Letter) {
                $candidates = $LLTable | Where-Object { 
                    $_.Section -eq "ll" -and 
                    $_.GradeInfo.Grade -eq $CourseInfo.Grade -and 
                    $_.GradePattern.Contains($CourseInfo.Letter) 
                }
                if ($candidates) { $matchedRow = $candidates[0] }
            }
            if (-not $matchedRow -and $DocenteClean) {
                $docWords = $DocenteClean.Split(' ') | Where-Object { $_.Length -ge 4 }
                foreach ($rowLL in ($LLTable | Where-Object { $_.Section -eq "ll" })) {
                    if ($CourseInfo.Grade -and $rowLL.GradeInfo.Grade -ne $CourseInfo.Grade) { continue }
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
            if (-not $matchedRow -and $CourseInfo.Grade) {
                $candidates = $LLTable | Where-Object { $_.Section -eq "ll" -and $_.GradeInfo.Grade -eq $CourseInfo.Grade }
                if ($candidates) { $matchedRow = $candidates[0] }
            }
            if ($matchedRow) { $expVal = $matchedRow.Expected }
        }
    } elseif ($SheetName -eq "ARTES") {
        if ($PlantillaExpected.ContainsKey("artes|$($CourseInfo.Grade)")) {
            $expVal = $PlantillaExpected["artes|$($CourseInfo.Grade)"]
        } else {
            $expVal = 6
        }
    } else {
        if ($CourseInfo.Grade -and $PlantillaExpected.ContainsKey("$PSheetClean|$($CourseInfo.Grade)")) {
            $expVal = $PlantillaExpected["$PSheetClean|$($CourseInfo.Grade)"]
        }
    }

    return $expVal
}
