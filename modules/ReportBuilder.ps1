<#
.SYNOPSIS
    Módulo para generar y actualizar el reporte de Progreso por Periodo.
#>

function Update-ProgresoReport {
    <#
    .SYNOPSIS
        Construye el nuevo ciclo en el libro de Progreso por Periodo combinando Phidias y la Plantilla.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [object]$ExcelApp,

        [Parameter(Mandatory = $true)]
        [string]$PathProgreso,

        [Parameter(Mandatory = $true)]
        [string]$PathSalida,

        [Parameter(Mandatory = $true)]
        [string]$TargetCycle,

        [Parameter(Mandatory = $true)]
        [array]$PhidiasEntries,

        [Parameter(Mandatory = $true)]
        [hashtable]$PlantillaExpected,

        [Parameter(Mandatory = $true)]
        [array]$LLTable
    )

    if (-not (Test-Path $PathProgreso)) {
        throw "No se encontró el archivo base de Progreso en: $PathProgreso"
    }

    # Copiar libro base a la ruta de salida (si está abierto en otra instancia de Excel, intentar cerrarlo primero)
    $leafSalida = Split-Path -Leaf $PathSalida
    try {
        Copy-Item -Path $PathProgreso -Destination $PathSalida -Force
    } catch {
        try {
            $activeExcel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
            foreach ($wbOpen in $activeExcel.Workbooks) {
                if ($wbOpen.Name -eq $leafSalida) {
                    $wbOpen.Close($false)
                    break
                }
            }
        } catch {}
        Start-Sleep -Milliseconds 500
        Copy-Item -Path $PathProgreso -Destination $PathSalida -Force
    }

    $wbPr = $ExcelApp.Workbooks.Open($PathSalida)
    try {
        $progresoSheets = @("MATEMATICAS", "INGLES", "FRANCES", "LENGUA Y LITERATURA", "CIENCIAS", "SOCIALES", "TECNOLOGIA", "EDU_FISICA", "ARTES")

        $sheetToPlantillaMap = @{
            "MATEMATICAS"         = "math"
            "INGLES"              = "eng"
            "FRANCES"             = "frances"
            "LENGUA Y LITERATURA" = "l l"
            "CIENCIAS"            = "ciencias"
            "SOCIALES"            = "sociales"
            "TECNOLOGIA"          = "tecnologia"
            "EDU_FISICA"          = "pe"
            "ARTES"               = "ed artistica"
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
            Write-Host "   -> Procesando '$sheetName'..." -ForegroundColor Cyan

            # Encabezados
            $sPr.Cells.Item(1, $newStartCol).Value2 = "CICLO $TargetCycle"
            $sPr.Cells.Item(3, $newStartCol).Value2 = "ESPERADO"
            $sPr.Cells.Item(3, $newStartCol + 1).Value2 = "TOTAL"
            $sPr.Cells.Item(3, $newStartCol + 2).Value2 = "PHIDIAS"
            $sPr.Cells.Item(3, $newStartCol + 3).Value2 = "PORCENTAJE"

            # Copia en bloque de formatos del ciclo previo.
            # PasteSpecial puede fallar si el SO descarta el portapapeles entre
            # Copy() y PasteSpecial() (condición de carrera normal en Windows).
            # En ese caso se omite el copiado de formato y el script continúa.
            try {
                $sPr.Range($sPr.Cells.Item(1, $lastCycleCol), $sPr.Cells.Item(3, $lastCycleCol + 3)).Copy() | Out-Null
                $sPr.Range($sPr.Cells.Item(1, $newStartCol), $sPr.Cells.Item(3, $newStartCol + 3)).PasteSpecial(-4122) | Out-Null
                $sPr.Range($sPr.Cells.Item(6, $lastCycleCol), $sPr.Cells.Item($nRows, $lastCycleCol + 3)).Copy() | Out-Null
                $sPr.Range($sPr.Cells.Item(6, $newStartCol), $sPr.Cells.Item($nRows, $newStartCol + 3)).PasteSpecial(-4122) | Out-Null
            } catch {
                Write-Warning "[$sheetName] Formato del ciclo previo no copiado (portapapeles no disponible). Se continua."
            }

            $pSheetClean = $sheetToPlantillaMap[$sheetName]

            $colLetEsp = Get-ColLetter $newStartCol
            $colLetTot = Get-ColLetter ($newStartCol + 1)
            $colLetPhi = Get-ColLetter ($newStartCol + 2)

            $curSecPr = "general"
            for ($r = 6; $r -le $nRows; $r++) {
                $col1 = $sPr.Cells.Item($r, 1).Text
                $curso = $sPr.Cells.Item($r, 2).Text
                $docente = $sPr.Cells.Item($r, 3).Text

                # Detección contextual de la sección según columna 1
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
                    elseif ($col1Clean.Contains("audiovisual") -or $col1Clean.Contains("audio visual")) { $curSecPr = "audiovisual" }
                    elseif ($col1Clean.Contains("danza")) { $curSecPr = "danza" }
                    elseif ($col1Clean.Contains("arte y diseno") -or $col1Clean.Contains("arte") -or ($col1Clean -eq "primaria" -and $curSecPr -eq "general")) { $curSecPr = "arte" }
                }

                # Omitir filas vacías, cabeceras o subtotales
                if (-not $curso -and -not $docente) { continue }
                if ($curso -match "CURSOS" -or $docente -match "DOCENTE") { continue }
                if ($curso -match "TOTALES GENERALES") { continue }

                $cCurso = Format-Text $curso
                $cDocente = Format-Text $docente
                $cInfo = Get-GradeAndSection $curso

                # 1. Obtener Valor Esperado de la Plantilla
                $expVal = Get-ExpectedValue `
                    -SheetName $sheetName `
                    -CurSecPr $curSecPr `
                    -CourseInfo $cInfo `
                    -DocenteClean $cDocente `
                    -PlantillaExpected $PlantillaExpected `
                    -LLTable $LLTable `
                    -PSheetClean $pSheetClean

                # 2. Obtener Calificaciones Subidas en Phidias
                $allowedKeywords = Get-PhidiasSubjectKeywords -SheetName $sheetName -CurSecPr $curSecPr
                $phiVal = Find-PhidiasEvaluatedCount `
                    -PhidiasEntries $PhidiasEntries `
                    -AllowedKeywords $allowedKeywords `
                    -DocenteClean $cDocente `
                    -CursoClean $cCurso `
                    -CourseInfo $cInfo

                # 3. Asignación a celdas
                $sPr.Cells.Item($r, $newStartCol).Value2 = $expVal
                $sPr.Cells.Item($r, $newStartCol + 1).Formula = "=$colLetEsp$r*`$D$r"
                $sPr.Cells.Item($r, $newStartCol + 2).Value2 = $phiVal
                $sPr.Cells.Item($r, $newStartCol + 3).Formula = "=IFERROR($colLetPhi$r/$colLetTot$r, 0)"
                $sPr.Cells.Item($r, $newStartCol + 3).NumberFormat = "0%"
            }

            # Formato condicional sobre la columna de porcentajes
            $pctRange = $sPr.Range($sPr.Cells.Item(6, $newStartCol + 3), $sPr.Cells.Item($nRows, $newStartCol + 3))
            $pctRange.FormatConditions.Delete()
            
            # Rojo: 0% a 89.99%
            $fcRed = $pctRange.FormatConditions.Add(1, 1, "0", "0,899999")
            $fcRed.Interior.Color = 13421823
            
            # Verde: 90% a 100%
            $fcGreen = $pctRange.FormatConditions.Add(1, 1, "0,9", "1")
            $fcGreen.Interior.Color = 13565902
            
            # Amarillo: > 100% (superó la meta)
            $fcYellow = $pctRange.FormatConditions.Add(1, 1, "1,000001", "10")
            $fcYellow.Interior.Color = 10092543
        }

        $wbPr.Save()
        $wbPr.Close($false)
    } finally {
        if ($wbPr) {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wbPr) | Out-Null
        }
    }
}
