<#
.SYNOPSIS
    Módulo para indexar y consultar información de cursos y docentes desde Archivo_Phidias.
#>

# Mapeo base de asignaturas permitidas por hoja en Phidias
$global:SheetAllowedSubjects = @{
    "MATEMATICAS"         = @("matemat", "mathem", "calcul", "trigono")
    "INGLES"              = @("ingles", "english", "communication and language")
    "FRANCES"             = @("franc")
    "LENGUA Y LITERATURA" = @("lengua", "literatura", "comunicacion", "filosof", "investig", "pensamiento global", "marketing", "diseño")
    "CIENCIAS"            = @("ciencias", "science", "biolog", "quim", "fisic", "environmental", "project based")
    "SOCIALES"            = @("sociales", "global", "enterprise", "ciudadan", "project based")
    "TECNOLOGIA"          = @("tecnolog", "technol")
    "EDU_FISICA"          = @("educacion fisica", "physical", "deport", "polimotor")
    "ARTES"               = @("arte", "arts", "danza", "musica", "music", "teatro", "drama", "performing", "audiovisual", "efectos especiales")
}

function Get-PhidiasEntries {
    <#
    .SYNOPSIS
        Lee el archivo de Phidias e indexa todas las entradas jerárquicas (Docente -> Asignatura -> Curso).
    #>
    param(
        [Parameter(Mandatory = $true)]
        [object]$ExcelApp,

        [Parameter(Mandatory = $true)]
        [string]$PathPhidias
    )

    if (-not (Test-Path $PathPhidias)) {
        throw "No se encontró el archivo de Phidias en: $PathPhidias"
    }

    $wbPh = $ExcelApp.Workbooks.Open($PathPhidias, 0, $true)
    try {
        $sPh = $wbPh.Sheets.Item(1)
        $nPhRows = $sPh.UsedRange.Rows.Count

        $rangeA = $sPh.Range("A1:A$nPhRows").Value2
        $rangeM = $sPh.Range("M1:M$nPhRows").Value2

        $entries = @()
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
                $subClean = Format-Text $curSubject
                if ($subClean -match "^club\b|^extracurricular\b") { continue }

                $subInfo = Get-GradeAndSection $curSubject
                $crsInfo = Get-GradeAndSection $curCourse
                if ($subInfo.Grade -and $crsInfo.Grade -ne $subInfo.Grade) {
                    $crsInfo = @{
                        Grade     = $subInfo.Grade
                        Letter    = ""
                        ShortCode = $subInfo.Grade
                    }
                }

                $entries += [PSCustomObject]@{
                    TeacherRaw   = $curTeacher
                    TeacherClean = Format-Text $curTeacher
                    SubjectRaw   = $curSubject
                    SubjectClean = $subClean
                    CourseRaw    = $curCourse
                    CourseClean  = Format-Text $curCourse
                    CourseInfo   = $crsInfo
                    Evaluados    = $intM
                }
            }
        }
        return ,$entries
    } finally {
        $wbPh.Close($false)
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wbPh) | Out-Null
    }
}

function Get-PhidiasSubjectKeywords {
    <#
    .SYNOPSIS
        Obtiene las palabras clave permitidas según la hoja y la sub-sección temática.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$SheetName,

        [Parameter(Mandatory = $false)]
        [string]$CurSecPr = "general"
    )

    if ($SheetName -eq "CIENCIAS") {
        if ($CurSecPr -eq "bio") { return @("biolog") }
        if ($CurSecPr -eq "quim") { return @("quim") }
        if ($CurSecPr -eq "fis") { return @("fisic") }
        if ($CurSecPr -eq "em") { return @("environmental") }
        return @("ciencias", "science", "project based")
    }

    if ($SheetName -eq "SOCIALES") {
        if ($CurSecPr -eq "global") { return @("global") }
        if ($CurSecPr -eq "enterprise") { return @("enterprise") }
        if ($CurSecPr -eq "ciudadania") { return @("ciudadan") }
        return @("sociales", "project based")
    }

    if ($SheetName -eq "LENGUA Y LITERATURA") {
        if ($CurSecPr -eq "filosofia") { return @("filosof", "investig") }
        if ($CurSecPr -eq "comunicacion") { return @("comunicacion", "communication", "pensamiento global", "marketing", "diseno", "diseño") }
        return @("lengua", "literatura", "language")
    }

    if ($SheetName -eq "ARTES") {
        if ($CurSecPr -eq "musica") { return @("musica", "music") }
        if ($CurSecPr -eq "teatro") { return @("teatro", "drama") }
        if ($CurSecPr -eq "audiovisual") { return @("audiovisual", "efectos especiales") }
        if ($CurSecPr -eq "danza") { return @("danza") }
        if ($CurSecPr -eq "arte") { return @("arte", "arts") }
        return @("performing", "arte", "arts", "danza", "musica")
    }

    if ($global:SheetAllowedSubjects.ContainsKey($SheetName)) {
        return $global:SheetAllowedSubjects[$SheetName]
    }

    return @()
}

function Find-PhidiasEvaluatedCount {
    <#
    .SYNOPSIS
        Realiza la búsqueda jerárquica de coincidencias entre la fila del reporte y Phidias.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [array]$PhidiasEntries,

        [Parameter(Mandatory = $true)]
        [string[]]$AllowedKeywords,

        [Parameter(Mandatory = $false)]
        [string]$DocenteClean,

        [Parameter(Mandatory = $false)]
        [string]$CursoClean,

        [Parameter(Mandatory = $true)]
        [hashtable]$CourseInfo
    )

    # Filtrar entradas de Phidias por las palabras clave de la materia / sección
    $candidates = $PhidiasEntries | Where-Object {
        $sub = $_.SubjectClean
        $ok = $false
        foreach ($kw in $AllowedKeywords) {
            if ($sub.Contains($kw)) { $ok = $true; break }
        }
        $ok
    }

    # Nivel 1: Match Docente exacto o contenido + Curso exacto o ShortCode
    if ($DocenteClean) {
        $m = $candidates | Where-Object {
            $isCourseMatch = ($_.CourseClean -eq $CursoClean) -or 
                             ($CourseInfo.ShortCode -and $_.CourseInfo.ShortCode -eq $CourseInfo.ShortCode)
            $isTeacherMatch = ($_.TeacherClean -eq $DocenteClean) -or 
                              ($_.TeacherClean.Contains($DocenteClean)) -or 
                              ($DocenteClean.Contains($_.TeacherClean))
            $isCourseMatch -and $isTeacherMatch
        }
        if ($m) {
            $best = $m | Sort-Object Evaluados -Descending | Select-Object -First 1
            return $best.Evaluados
        }
    }

    # Nivel 2: Match Docente (por palabras de apellido con longitud >= 4) + Curso
    if ($DocenteClean) {
        $docWords = $DocenteClean.Split(' ') | Where-Object { $_.Length -ge 4 }
        foreach ($entry in $candidates) {
            $isCourseMatch = ($entry.CourseClean -eq $CursoClean) -or 
                             ($CourseInfo.ShortCode -and $entry.CourseInfo.ShortCode -eq $CourseInfo.ShortCode)
            if ($isCourseMatch) {
                $mW = 0
                foreach ($w in $docWords) {
                    if ($entry.TeacherClean.Contains($w)) { $mW++ }
                }
                if ($mW -ge 1) {
                    return $entry.Evaluados
                }
            }
        }
    }

    # Nivel 3: Match Curso + Materia (cuando el docente cambió o la celda estaba vacía)
    $m = $candidates | Where-Object {
        ($_.CourseClean -eq $CursoClean) -or 
        ($CourseInfo.ShortCode -and $_.CourseInfo.ShortCode -eq $CourseInfo.ShortCode)
    }
    if ($m) {
        $best = $m | Sort-Object Evaluados -Descending | Select-Object -First 1
        return $best.Evaluados
    }

    # Nivel 4: Fallback especial para cursos consolidados por grado (ej. Artes Fifth)
    if ($DocenteClean) {
        $docWords = $DocenteClean.Split(' ') | Where-Object { $_.Length -ge 4 }
        foreach ($entry in $candidates) {
            if ($CourseInfo.Grade -and $entry.CourseInfo.Grade -eq $CourseInfo.Grade) {
                $mW = 0
                foreach ($w in $docWords) {
                    if ($entry.TeacherClean.Contains($w)) { $mW++ }
                }
                if ($mW -ge 1) {
                    return $entry.Evaluados
                }
            }
        }
    }

    return 0
}
