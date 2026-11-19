<#
.SYNOPSIS
    Módulo común con funciones de utilidad para normalización de texto y mapeo de cursos.
#>

function Format-Text {
    <#
    .SYNOPSIS
        Normaliza una cadena de texto eliminando acentos, caracteres especiales y espacios múltiples.
    #>
    param(
        [Parameter(Mandatory = $false)]
        [string]$str
    )

    if (-not $str) { return "" }
    
    # Separar caracteres de sus marcas diacríticas (FormD)
    $normalized = $str.Normalize([System.Text.NormalizationForm]::FormD)
    $sb = New-Object System.Text.StringBuilder
    foreach ($c in $normalized.ToCharArray()) {
        if ([System.Globalization.CharUnicodeInfo]::GetUnicodeCategory($c) -ne [System.Globalization.UnicodeCategory]::NonSpacingMark) {
            [void]$sb.Append($c)
        }
    }
    
    # Componer de nuevo en FormC y limpiar caracteres invisibles/especiales
    $s = $sb.ToString().Normalize([System.Text.NormalizationForm]::FormC)
    $s = $s.Replace([char]160, ' ').ToLower().Trim()
    $s = [System.Text.RegularExpressions.Regex]::Replace($s, "[^a-z0-9 ]", " ")
    $s = [System.Text.RegularExpressions.Regex]::Replace($s, "\s+", " ").Trim()
    return $s
}

function Get-ColLetter {
    <#
    .SYNOPSIS
        Convierte un índice de columna numérico (1-based) en su letra correspondiente de Excel (A, B, ..., Z, AA, ...).
    #>
    param(
        [Parameter(Mandatory = $true)]
        [int]$colNum
    )

    $div = $colNum
    $colLetter = ""
    while ($div -gt 0) {
        $mod = ($div - 1) % 26
        $colLetter = [char](65 + $mod) + $colLetter
        $div = [int](($div - $mod) / 26)
    }
    return $colLetter
}

function Get-GradeAndSection {
    <#
    .SYNOPSIS
        Analiza el nombre de un curso para extraer el grado normalizado y la sección (letra).
    #>
    param(
        [Parameter(Mandatory = $false)]
        [string]$course
    )

    $c = Format-Text $course

    # Asignación en expresión única — evita falsos positivos de PSScriptAnalyzer
    # sobre variables asignadas en ramas elseif separadas.
    $grade = $(
        if     ($c -match "\bpioneros\b|\bpn\b")        { "pn" }
        elseif ($c -match "\bfirst garten\b|\bfg\b")    { "fg" }
        elseif ($c -match "\bgarten\b|\bgr\b|^g\s")     { "gr" }
        elseif ($c -match "\btransition\b|\btr\b|^t\s") { "tr" }
        elseif ($c -match "\btenth\b|\b10[a-g]?\b")     { "10" }
        elseif ($c -match "\beleventh\b|\b11[a-g]?\b")  { "11" }
        elseif ($c -match "\bfirst\b|\b1[a-g]?\b")      { "1" }
        elseif ($c -match "\bsecond\b|\b2[a-g]?\b")     { "2" }
        elseif ($c -match "\bthird\b|\b3[a-g]?\b")      { "3" }
        elseif ($c -match "\bfourth\b|\b4[a-g]?\b")     { "4" }
        elseif ($c -match "\bfifth\b|\b5[a-g]?\b")      { "5" }
        elseif ($c -match "\bsixth\b|\b6[a-g]?\b")      { "6" }
        elseif ($c -match "\bseventh\b|\b7[a-g]?\b")    { "7" }
        elseif ($c -match "\beighth\b|\b8[a-g]?\b")     { "8" }
        elseif ($c -match "\bninth\b|\b9[a-g]?\b")      { "9" }
        else                                             { "" }
    )

    $letter = $(
        if ($c -match "\b(?:10|11|[1-9])([a-g])\b") { $matches[1].ToLower() }
        elseif ($c -match "\b([a-g])\b")            { $matches[1].ToLower() }
        else                                        { "" }
    )

    return @{
        Grade     = $grade
        Letter    = $letter
        ShortCode = "$grade$letter"
    }
}
