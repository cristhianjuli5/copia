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
    $grade = ""
    $letter = ""

    if     ($c -match "\bpioneros\b|\bpn\b")        { $grade = "pn" }
    elseif ($c -match "\bfirst garten\b|\bfg\b")    { $grade = "fg" }
    elseif ($c -match "\bgarten\b|\bgr\b|^g\s")     { $grade = "gr" }
    elseif ($c -match "\btransition\b|\btr\b|^t\s") { $grade = "tr" }
    elseif ($c -match "\bfirst\b|\b1\b")            { $grade = "1" }
    elseif ($c -match "\bsecond\b|\b2\b")           { $grade = "2" }
    elseif ($c -match "\bthird\b|\b3\b")            { $grade = "3" }
    elseif ($c -match "\bfourth\b|\b4\b")           { $grade = "4" }
    elseif ($c -match "\bfifth\b|\b5\b")            { $grade = "5" }
    elseif ($c -match "\bsixth\b|\b6\b")            { $grade = "6" }
    elseif ($c -match "\bseventh\b|\b7\b")          { $grade = "7" }
    elseif ($c -match "\beighth\b|\b8\b")           { $grade = "8" }
    elseif ($c -match "\bninth\b|\b9\b")            { $grade = "9" }
    elseif ($c -match "\btenth\b|\b10\b")           { $grade = "10" }
    elseif ($c -match "\beleventh\b|\b11\b")        { $grade = "11" }

    if ($c -match "\b([a-g])\b") {
        $letter = $matches[1].ToLower()
    }

    return @{
        Grade     = $grade
        Letter    = $letter
        ShortCode = "$grade$letter"
    }
}
