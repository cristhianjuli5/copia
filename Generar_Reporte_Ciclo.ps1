<#
.SYNOPSIS
    Generador Automático de Reporte de Progreso por Periodo.
    Orquesta la indexación de Phidias, extracción de metas de la Plantilla y actualización del reporte.

.PARAMETER TargetCycle
    Número del ciclo a generar (por ejemplo: "24").
#>

[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string]$TargetCycle = "24"
)

$ErrorActionPreference = "Stop"

# Establecer directorio base de trabajo
$dir = $PSScriptRoot
if (-not $dir) { 
    $dir = Split-Path -Parent $MyInvocation.MyCommand.Definition
}
if (-not $dir) { 
    $dir = Get-Location 
}

# Rutas de archivos
$pathProgreso  = Join-Path $dir "PROGRESO_POR_PERIODO.xlsx"
$pathPhidias   = Join-Path $dir "Archivo_Phidias.xls.xlsx"
$pathPlantilla = Join-Path $dir "PLANTILLA CA PROGRESO POR PERIODO.xlsx"
$pathSalida    = Join-Path $dir "PROGRESO_POR_PERIODO_ACTUALIZADO.xlsx"
$modulesDir    = Join-Path $dir "modules"

Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host "   AUTOMATIZACION DE REPORTE DE PROGRESO POR PERIODO     " -ForegroundColor Cyan
Write-Host "   Ciclo a procesar: CICLO $TargetCycle                  " -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan

# Validar existencia de archivos fuente
$missingFiles = @()
foreach ($f in @($pathPhidias, $pathPlantilla, $pathProgreso)) {
    if (-not (Test-Path $f)) { $missingFiles += (Split-Path -Leaf $f) }
}
if ($missingFiles.Count -gt 0) {
    Write-Error "Faltan los siguientes archivos requeridos en el directorio:`n- $($missingFiles -join "`n- ")"
    exit 1
}

# Cargar módulos especializados
try {
    . (Join-Path $modulesDir "Common.ps1")
    . (Join-Path $modulesDir "PhidiasParser.ps1")
    . (Join-Path $modulesDir "PlantillaParser.ps1")
    . (Join-Path $modulesDir "ReportBuilder.ps1")
} catch {
    Write-Error "Error al cargar los modulos desde $modulesDir : $_"
    exit 1
}

# Inicializar motor de Excel COM
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.ScreenUpdating = $false

try {
    # ----------------------------------------------------
    # 1. PARSE ARCHIVO PHIDIAS
    # ----------------------------------------------------
    Write-Host "`n[1/3] Indexando Archivo_Phidias.xls.xlsx..." -ForegroundColor Yellow
    $phidiasEntries = Get-PhidiasEntries -ExcelApp $excel -PathPhidias $pathPhidias
    Write-Host "   -> $($phidiasEntries.Count) cursos indexados de Phidias." -ForegroundColor Green

    # ----------------------------------------------------
    # 2. PARSE PLANTILLA CA PROGRESO POR PERIODO
    # ----------------------------------------------------
    Write-Host "`n[2/3] Leyendo PLANTILLA CA PROGRESO POR PERIODO..." -ForegroundColor Yellow
    $plantillaData = Get-PlantillaData -ExcelApp $excel -PathPlantilla $pathPlantilla -TargetCycle $TargetCycle
    $plantillaExpected = $plantillaData.PlantillaExpected
    $llTable = $plantillaData.LLTable
    Write-Host "   -> $($plantillaExpected.Count) metas esperadas leidas de la plantilla." -ForegroundColor Green

    # ----------------------------------------------------
    # 3. ACTUALIZAR PROGRESO_POR_PERIODO
    # ----------------------------------------------------
    Write-Host "`n[3/3] Actualizando PROGRESO_POR_PERIODO con reglas inteligentes completas..." -ForegroundColor Yellow
    Update-ProgresoReport `
        -ExcelApp $excel `
        -PathProgreso $pathProgreso `
        -PathSalida $pathSalida `
        -TargetCycle $TargetCycle `
        -PhidiasEntries $phidiasEntries `
        -PlantillaExpected $plantillaExpected `
        -LLTable $llTable

    Write-Host "`n==========================================================" -ForegroundColor Green
    Write-Host " [EXITO] REPORTE ACTUALIZADO CON COINCIDENCIAS EXACTAS!   " -ForegroundColor Green
    Write-Host " Archivo generado: $pathSalida" -ForegroundColor Green
    Write-Host "==========================================================" -ForegroundColor Green

} catch {
    Write-Host "`n==========================================================" -ForegroundColor Red
    Write-Host " [ERROR] Ocurrio una excepcion durante el proceso:" -ForegroundColor Red
    Write-Host " $_" -ForegroundColor Red
    Write-Host "==========================================================" -ForegroundColor Red
    exit 1
} finally {
    if ($excel) {
        $excel.ScreenUpdating = $true
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}
