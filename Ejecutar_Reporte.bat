@echo off
chcp 65001 >nul
title Generador Automatico de Reporte de Progreso
cls
echo ==========================================================
echo    AUTOMATIZACION DE REPORTE DE PROGRESO POR PERIODO
echo ==========================================================
echo.
echo Este programa lee automaticamente:
echo   1. Archivo_Phidias.xls.xlsx
echo   2. PLANTILLA CA PROGRESO POR PERIODO.xlsx
echo   3. PROGRESO_POR_PERIODO.xlsx
echo.
echo Y genera el nuevo ciclo con formulas y colores en:
echo   - PROGRESO_POR_PERIODO_ACTUALIZADO.xlsx
echo ==========================================================
echo.
set /p CICLO="Ingresa el numero de Ciclo a generar (Ejemplo: 24) [Presiona ENTER para Ciclo 24]: "
if "%CICLO%"=="" set CICLO=24
echo.
echo [INFO] Procesando CICLO %CICLO%... Por favor espera unos segundos.
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Generar_Reporte_Ciclo.ps1" -TargetCycle "%CICLO%"
if %errorlevel% equ 0 (
    echo.
    echo ==========================================================
    echo  [LISTO] PROCESO COMPLETADO EXITOSAMENTE!
    echo  Abriendo el archivo de Excel resultante...
    echo ==========================================================
    start "" "%~dp0PROGRESO_POR_PERIODO_ACTUALIZADO.xlsx"
) else (
    echo.
    echo ==========================================================
    echo  [ALERTA] Ocurrio un inconveniente durante el proceso.
    echo ==========================================================
)
echo.
pause
