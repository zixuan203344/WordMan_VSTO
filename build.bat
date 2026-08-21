@echo off
setlocal

REM ============================================================
REM  WordMan VSTO - One-click build script (Debug + Release)
REM  Usage: double-click this file to build both configurations.
REM  Output: bin\Debug, bin\Release
REM  Note:  file must be saved as UTF-8. Requires VS installed.
REM         If VS is installed elsewhere, edit the MSBUILD path.
REM ============================================================

cd /d "%~dp0"

REM Clear oversized env var to avoid MSBuild node crash (MSB4166)
set "ACC_PRODUCT_CONFIG_V3="

REM --- Locate MSBuild.exe (edit this path if VS is installed elsewhere) ---
set "MSBUILD=D:\Program Files\Microsoft Visual Studio\18\Enterprise\MSBuild\Current\Bin\MSBuild.exe"

if not exist "%MSBUILD%" (
    echo [ERROR] MSBuild.exe not found at:
    echo         %MSBUILD%
    echo         Please edit the "MSBUILD" path in this script.
    pause
    exit /b 1
)

echo ============================================================
echo  WordMan build started: Debug + Release
echo ============================================================

set "FAILED=0"

for %%C in (Debug Release) do (
    echo.
    echo ---------- Building %%C ----------
    "%MSBUILD%" "%~dp0WordMan.csproj" -p:Configuration=%%C -t:Build -v:minimal -nologo
    if errorlevel 1 (
        echo [FAILED] %%C build finished with errors.
        set "FAILED=1"
    ) else (
        echo [OK] %%C build succeeded.
    )
)

echo.
if "%FAILED%"=="1" (
    echo ============================================================
    echo  BUILD RESULT: FAILED - see errors above
    echo ============================================================
) else (
    echo ============================================================
    echo  BUILD RESULT: ALL SUCCEEDED
    echo  Output: bin\Debug , bin\Release
    echo ============================================================
)

pause
endlocal
