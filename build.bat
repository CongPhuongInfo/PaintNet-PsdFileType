@echo off
setlocal
REM ===========================================================================
REM  build.bat - builds Photoshop.dll (VB.NET port of PsdPlugin) with the
REM  .NET 9 SDK, and copies it into Paint.NET's FileTypes plugin folder.
REM ===========================================================================

REM Adjust this if Paint.NET is installed elsewhere.
set PAINTNET_DIR=C:\Program Files\paint.net
set FILETYPES_DIR=%PAINTNET_DIR%\FileTypes

where dotnet >nul 2>nul
if errorlevel 1 (
    echo ERROR: dotnet SDK not found on PATH. Install the .NET 9 SDK from
    echo        https://dotnet.microsoft.com/download/dotnet/9.0 and re-run.
    exit /b 1
)

echo Using:
dotnet --version

echo.
echo Building Photoshop.vbproj (net9.0-windows) ...
dotnet build "%~dp0Photoshop.vbproj" -c Release

if errorlevel 1 (
    echo.
    echo BUILD FAILED.
    exit /b 1
)

echo.
echo BUILD SUCCEEDED.

set OUT_DLL=%~dp0bin\Release\net9.0-windows\Photoshop.dll

if not exist "%OUT_DLL%" (
    echo.
    echo WARNING: Expected output DLL not found at:
    echo   %OUT_DLL%
    echo Check bin\Release\ for the actual output path and copy it manually.
    exit /b 1
)

if not exist "%FILETYPES_DIR%" (
    echo.
    echo WARNING: FileTypes folder not found at:
    echo   %FILETYPES_DIR%
    echo Paint.NET may not be installed there - edit PAINTNET_DIR near the top
    echo of this script, or copy "%OUT_DLL%" there manually.
    exit /b 1
)

echo.
echo Copying Photoshop.dll to "%FILETYPES_DIR%" ...
copy /Y "%OUT_DLL%" "%FILETYPES_DIR%\Photoshop.dll" >nul

if errorlevel 1 (
    echo.
    echo ERROR: Copy failed. If Paint.NET is currently running, close it and
    echo        re-run this script - the DLL may be locked while Paint.NET
    echo        has it loaded. You may also need to run this script as
    echo        Administrator to write into Program Files.
    exit /b 1
)

echo.
echo DONE. Photoshop.dll installed to "%FILETYPES_DIR%".
echo Restart Paint.NET to pick up the change.

endlocal
