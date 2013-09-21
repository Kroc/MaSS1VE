@ECHO OFF
CLS

REM Compile the app
ECHO * Compiling MaSS1VE.exe
"%PROGRAMFILES%\Microsoft Visual Studio\VB98\vb6.exe" /make "..\MaSS1VE.vbp" /outdir "..\RELEASE"

REM Get the version number from the MaSS1VE executable
FOR /f "delims=" %%R IN ('CScript //NOLOGO GetFileVersion.vbs ..\RELEASE\MaSS1VE.exe') DO (SET "EXEVER=%%R")
ECHO - Version number is: %EXEVER%

REM Convert this to a shorter version number in the VB6 style. VB6 does not use the Build value
FOR /f "tokens=1,2,3,4 delims=." %%A IN ("%EXEVER%") DO (SET "Major=%%A" & SET "Minor=%%B" & SET "Build=%%C" & SET "Revision=%%D")
SET "VB6VER=%MAJOR%.%MINOR%.%REVISION%"

REM Now build the installer
ECHO.
ECHO * Building Installer...
ECHO.
"%PROGRAMFILES%\NSIS\makensis.exe" /DPRODUCT_VERSION_WIN="%EXEVER%" /DPRODUCT_VERSION_VB6="%VB6VER%" /V2 "Installer Source.nsi"

ECHO.
PAUSE