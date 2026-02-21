@echo off

title Windows 21H2to7 Language Pack Installer

echo.
echo Language Pack for Windows 21H2to7
echo ==========================================================================
echo.
echo This script will run Scripts\installer.ps1 with admin privileges
echo Make sure that you've downloaded this installer from https://github.com/ImSwordQueen/Win10to7LanguagePacks
echo and not any other site that isn't that link above.
echo.
echo.
echo If you are sure that you've downloaded this from the correct source and you are sure
echo the script is safe.
echo.

pause

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0\scripts\installer.ps1" %*

::THIS ONLY LOADS THE POWERSHELL SCRIPT TO INSTALL THE LANGUAGE PACK.
::THIS IS NOT THE INSTALLER ITSELF. CHECK "Scripts\installer.ps1" FOR THE INSTALLER CODE.
