@echo off
rem Inventaire des macros VBA : choisit un dossier, lance l'analyse, ouvre le rapport.
setlocal
set "EXE=%~dp0inventaire-vba.exe"
if not exist "%EXE%" set "EXE=inventaire-vba.exe"
for /f "usebackq delims=" %%D in (`powershell -NoProfile -STA -Command "Add-Type -AssemblyName System.Windows.Forms; $d = New-Object System.Windows.Forms.FolderBrowserDialog; $d.Description = 'Choisissez le dossier ou le partage ' + [char]0xE0 + ' inventorier (lecture seule)'; if ($d.ShowDialog() -eq 'OK') { $d.SelectedPath }"`) do set "DOSSIER=%%D"
if not defined DOSSIER exit /b 1
set "RESULTATS=%USERPROFILE%\Documents\Inventaire VBA"
if not exist "%RESULTATS%" mkdir "%RESULTATS%"
cd /d "%RESULTATS%"
"%EXE%" "%DOSSIER%" --ouvrir
echo.
pause
