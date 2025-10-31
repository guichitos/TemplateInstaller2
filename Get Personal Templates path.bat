@echo off
setlocal EnableDelayedExpansion
set "KEY=HKCU\Software\Microsoft\Office\16.0\PowerPoint\Options"
set "VALUE=PersonalTemplates"
set "Data="

echo Revisando la clave del Registro: "%KEY%"

for /f "skip=2 tokens=1,2,*" %%A in ('reg query "%KEY%" /v "%VALUE%" 2^>nul') do (
    set "Data=%%C"
)

if defined Data (
    echo %VALUE% = !Data!
) else (
    echo No se pudo encontrar el valor "%VALUE%" en "%KEY%".
    exit /b 1
)

endlocal
