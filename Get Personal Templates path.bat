@echo off
setlocal EnableDelayedExpansion

set "VALUE=PersonalTemplates"
set "VERSION=16.0"

for %%A in (PowerPoint Word Excel) do (
    set "KEY=HKCU\Software\Microsoft\Office\%VERSION%\%%A\Options"
    set "Data="

    echo Revisando la clave del Registro: "!KEY!"

    for /f "skip=2 tokens=1,2,*" %%B in ('reg query "!KEY!" /v "%VALUE%" 2^>nul') do (
        set "Data=%%D"
    )

    if defined Data (
        echo %VALUE% = !Data!
    ) else (
        echo No se pudo encontrar el valor "%VALUE%" en "!KEY!".
    )

    echo.
)

endlocal
pause
