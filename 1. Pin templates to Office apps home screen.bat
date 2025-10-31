rem _Main_installer.bat
@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul

rem ======================================================
rem === UNIVERSAL OFFICE TEMPLATES INSTALLER (MAIN) ======
rem ------------------------------------------------------
rem Entry point for the modular installer system.
rem Coordinates environment checks, closes Office apps,
rem installs base templates for Word, PowerPoint, and Excel,
rem copies user custom templates, and logs all operations.
rem ======================================================

rem === DESIGN / DEBUG MODE CONTROL ======================
rem If IsDesignModeEnabled=true  → shows console messages and generates log.
rem If IsDesignModeEnabled=false → runs silently (no output, no log file).
set "IsDesignModeEnabled=true"

rem === Base paths and library references ================
set "BaseDirectoryPath=%~dp0"
set "LogsDirectoryPath=%BaseDirectoryPath%logs"
set "LogFilePath=%LogsDirectoryPath%\install_log_all.txt"

echo Executing. Please wait...
rem === Initialize log only if design mode is enabled ====
if /I "%IsDesignModeEnabled%"=="true" (
    if not exist "%LogsDirectoryPath%" mkdir "%LogsDirectoryPath%"
    echo. > "%LogFilePath%"
    echo [%DATE% %TIME%] --- START TEMPLATES INSTALLATION --- >> "%LogFilePath%"
)

rem === Header message =============
if /I "%IsDesignModeEnabled%"=="true" (
    title TEMPLATE INSTALLER - DEBUG MODE
    echo [DEBUG] Design mode is enabled.
    echo [INFO] Script is running from: %BaseDirectoryPath%
)

rem === Environment verification and Office shutdown =====
if /I "%IsDesignModeEnabled%"=="true" (
    echo.
    echo [INFO] Verifying environment and closing Office applications...
    call :CheckEnvironment "%LogFilePath%"
    call :CloseOfficeApps "%LogFilePath%"
    echo [OK] Environment verification and Office app closure completed.
    echo [OK] Environment verification and Office app closure completed. >> "%LogFilePath%"
) else (
    call :CheckEnvironment "" >nul 2>&1
    call :CloseOfficeApps "" >nul 2>&1
)



rem === Install base templates for Word, PowerPoint, Excel ===
if /I "%IsDesignModeEnabled%"=="true" (
    echo.
    echo [INFO] Starting base template installation phase...
    rem --- Word templates (Normal.dotx / Normal.dotm) ---
    call :InstallApp "WORD" "Normal.dotx" "%APPDATA%\Microsoft\Templates" "Normal.dotx" "%LogFilePath%" "%BaseDirectoryPath%" "%IsDesignModeEnabled%"
    call :InstallApp "WORD" "Normal.dotm" "%APPDATA%\Microsoft\Templates" "Normal.dotm" "%LogFilePath%" "%BaseDirectoryPath%" "%IsDesignModeEnabled%"
    rem --- PowerPoint templates (Blank.potx / Blank.potm) ---
    call :InstallApp "POWERPOINT" "Blank.potx" "%APPDATA%\Microsoft\Templates" "Blank.potx" "%LogFilePath%" "%BaseDirectoryPath%" "%IsDesignModeEnabled%"
    call :InstallApp "POWERPOINT" "Blank.potm" "%APPDATA%\Microsoft\Templates" "Blank.potm" "%LogFilePath%" "%BaseDirectoryPath%" "%IsDesignModeEnabled%"
    rem --- Excel templates (Book / Sheet in xltx & xltm) ---
    call :InstallApp "EXCEL" "Book.xltx" "%APPDATA%\Microsoft\Excel\XLSTART" "Book.xltx" "%LogFilePath%" "%BaseDirectoryPath%" "%IsDesignModeEnabled%"
    call :InstallApp "EXCEL" "Book.xltm" "%APPDATA%\Microsoft\Excel\XLSTART" "Book.xltm" "%LogFilePath%" "%BaseDirectoryPath%" "%IsDesignModeEnabled%"
    call :InstallApp "EXCEL" "Sheet.xltx" "%APPDATA%\Microsoft\Excel\XLSTART" "Sheet.xltx" "%LogFilePath%" "%BaseDirectoryPath%" "%IsDesignModeEnabled%"
    call :InstallApp "EXCEL" "Sheet.xltm" "%APPDATA%\Microsoft\Excel\XLSTART" "Sheet.xltm" "%LogFilePath%" "%BaseDirectoryPath%" "%IsDesignModeEnabled%"
) else (
    rem --- Word templates (Normal.dotx / Normal.dotm) ---
    call :InstallApp "WORD" "Normal.dotx" "%APPDATA%\Microsoft\Templates" "Normal.dotx" "" "%BaseDirectoryPath%" "%IsDesignModeEnabled%" >nul 2>&1
    call :InstallApp "WORD" "Normal.dotm" "%APPDATA%\Microsoft\Templates" "Normal.dotm" "" "%BaseDirectoryPath%" "%IsDesignModeEnabled%" >nul 2>&1
    rem --- PowerPoint templates (Blank.potx / Blank.potm) ---
    call :InstallApp "POWERPOINT" "Blank.potx" "%APPDATA%\Microsoft\Templates" "Blank.potx" "" "%BaseDirectoryPath%" "%IsDesignModeEnabled%" >nul 2>&1
    call :InstallApp "POWERPOINT" "Blank.potm" "%APPDATA%\Microsoft\Templates" "Blank.potm" "" "%BaseDirectoryPath%" "%IsDesignModeEnabled%" >nul 2>&1
    rem --- Excel templates (Book / Sheet in xltx & xltm) ---
    call :InstallApp "EXCEL" "Book.xltx" "%APPDATA%\Microsoft\Excel\XLSTART" "Book.xltx" "" "%BaseDirectoryPath%" "%IsDesignModeEnabled%" >nul 2>&1
    call :InstallApp "EXCEL" "Book.xltm" "%APPDATA%\Microsoft\Excel\XLSTART" "Book.xltm" "" "%BaseDirectoryPath%" "%IsDesignModeEnabled%" >nul 2>&1
    call :InstallApp "EXCEL" "Sheet.xltx" "%APPDATA%\Microsoft\Excel\XLSTART" "Sheet.xltx" "" "%BaseDirectoryPath%" "%IsDesignModeEnabled%" >nul 2>&1
    call :InstallApp "EXCEL" "Sheet.xltm" "%APPDATA%\Microsoft\Excel\XLSTART" "Sheet.xltm" "" "%BaseDirectoryPath%" "%IsDesignModeEnabled%" >nul 2>&1
)


rem === Detect Office personal template directories ==========
if /I "%IsDesignModeEnabled%"=="true" (
    echo.
    echo [INFO] Detecting Office personal template folders...
)

if /I "%IsDesignModeEnabled%"=="true" (
    call :DetectOfficePaths "%LogFilePath%" "%IsDesignModeEnabled%"
) else (
    call :DetectOfficePaths "" "%IsDesignModeEnabled%" >nul 2>&1
)

rem === Copy custom templates and update registry MRUs ===
if /I "%IsDesignModeEnabled%"=="true" (
    echo.
    echo [INFO] Starting custom template copy phase...
)

if /I "%IsDesignModeEnabled%"=="true" (
    echo [DEBUG] Executing internal copy routine for custom templates... >> "%LogFilePath%"
    call :CopyAll "%LogFilePath%" "%BaseDirectoryPath%" "%IsDesignModeEnabled%"
) else (
    call :CopyAll "%LogFilePath%" "%BaseDirectoryPath%" "%IsDesignModeEnabled%" >nul 2>&1
)

rem === Finalization and optional pause ==================
if /I "%IsDesignModeEnabled%"=="true" (
    echo [%DATE% %TIME%] --- UNIVERSAL INSTALLATION COMPLETED --- >> "%LogFilePath%"
    echo.
    echo [FINAL] Universal Office Template installation completed successfully.
    echo Log file saved at: "%LogFilePath%"
    echo ----------------------------------------------------
    pause
)
Echo Successfully executed.
pause
goto :EndOfScript

:GenerateFutureTCode
rem Args: OUTPUT_VAR LOG_FILE DESIGN_MODE
set "OutputVarName=%~1"
set "LogFilePath=%~2"
set "DesignMode=%~3"

if not defined DesignMode set "DesignMode=true"

setlocal EnableExtensions EnableDelayedExpansion
set "OutputVarName=%OutputVarName%"
set "LogFilePath=%LogFilePath%"
set "DesignMode=%DesignMode%"

rem 1. Obtener fecha y hora actual en UTC
for /f "tokens=2 delims==." %%a in ('wmic os get localdatetime /value') do set "dt=%%a"
set "YYYY=!dt:~0,4!"
set "MM=!dt:~4,2!"
set "DD=!dt:~6,2!"
set "hh=!dt:~8,2!"
set "nn=!dt:~10,2!"
set "ss=!dt:~12,2!"

rem 2. Calcular la fecha dentro de 10 años usando PowerShell (solo para precisión)
for /f %%F in ('powershell -NoLogo -Command "(Get-Date).AddYears(10).ToFileTimeUtc().ToString(''X16'')"') do set "HEX=%%F"

rem 3. Formar el código tipo [Txxxxxxxxxxxxxxxx]
set "TCODE=[T!HEX!]"

if /I "!DesignMode!"=="true" (
    echo ==============================================
    echo Fecha actual: !YYYY!-!MM!-!DD! !hh!:!nn!:!ss!
    echo Fecha futura (+10 años): (calculada por PowerShell)
    echo Código T: !TCODE!
    echo ==============================================
    if defined LogFilePath (
        echo [%DATE% %TIME%] Generated future template code: !TCODE! >> "!LogFilePath!"
    )
)

endlocal & (
    if defined OutputVarName set "%OutputVarName%=%TCODE%"
)
exit /b

:InstallApp
rem Args: APP SRC_NAME DST_DIR DST_NAME LOG_FILE BASE_DIR DESIGN_MODE
setlocal
set "AppName=%~1"
set "SourceFileName=%~2"
set "DestinationDirectory=%~3"
set "DestinationFileName=%~4"
set "LogFilePath=%~5"
set "SourceDirectory=%~6"
set "DesignMode=%~7"

set "SourceFilePath=%SourceDirectory%%SourceFileName%"
set "DestinationFilePath=%DestinationDirectory%\%DestinationFileName%"
set "BackupFilePath=%DestinationDirectory%\%~n4_backup%~x4"

if not exist "%SourceFilePath%" (
    if /I "%DesignMode%"=="true" echo [ERROR] Source file not found: "%SourceFilePath%"
    if /I "%DesignMode%"=="true" if defined LogFilePath echo [ERROR] Source file not found "%SourceFilePath%". >> "%LogFilePath%"
    endlocal
    exit /b
)

if not exist "%DestinationDirectory%" mkdir "%DestinationDirectory%" 2>nul

if exist "%DestinationFilePath%" (
    copy /Y "%DestinationFilePath%" "%BackupFilePath%" >nul 2>&1
    if /I "%DesignMode%"=="true" (
        echo [BACKUP] Created for %AppName% template at "%BackupFilePath%"
        if defined LogFilePath echo [BACKUP] Created for %AppName% template at "%BackupFilePath%" >> "%LogFilePath%"
    )
)

copy /Y "%SourceFilePath%" "%DestinationFilePath%" >nul 2>&1
if exist "%DestinationFilePath%" (
    if /I "%DesignMode%"=="true" (
        echo [OK] Installed %AppName% template at "%DestinationFilePath%"
        if defined LogFilePath echo [OK] Installed %AppName% template at "%DestinationFilePath%" >> "%LogFilePath%"
    )
) else (
    if /I "%DesignMode%"=="true" (
        echo [ERROR] Copy failed for "%SourceFilePath%"
        if defined LogFilePath echo [ERROR] Copy failed for "%SourceFilePath%" >> "%LogFilePath%"
    )
)

endlocal
exit /b

:CheckEnvironment
rem Args: LOG_FILE
set "LOG_FILE=%~1"

if defined LOG_FILE (
    echo [%DATE% %TIME%] Checking environment... >> "%LOG_FILE%"
)

openfiles >nul 2>&1
if %errorlevel% NEQ 0 (
    if defined LOG_FILE (
        echo [%DATE% %TIME%] Elevation required. Attempting to relaunch as admin... >> "%LOG_FILE%"
    )
    powershell -NoProfile -ExecutionPolicy Bypass -Command ^
        "Start-Process 'cmd.exe' -ArgumentList '/c','\"%~f0\"' -Verb RunAs"
    exit /b
)

if defined LOG_FILE (
    echo [%DATE% %TIME%] Environment check passed (already running as admin). >> "%LOG_FILE%"
)

exit /b

:DetectOfficePaths
rem Args: LOG_FILE, DESIGN_MODE
setlocal enabledelayedexpansion
set "LOG_FILE=%~1"
set "DESIGN_MODE=%~2"

set "WORD_PATH="
set "PPT_PATH="
set "EXCEL_PATH="
set "WORD_PATH_FALLBACK=0"
set "PPT_PATH_FALLBACK=0"
set "EXCEL_PATH_FALLBACK=0"
set "DEFAULT_CUSTOM_DIR=%USERPROFILE%\Documents\Custom Office Templates"
set "OFFICE_VERSIONS=16.0 15.0 14.0 12.0"

for %%V in (!OFFICE_VERSIONS!) do (
    if not defined WORD_PATH (
        for /f "tokens=1,2,*" %%A in (
          'reg query "HKCU\Software\Microsoft\Office\%%V\Word\Options" /v "PersonalTemplates" 2^>nul ^| find /I "PersonalTemplates"'
        ) do set "WORD_PATH=%%C"
    )
    if not defined PPT_PATH (
        for /f "tokens=1,2,*" %%A in (
          'reg query "HKCU\Software\Microsoft\Office\%%V\PowerPoint\Options" /v "PersonalTemplates" 2^>nul ^| find /I "PersonalTemplates"'
        ) do set "PPT_PATH=%%C"
    )
    if not defined EXCEL_PATH (
        for /f "tokens=1,2,*" %%A in (
          'reg query "HKCU\Software\Microsoft\Office\%%V\Excel\Options" /v "PersonalTemplates" 2^>nul ^| find /I "PersonalTemplates"'
        ) do set "EXCEL_PATH=%%C"
    )
)

for %%V in (!OFFICE_VERSIONS!) do (
    if not defined WORD_PATH (
        for /f "tokens=1,2,*" %%A in (
          'reg query "HKCU\Software\Microsoft\Office\%%V\Common\General" /v "UserTemplates" 2^>nul ^| find /I "UserTemplates"'
        ) do set "WORD_PATH=%%C"
    )
    if not defined PPT_PATH (
        for /f "tokens=1,2,*" %%A in (
          'reg query "HKCU\Software\Microsoft\Office\%%V\Common\General" /v "UserTemplates" 2^>nul ^| find /I "UserTemplates"'
        ) do set "PPT_PATH=%%C"
    )
    if not defined EXCEL_PATH (
        for /f "tokens=1,2,*" %%A in (
          'reg query "HKCU\Software\Microsoft\Office\%%V\Common\General" /v "UserTemplates" 2^>nul ^| find /I "UserTemplates"'
        ) do set "EXCEL_PATH=%%C"
    )
)

if not defined WORD_PATH if exist "!DEFAULT_CUSTOM_DIR!" (
    set "WORD_PATH=!DEFAULT_CUSTOM_DIR!"
    set "WORD_PATH_FALLBACK=1"
)
if not defined PPT_PATH if exist "!DEFAULT_CUSTOM_DIR!" (
    set "PPT_PATH=!DEFAULT_CUSTOM_DIR!"
    set "PPT_PATH_FALLBACK=1"
)
if not defined EXCEL_PATH if exist "!DEFAULT_CUSTOM_DIR!" (
    set "EXCEL_PATH=!DEFAULT_CUSTOM_DIR!"
    set "EXCEL_PATH_FALLBACK=1"
)

call :CleanPath WORD_PATH
call :CleanPath PPT_PATH
call :CleanPath EXCEL_PATH

if /I "!DESIGN_MODE!"=="true" (
    if defined WORD_PATH (
        if "!WORD_PATH_FALLBACK!"=="1" (
            echo [INFO] Word templates folder defaulted to: !WORD_PATH!
        ) else (
            echo [DEBUG] Word templates folder detected: !WORD_PATH!
        )
    ) else (
        echo [WARNING] No Word templates folder detected from registry.
    )
    if defined PPT_PATH (
        if "!PPT_PATH_FALLBACK!"=="1" (
            echo [INFO] PowerPoint templates folder defaulted to: !PPT_PATH!
        ) else (
            echo [DEBUG] PowerPoint templates folder detected: !PPT_PATH!
        )
    ) else (
        echo [WARNING] No PowerPoint templates folder detected from registry.
    )
    if defined EXCEL_PATH (
        if "!EXCEL_PATH_FALLBACK!"=="1" (
            echo [INFO] Excel templates folder defaulted to: !EXCEL_PATH!
        ) else (
            echo [DEBUG] Excel templates folder detected: !EXCEL_PATH!
        )
    ) else (
        echo [WARNING] No Excel templates folder detected from registry.
    )
)

if defined LOG_FILE (
    if defined WORD_PATH  echo [%DATE% %TIME%] Word templates folder detected or defaulted: !WORD_PATH! >> "!LOG_FILE!"
    if not defined WORD_PATH echo [%DATE% %TIME%] Word templates folder not found in registry. >> "!LOG_FILE!"
    if defined WORD_PATH if "!WORD_PATH_FALLBACK!"=="1" echo [%DATE% %TIME%] Word templates folder defaulted to user documents. >> "!LOG_FILE!"
    if defined PPT_PATH   echo [%DATE% %TIME%] PowerPoint templates folder detected or defaulted: !PPT_PATH! >> "!LOG_FILE!"
    if not defined PPT_PATH echo [%DATE% %TIME%] PowerPoint templates folder not found in registry. >> "!LOG_FILE!"
    if defined PPT_PATH if "!PPT_PATH_FALLBACK!"=="1" echo [%DATE% %TIME%] PowerPoint templates folder defaulted to user documents. >> "!LOG_FILE!"
    if defined EXCEL_PATH echo [%DATE% %TIME%] Excel templates folder detected or defaulted: !EXCEL_PATH! >> "!LOG_FILE!"
    if not defined EXCEL_PATH echo [%DATE% %TIME%] Excel templates folder not found in registry. >> "!LOG_FILE!"
    if defined EXCEL_PATH if "!EXCEL_PATH_FALLBACK!"=="1" echo [%DATE% %TIME%] Excel templates folder defaulted to user documents. >> "!LOG_FILE!"
)

endlocal & (
    set "WORD_PATH=%WORD_PATH%"
    set "PPT_PATH=%PPT_PATH%"
    set "EXCEL_PATH=%EXCEL_PATH%"
)
exit /b

:CleanPath
setlocal enabledelayedexpansion
set "VAR_NAME=%~1"
for /f "tokens=2 delims==" %%A in ('set !VAR_NAME! 2^>nul') do set "VALUE=%%A"

if defined VALUE (
    set "VALUE=!VALUE:"=!"
    for /f "tokens=* delims= " %%Z in ('echo(!VALUE!') do set "VALUE=%%Z"
    set "VALUE=!VALUE:        =!"
    if "!VALUE:~0,1!"=="\" set "VALUE=C:!VALUE!"
    if "!VALUE:~-1!"=="\" set "VALUE=!VALUE:~0,-1!"
)

endlocal & if defined VALUE set "%~1=%VALUE%"
exit /b

:Log
rem Args: LOG_FILE, MESSAGE
set "LOG_FILE=%~1"
shift
echo [%DATE% %TIME%] %* >> "%LOG_FILE%"
exit /b

:CloseOfficeApps
rem Args: LOG_FILE
set "LOG_FILE=%~1"
call :Log "%LOG_FILE%" Closing Office apps...
taskkill /IM WINWORD.EXE /F >nul 2>&1
taskkill /IM POWERPNT.EXE /F >nul 2>&1
taskkill /IM EXCEL.EXE /F >nul 2>&1
exit /b

:CopyAll
rem Args: LOG_FILE BASE_DIR IsDesignModeEnabled
setlocal enabledelayedexpansion
set "LOG_FILE=%~1"
set "BASE_DIR=%~2"
set "IsDesignModeEnabled=%~3"

set /a TOTAL_FILES=0
set /a TOTAL_ERRORS=0

if defined WORD_PATH if "!WORD_PATH:~-1!"=="\" set "WORD_PATH=!WORD_PATH:~0,-1!"
if defined PPT_PATH  if "!PPT_PATH:~-1!"=="\"  set "PPT_PATH=!PPT_PATH:~0,-1!"
if defined EXCEL_PATH if "!EXCEL_PATH:~-1!"=="\" set "EXCEL_PATH=!EXCEL_PATH:~0,-1!"

if /I "%IsDesignModeEnabled%"=="true" (
    title COPY_TEMPLATES DEBUG MODE
    echo [DEBUG] copy_templates.bat started
    echo [INFO] Script running from: %~dp0
    echo [INFO] Arguments: %*
    echo.
)

if /I "%IsDesignModeEnabled%"=="true" (
    echo.
    echo [DEBUG] === Template destinations received ===
    echo   WORD_PATH = !WORD_PATH!
    echo   PPT_PATH  = !PPT_PATH!
    echo   EXCEL_PATH= !EXCEL_PATH!
    echo [DEBUG] ==================================
    echo.
)

rem ==========================================================
rem === STAGE 1: DETECT MRU PATHS ============================
rem ==========================================================
call :DetectMRUPath POWERPOINT
if /I "%IsDesignModeEnabled%"=="true" (
    setlocal enabledelayedexpansion
    echo [DEBUG] PowerPoint MRU detected: !PPT_MRU_PATH!
    endlocal
)
if defined LOG_FILE echo [DEBUG] PowerPoint MRU detected: %PPT_MRU_PATH% >> "%LOG_FILE%"

call :DetectMRUPath WORD
if /I "%IsDesignModeEnabled%"=="true" echo [DEBUG] Word MRU detected: %WORD_MRU_PATH%
if defined LOG_FILE echo [DEBUG] Word MRU detected: %WORD_MRU_PATH% >> "%LOG_FILE%"

call :DetectMRUPath EXCEL
if /I "%IsDesignModeEnabled%"=="true" echo [DEBUG] Excel MRU detected: %EXCEL_MRU_PATH%
if defined LOG_FILE echo [DEBUG] Excel MRU detected: %EXCEL_MRU_PATH% >> "%LOG_FILE%"

if /I "%IsDesignModeEnabled%"=="true" (
    echo.
)

rem ==========================================================
rem === STAGE 2: FILE LISTING AND VALIDATION ================
rem ==========================================================
if /I "%IsDesignModeEnabled%"=="true" (
    echo [DEBUG] --- Scanning BASE_DIR for templates --- >> "%LOG_FILE%"
    echo [INFO] Searching templates in "%BASE_DIR%"...
    echo -----------------------------------------------
    dir /b "%BASE_DIR%\*.dot*" "%BASE_DIR%\*.pot*" "%BASE_DIR%\*.xlt*" 2>nul
    echo -----------------------------------------------
    echo.
)

if errorlevel 1 (
    if /I "%IsDesignModeEnabled%"=="true" (
        echo [WARNING] No template files found in "%BASE_DIR%".
        echo [WARNING] No .dotx / .potx / .xltx files detected. >> "%LOG_FILE%"
    )
)

rem ==========================================================
rem === STAGE 3: DESTINATION PATH VALIDATION ================
rem ==========================================================
if /I "%IsDesignModeEnabled%"=="true" (
    echo [DEBUG] Verifying destination paths...
)
for %%P in ("!WORD_PATH!" "!PPT_PATH!" "!EXCEL_PATH!") do (
    if exist "%%~P" (
        if /I "%IsDesignModeEnabled%"=="true" echo [OK] Valid folder: %%~P
    ) else (
        if /I "%IsDesignModeEnabled%"=="true" (
            echo [ERROR] Missing folder: %%~P
            echo [ERROR] Missing folder: %%~P >> "%LOG_FILE%"
        )
    )
)
if /I "%IsDesignModeEnabled%"=="true" (
    echo.
)

rem ==========================================================
rem === STAGE 4: FILE COPY AND REGISTRATION =================
rem ==========================================================
if /I "%IsDesignModeEnabled%"=="true" (
    echo [DEBUG] Starting file copy stage...
    echo [DEBUG] BASE_DIR = "%BASE_DIR%"
    echo -----------------------------------------------
)

for %%F in ("%BASE_DIR%\*.dotx" "%BASE_DIR%\*.dotm" "%BASE_DIR%\*.potx" "%BASE_DIR%\*.potm" "%BASE_DIR%\*.xltx" "%BASE_DIR%\*.xltm") do (
    if exist "%%~fF" (
        set "FN=%%~nxF"
        set "EXT=%%~xF"

        rem === Skip generic templates ===
        set "SKIP=0"
        if /I "!FN!"=="Normal.dotx" set "SKIP=1"
        if /I "!FN!"=="Blank.potx" set "SKIP=1"
        if /I "!FN!"=="Book.xltx" set "SKIP=1"
        if /I "!FN!"=="Normal.dotm" set "SKIP=1"
        if /I "!FN!"=="Blank.potm" set "SKIP=1"
        if /I "!FN!"=="Book.xltm" set "SKIP=1"
        if /I "!FN!"=="Sheet.xltx" set "SKIP=1"
        if /I "!FN!"=="Sheet.xltm" set "SKIP=1"

        rem === Determine destination ===
        set "DEST="
        if /I "!EXT!"==".dotx" set "DEST=!WORD_PATH!"
        if /I "!EXT!"==".dotm" set "DEST=!WORD_PATH!"
        if /I "!EXT!"==".potx" set "DEST=!PPT_PATH!"
        if /I "!EXT!"==".potm" set "DEST=!PPT_PATH!"
        if /I "!EXT!"==".xltx" set "DEST=!EXCEL_PATH!"
        if /I "!EXT!"==".xltm" set "DEST=!EXCEL_PATH!"

        if /I "%IsDesignModeEnabled%"=="true" (
            echo.
            echo [DEBUG] Processing file: !FN!
            echo [DEBUG] Extension detected: !EXT!
        )

        if "!SKIP!"=="1" (
            if /I "%IsDesignModeEnabled%"=="true" (
                echo [INFO] Skipped generic file: !FN!
                echo [INFO] Skipped generic: !FN! >> "%LOG_FILE%"
            )
        ) else if defined DEST (
            if /I "%IsDesignModeEnabled%"=="true" (
                echo [DEBUG] Destination assigned: !DEST!
                echo [ACTION] Copying !FN! → !DEST!
                echo [DEBUG] Copying !FN! to !DEST! >> "%LOG_FILE%"
            )
            mkdir "!DEST!" 2>nul
            if /I "%IsDesignModeEnabled%"=="true" (
                copy /Y "%%~fF" "!DEST!\" >> "%LOG_FILE%" 2>&1
            ) else (
                copy /Y "%%~fF" "!DEST!\" >nul 2>&1
            )

            if exist "!DEST!\!FN!" (
                if /I "%IsDesignModeEnabled%"=="true" (
                    echo [OK] Successfully copied: !FN!
                    echo [RESULT] Success → !FN! >> "%LOG_FILE%"
                )
                set /a TOTAL_FILES+=1
                echo /** template copied
                rem === ADDED: MRU registration ===
                if /I "!EXT!"==".potx" call :SimulateRegEntry POWERPOINT "!FN!" "!DEST!\!FN!" "%LOG_FILE%"
                if /I "!EXT!"==".potm" call :SimulateRegEntry POWERPOINT "!FN!" "!DEST!\!FN!" "%LOG_FILE%"
                if /I "!EXT!"==".dotx" call :SimulateRegEntry WORD "!FN!" "!DEST!\!FN!" "%LOG_FILE%"
                if /I "!EXT!"==".dotm" call :SimulateRegEntry WORD "!FN!" "!DEST!\!FN!" "%LOG_FILE%"
                if /I "!EXT!"==".xltx" call :SimulateRegEntry EXCEL "!FN!" "!DEST!\!FN!" "%LOG_FILE%"
                if /I "!EXT!"==".xltm" call :SimulateRegEntry EXCEL "!FN!" "!DEST!\!FN!" "%LOG_FILE%"
                rem === END ADDED ===

            ) else (
                if /I "%IsDesignModeEnabled%"=="true" (
                    echo [ERROR] Failed to copy: !FN!
                    echo [RESULT] Error → !FN! >> "%LOG_FILE%"
                )
                set /a TOTAL_ERRORS+=1
            )
        ) else (
            if /I "%IsDesignModeEnabled%"=="true" (
                echo [WARNING] No destination assigned for !FN!
                echo [WARNING] No destination → !FN! >> "%LOG_FILE%"
            )
        )
        if /I "%IsDesignModeEnabled%"=="true" (
            echo -----------------------------------------------
        )
    )
)

if /I "%IsDesignModeEnabled%"=="true" (
    echo [DEBUG] Copy loop finished
    echo [DEBUG] TOTAL_FILES=!TOTAL_FILES! TOTAL_ERRORS=!TOTAL_ERRORS!
)

rem ==========================================================
rem === STAGE 5: FINAL SUMMARY ==============================
rem ==========================================================
if /I "%IsDesignModeEnabled%"=="true" (
    echo.
    echo [FINAL] Copy phase completed.
    echo   Files copied: !TOTAL_FILES!
    echo   Files with errors: !TOTAL_ERRORS!
    echo ----------------------------------------------------------
    echo [DEBUG] Total copied: !TOTAL_FILES!, errors: !TOTAL_ERRORS! >> "%LOG_FILE%"
    echo.
)

endlocal
exit /b

rem ==========================================================
rem === Registry helper routines (formerly registry_tools) ===
rem ==========================================================

:DetectAdalContainer
rem Args: OUT_ID_VAR OUT_PATH_VAR [APP_REG_NAME]
set "TARGET_ID=%~1"
set "TARGET_PATH=%~2"
set "TARGET_APP=%~3"
setlocal enabledelayedexpansion
set "FOUND_ID="
set "FOUND_PATH="
set "APP_FILTER=%TARGET_APP%"
set "APP_LIST=PowerPoint Word Excel"
if defined APP_FILTER set "APP_LIST=!APP_FILTER!"

rem --- Buscar primero en las rutas comunes de Recent Templates ---
for %%V in (16.0 15.0 14.0 12.0) do (
    for %%A in (!APP_LIST!) do (
        set "BASE=HKCU\Software\Microsoft\Office\%%V\%%A\Recent Templates"
        for /f "tokens=* delims=" %%K in ('reg query "!BASE!" 2^>nul ^| findstr /R /C:"ADAL_" /C:"Livelid_"') do (
            set "FOUND_PATH=%%K"
            set "FOUND_ID=%%~nK"
            goto :dac_found
        )
    )
)

rem --- Búsqueda amplia dentro de toda la rama Office ---
for %%S in ("ADAL_" "Livelid_") do (
    for /f "tokens=* delims=" %%L in ('reg query "HKCU\Software\Microsoft\Office" /f %%~S /s 2^>nul ^| findstr /I "HKEY_CURRENT_USER"') do (
        set "LINE=%%L"
        if not defined APP_FILTER (
            set "FOUND_PATH=!LINE!"
            set "FOUND_ID=%%~nL"
            goto :dac_found
        ) else (
            echo(!LINE!| findstr /I "\\!APP_FILTER!\\Recent Templates" >nul
            if not errorlevel 1 (
                set "FOUND_PATH=!LINE!"
                set "FOUND_ID=%%~nL"
                goto :dac_found
            )
        )
    )
)

:dac_not_found
for %%# in (1) do (
    endlocal
    if not "%TARGET_ID%"=="" set "%TARGET_ID%="
    if not "%TARGET_PATH%"=="" set "%TARGET_PATH%="
    exit /b 1
)

:dac_found
for %%# in (1) do (
    endlocal
    if not "%TARGET_ID%"=="" set "%TARGET_ID%=%FOUND_ID%"
    if not "%TARGET_PATH%"=="" set "%TARGET_PATH%=%FOUND_PATH%"
    exit /b 0
)

:DetectMRUPath
rem Args: APP_NAME
setlocal enabledelayedexpansion
set "APP_NAME=%~1"
call :ResolveAppProperties "!APP_NAME!"
if not defined PROP_REG_NAME (
    endlocal
    exit /b 1
)
set "MRU_VAR=!PROP_MRU_VAR!"
set "MRU_PATH="
set "MRU_CONTAINER_PATH="

call :DetectAdalContainer MRU_CONTAINER_ID MRU_CONTAINER_PATH "!PROP_REG_NAME!"
if not errorlevel 1 if defined MRU_CONTAINER_PATH (
    set "MRU_PATH=!MRU_CONTAINER_PATH!\File MRU"
)

for %%V in (16.0 15.0 14.0 12.0) do (
    if not defined MRU_PATH (
        set "BASE=HKCU\Software\Microsoft\Office\%%V\!PROP_REG_NAME!\Recent Templates"
        for /f "delims=" %%K in ('reg query "!BASE!" /s /v "File MRU" 2^>nul ^| findstr /I "HKEY_CURRENT_USER"') do (
            set "MRU_PATH=%%K\File MRU"
            goto :found
        )
    )
)
:found
if not defined MRU_PATH (
    set "MRU_PATH=HKCU\Software\Microsoft\Office\16.0\!PROP_REG_NAME!\Recent Templates\File MRU"
)
endlocal & set "%MRU_VAR%=%MRU_PATH%"
exit /b


:SimulateRegEntry

rem Args: APP_NAME FILE_NAME FULL_PATH LOG_FILE
setlocal enabledelayedexpansion
set "APP_NAME=%~1"
set "FILE_NAME=%~2"
set "FULL_PATH=%~3"
set "LOG_FILE=%~4"

call :ResolveAppProperties "!APP_NAME!"
if not defined PROP_REG_NAME (
    endlocal
    exit /b 1
)
set "MRU_VAR=!PROP_MRU_VAR!"
set "SCRIPT_DIR=%~dp0"
set "COUNTER_VAR=!PROP_COUNTER_VAR!"
set "LOCAL_LOGGING=true"

if /I "%IsDesignModeEnabled%"=="false" set "LOCAL_LOGGING=false"
for /f "tokens=2 delims==" %%V in ('set !MRU_VAR! 2^>nul') do set "MRU_PATH=%%V"
set "MRU_CONTAINER_PATH="
set "MRU_CONTAINER_ID="

call :DetectAdalContainer MRU_CONTAINER_ID MRU_CONTAINER_PATH "!PROP_REG_NAME!"
if not errorlevel 1 if defined MRU_CONTAINER_PATH set "MRU_PATH=!MRU_CONTAINER_PATH!\File MRU"
if not defined MRU_PATH (
    call :DetectMRUPath "!APP_NAME!"
    for /f "tokens=2 delims==" %%V in ('set !MRU_VAR! 2^>nul') do set "MRU_PATH=%%V"
)

if not defined MRU_PATH (
    set "MRU_PATH=HKCU\Software\Microsoft\Office\16.0\!PROP_REG_NAME!\Recent Templates\File MRU"
)

set "%MRU_VAR%=%MRU_PATH%"

call :ShiftMRUEntries "!PROP_REG_NAME!" "!MRU_PATH!" "%IsDesignModeEnabled%" "!LOCAL_LOGGING!" "%LOG_FILE%"
for /f "tokens=2 delims==" %%C in ('set !COUNTER_VAR! 2^>nul') do set "CURRENT_COUNT=%%C"
if not defined CURRENT_COUNT set "CURRENT_COUNT=0"
set /a LOCAL_COUNT=CURRENT_COUNT+1
set "REG_VALUE=Item 1"
set "REG_DATA=[F00000000][T01ED6D7E58D00000][O00000000]*%FULL_PATH%"
if /I "%IsDesignModeEnabled%"=="true" echo [DEBUG] Escribiendo !REG_VALUE! en "!MRU_PATH!"
reg add "!MRU_PATH!" /v "!REG_VALUE!" /t REG_SZ /d "!REG_DATA!" /f >nul 2>&1
if errorlevel 1 (
    if /I "!LOCAL_LOGGING!"=="true" if defined LOG_FILE echo [ERROR] Falló al escribir !REG_VALUE! >> "%LOG_FILE%"
    if /I "%IsDesignModeEnabled%"=="true" echo [ERROR] Falló al escribir !REG_VALUE!
) else (
    if /I "%IsDesignModeEnabled%"=="true" echo [OK] !REG_VALUE! agregado correctamente
)
for %%N in ("!FILE_NAME!") do set "BASENAME=%%~nN"
set "META_VALUE=Item Metadata 1"
set "META_DATA=<Metadata><AppSpecific><id>%FULL_PATH%</id><nm>!BASENAME!</nm><du>%FULL_PATH%</du></AppSpecific></Metadata>"
if /I "%IsDesignModeEnabled%"=="true" echo [DEBUG] Escribiendo !META_VALUE! en "!MRU_PATH!"
reg add "!MRU_PATH!" /v "!META_VALUE!" /t REG_SZ /d "!META_DATA!" /f >nul 2>&1
if errorlevel 1 (
    if /I "!LOCAL_LOGGING!"=="true" if defined LOG_FILE echo [ERROR] Falló al escribir !META_VALUE! >> "%LOG_FILE%"
    if /I "%IsDesignModeEnabled%"=="true" echo [ERROR] Falló al escribir !META_VALUE!
) else (
    if /I "%IsDesignModeEnabled%"=="true" echo [OK] !META_VALUE! agregado correctamente
)
if /I "!LOCAL_LOGGING!"=="true" (
  if defined LOG_FILE (
    (
      echo [REG ENTRY]
      echo REG ADD "!MRU_PATH!" /v "!REG_VALUE!" /t REG_SZ /d "!REG_DATA!" /f
      echo REG ADD "!MRU_PATH!" /v "!META_VALUE!" /t REG_SZ /d "!META_DATA!" /f
      echo [INFO] Archivo: "!FILE_NAME!"
      echo.
    ) >> "%LOG_FILE%"
  )
)
endlocal & set "%COUNTER_VAR%=%LOCAL_COUNT%"
exit /b

:ResolveAppProperties
rem Internal helper. Args: APP_NAME
set "APP_UP=%~1"
if /I "%APP_UP%"=="WORD" (
    set "PROP_REG_NAME=Word"
    set "PROP_MRU_VAR=WORD_MRU_PATH"
    set "PROP_COUNTER_VAR=GLOBAL_ITEM_COUNT_WORD"
) else if /I "%APP_UP%"=="POWERPOINT" (
    set "PROP_REG_NAME=PowerPoint"
    set "PROP_MRU_VAR=PPT_MRU_PATH"
    set "PROP_COUNTER_VAR=GLOBAL_ITEM_COUNT"
) else if /I "%APP_UP%"=="EXCEL" (
    set "PROP_REG_NAME=Excel"
    set "PROP_MRU_VAR=EXCEL_MRU_PATH"
    set "PROP_COUNTER_VAR=GLOBAL_ITEM_COUNT_EXCEL"
) else (
    set "PROP_REG_NAME="
    set "PROP_MRU_VAR="
    set "PROP_COUNTER_VAR="
)
exit /b

:ShiftMRUEntries
rem Args: APP_KEY MRU_PATH DESIGN_MODE LOCAL_LOGGING LOG_FILE
setlocal EnableDelayedExpansion
set "APP_KEY=%~1"
set "TARGET_MRU=%~2"
set "DESIGN_MODE=%~3"
set "LOCAL_LOGGING=%~4"
set "LOG_FILE=%~5"
set "OFFSET=1"

if not defined TARGET_MRU (
    endlocal
    exit /b 0
)

if /I "%DESIGN_MODE%"=="true" echo [DEBUG] Ajustando índices MRU para %APP_KEY%...
if /I "%LOCAL_LOGGING%"=="true" if defined LOG_FILE echo [DEBUG] Ajustando índices MRU para %APP_KEY% >> "%LOG_FILE%"

set "TMP_FILE=%TEMP%\mru_shift_%RANDOM%.txt"
if exist "!TMP_FILE!" del "!TMP_FILE!" >nul 2>&1

set "FOUND_VALUES="

for /f "skip=2 tokens=* delims=" %%L in ('reg query "!TARGET_MRU!" 2^>nul') do (
    set "LINE=%%L"
    if not "!LINE!"=="" (
        set "HASREG=!LINE:REG_SZ=!"
        if not "!HASREG!"=="!LINE!" (
            set "WORK_LINE=!LINE:REG_SZ=|!"
            for /f "tokens=1 delims=|" %%P in ("!WORK_LINE!") do set "VALUE_NAME_RAW=%%P"
            call :TrimWhitespaceVar VALUE_NAME_RAW
            if defined VALUE_NAME_RAW (
                set "FIRST="
                set "SECOND="
                set "THIRD="
                for /f "tokens=1-3" %%a in ("!VALUE_NAME_RAW!") do (
                    if not defined FIRST set "FIRST=%%a"
                    if not defined SECOND set "SECOND=%%b"
                    if not defined THIRD set "THIRD=%%c"
                )
                set "BASE="
                set "INDEX="
                if /I "!FIRST!"=="Item" (
                    if /I "!SECOND!"=="Metadata" (
                        set "BASE=Item Metadata"
                        set "INDEX=!THIRD!"
                    ) else (
                        set "BASE=Item"
                        set "INDEX=!SECOND!"
                    )
                )
                if defined INDEX (
                    echo(!INDEX!| findstr /R "^[0-9][0-9]*$" >nul
                    if not errorlevel 1 (
                        set "FOUND_VALUES=1"
                        set "PAD=0000000000!INDEX!"
                        set "PAD=!PAD:~-10!"
                        >>"!TMP_FILE!" echo(!PAD!^|!VALUE_NAME_RAW!
                    )
                )
            )
        )
    )
)

if not defined FOUND_VALUES (
    if /I "%DESIGN_MODE%"=="true" echo [DEBUG] No se encontraron entradas MRU previas para %APP_KEY%.
    if /I "%LOCAL_LOGGING%"=="true" if defined LOG_FILE echo [DEBUG] Sin entradas MRU para desplazar en %APP_KEY% >> "%LOG_FILE%"
    if exist "!TMP_FILE!" del "!TMP_FILE!" >nul 2>&1
    endlocal
    exit /b 0
)

for /f "usebackq tokens=1* delims=|" %%A in (`sort /R "!TMP_FILE!"`) do (
    call :ShiftMRURename "%%B" "%OFFSET%" "!TARGET_MRU!" "%DESIGN_MODE%" "%LOCAL_LOGGING%" "%LOG_FILE%" "%APP_KEY%"
)

if exist "!TMP_FILE!" del "!TMP_FILE!" >nul 2>&1

if /I "%DESIGN_MODE%"=="true" echo [DEBUG] Reindexado MRU completado para %APP_KEY%.
if /I "%LOCAL_LOGGING%"=="true" if defined LOG_FILE echo [DEBUG] Reindexado MRU completado para %APP_KEY% >> "%LOG_FILE%"

endlocal
exit /b 0

:ShiftMRURename
rem Args: ORIGINAL_NAME OFFSET MRU_PATH DESIGN_MODE LOCAL_LOGGING LOG_FILE APP_KEY
setlocal EnableDelayedExpansion
set "ORIGINAL_NAME=%~1"
set "OFFSET=%~2"
set "MRU_PATH=%~3"
set "DESIGN_MODE=%~4"
set "LOCAL_LOGGING=%~5"
set "LOG_FILE=%~6"
set "APP_KEY=%~7"

if "%ORIGINAL_NAME%"=="" (
    endlocal
    exit /b 0
)

set "FIRST="
set "SECOND="
set "THIRD="
for /f "tokens=1-3" %%a in ("!ORIGINAL_NAME!") do (
    if not defined FIRST set "FIRST=%%a"
    if not defined SECOND set "SECOND=%%b"
    if not defined THIRD set "THIRD=%%c"
)

set "BASE="
set "INDEX="
if /I "!FIRST!"=="Item" (
    if /I "!SECOND!"=="Metadata" (
        set "BASE=Item Metadata"
        set "INDEX=!THIRD!"
    ) else (
        set "BASE=Item"
        set "INDEX=!SECOND!"
    )
)

if not defined INDEX (
    endlocal
    exit /b 0
)

set /a NEW_INDEX=INDEX+OFFSET
if /I "!BASE!"=="Item Metadata" (
    set "NEW_NAME=Item Metadata !NEW_INDEX!"
) else (
    set "NEW_NAME=Item !NEW_INDEX!"
)

set "DATA_LINE="
for /f "skip=2 tokens=* delims=" %%L in ('reg query "!MRU_PATH!" /v "!ORIGINAL_NAME!" 2^>nul') do set "DATA_LINE=%%L"
if not defined DATA_LINE (
    endlocal
    exit /b 0
)

set "DATA_LINE=!DATA_LINE:*REG_SZ=!"
call :TrimWhitespaceVar DATA_LINE
set "DATA=!DATA_LINE!"

if /I "%DESIGN_MODE%"=="true" echo [DEBUG] Renombrando "!ORIGINAL_NAME!" a "!NEW_NAME!" en "!MRU_PATH!" para %APP_KEY%.
if /I "%LOCAL_LOGGING%"=="true" if defined LOG_FILE echo [DEBUG] Renombrando !ORIGINAL_NAME! a !NEW_NAME! en !MRU_PATH! >> "%LOG_FILE%"

reg add "!MRU_PATH!" /v "!NEW_NAME!" /t REG_SZ /d "!DATA!" /f >nul
reg delete "!MRU_PATH!" /v "!ORIGINAL_NAME!" /f >nul

endlocal
exit /b 0

:TrimWhitespaceVar
rem Args: VAR_NAME
setlocal EnableDelayedExpansion
set "VALUE=!%~1!"
:TrimLeadingWS
if defined VALUE if "!VALUE:~0,1!"==" " (
    set "VALUE=!VALUE:~1!"
    goto :TrimLeadingWS
)
:TrimTrailingWS
if defined VALUE if "!VALUE:~-1!"==" " (
    set "VALUE=!VALUE:~0,-1!"
    goto :TrimTrailingWS
)
endlocal & set "%~1=%VALUE%"
exit /b 0

:EndOfScript
endlocal
exit /b
