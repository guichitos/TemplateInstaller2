@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul

rem ===========================================================
rem === UNIVERSAL OFFICE TEMPLATE UNINSTALLER (v1.2) ==========
rem -----------------------------------------------------------
rem Uses the same hardcoded base paths as main_installer.bat
rem to remove the default XML-based templates (.dotx/.potx/.xltx)
rem and their macro-enabled counterparts (.dotm/.potm/.xltm),
rem restoring backups if available.
rem ===========================================================

rem === Mode and logging configuration ========================
rem Toggle this flag to control diagnostic output for this script only.
rem true  = verbose mode with console messages, logging, and final pause.
rem false = silent mode (no console output or pause).
set "IsInDesignModeEnabled=false"

set "BaseDirectoryPath=%~dp0"
set "LibraryDirectoryPath=%BaseDirectoryPath%lib"
set "LogsDirectoryPath=%BaseDirectoryPath%logs"
set "LogFilePath=%LogsDirectoryPath%\uninstall_log.txt"

if /I "%IsInDesignModeEnabled%"=="true" (
    if not exist "%LogsDirectoryPath%" mkdir "%LogsDirectoryPath%"
    echo [%DATE% %TIME%] --- START UNINSTALL --- > "%LogFilePath%"
    title OFFICE TEMPLATE UNINSTALLER - DEBUG MODE
    echo [DEBUG] Running from: %BaseDirectoryPath%
)

rem === Define base template paths (same as main_installer.bat) ===
set "WORD_PATH=%APPDATA%\Microsoft\Templates"
set "PPT_PATH=%APPDATA%\Microsoft\Templates"
set "EXCEL_PATH=%APPDATA%\Microsoft\Excel\XLSTART"

if /I "%IsInDesignModeEnabled%"=="true" (
    echo.
    echo [TARGET CLEANUP PATHS]
    echo ----------------------------
    echo WORD PATH:       %WORD_PATH%
    echo POWERPOINT PATH: %PPT_PATH%
    echo EXCEL PATH:      %EXCEL_PATH%
    echo ----------------------------
)

if /I "%IsInDesignModeEnabled%"=="true" (
    echo [INFO] --- TARGET CLEANUP PATHS --- >> "%LogFilePath%"
    echo Word path: %WORD_PATH% >> "%LogFilePath%"
    echo PowerPoint path: %PPT_PATH% >> "%LogFilePath%"
    echo Excel path: %EXCEL_PATH% >> "%LogFilePath%"
    echo ---------------------------- >> "%LogFilePath%"
)

rem === Define files ==========================================
set "WordFile=%WORD_PATH%\Normal.dotx"
set "WordBackup=%WORD_PATH%\Normal_backup.dotx"
set "WordMacroFile=%WORD_PATH%\Normal.dotm"
set "WordMacroBackup=%WORD_PATH%\Normal_backup.dotm"

set "PptFile=%PPT_PATH%\Blank.potx"
set "PptBackup=%PPT_PATH%\Blank_backup.potx"
set "PptMacroFile=%PPT_PATH%\Blank.potm"
set "PptMacroBackup=%PPT_PATH%\Blank_backup.potm"

set "ExcelBookFile=%EXCEL_PATH%\Book.xltx"
set "ExcelBookBackup=%EXCEL_PATH%\Book_backup.xltx"
set "ExcelBookMacroFile=%EXCEL_PATH%\Book.xltm"
set "ExcelBookMacroBackup=%EXCEL_PATH%\Book_backup.xltm"

set "ExcelSheetFile=%EXCEL_PATH%\Sheet.xltx"
set "ExcelSheetBackup=%EXCEL_PATH%\Sheet_backup.xltx"
set "ExcelSheetMacroFile=%EXCEL_PATH%\Sheet.xltm"
set "ExcelSheetMacroBackup=%EXCEL_PATH%\Sheet_backup.xltm"

rem === Folder existence check ================================
for %%D in ("%WORD_PATH%" "%PPT_PATH%" "%EXCEL_PATH%") do (
    if not exist "%%~D" (
        call :Log "%LogFilePath%" "[WARN] Missing folder: %%~D"
    )
)

rem === Helper routine: delete & restore =======================
call :ProcessFile "Word (.dotx)" "%WordFile%" "%WordBackup%" "%LogFilePath%"
call :ProcessFile "Word (.dotm)" "%WordMacroFile%" "%WordMacroBackup%" "%LogFilePath%"
call :ProcessFile "PowerPoint (.potx)" "%PptFile%" "%PptBackup%" "%LogFilePath%"
call :ProcessFile "PowerPoint (.potm)" "%PptMacroFile%" "%PptMacroBackup%" "%LogFilePath%"
call :ProcessFile "Excel Book (.xltx)" "%ExcelBookFile%" "%ExcelBookBackup%" "%LogFilePath%"
call :ProcessFile "Excel Book (.xltm)" "%ExcelBookMacroFile%" "%ExcelBookMacroBackup%" "%LogFilePath%"
call :ProcessFile "Excel Sheet (.xltx)" "%ExcelSheetFile%" "%ExcelSheetBackup%" "%LogFilePath%"
call :ProcessFile "Excel Sheet (.xltm)" "%ExcelSheetMacroFile%" "%ExcelSheetMacroBackup%" "%LogFilePath%"

call :Finalize "%LogFilePath%"

endlocal
exit /b


:ProcessFile
rem ===========================================================
rem Args: AppName, TargetFile, BackupFile, LogFile
rem ===========================================================
setlocal enabledelayedexpansion
set "AppName=%~1"
set "TargetFile=%~2"
set "BackupFile=%~3"
set "LogFile=%~4"

call :Log "%LogFile%" ""
call :Log "%LogFile%" "[INFO] Processing %AppName%..."

rem === Step 1: Always delete current template (factory reset) ===
if exist "%TargetFile%" (
    del /F /Q "%TargetFile%" >nul 2>&1
    if exist "%TargetFile%" (
        set "Message=[ERROR] Could not delete %TargetFile%. File may be locked."
        call :Log "%LogFile%" "!Message!"
    ) else (
        set "Message=[OK] Deleted %TargetFile%"
        call :Log "%LogFile%" "!Message!"
    )
) else (
    set "Message=[INFO] %TargetFile% not found."
    call :Log "%LogFile%" "!Message!"
)

rem === Step 2: Restore from backup if available ===
if exist "%BackupFile%" (
    copy /Y "%BackupFile%" "%TargetFile%" >nul 2>&1
    if exist "%TargetFile%" (
        del /F /Q "%BackupFile%" >nul 2>&1
        if exist "%BackupFile%" (
            set "Message=[WARN] Restored %TargetFile% but could not delete backup."
            call :Log "%LogFile%" "!Message!"
        ) else (
            set "Message=[OK] Restored %TargetFile% and deleted backup."
            call :Log "%LogFile%" "!Message!"
        )
    ) else (
        set "Message=[ERROR] Backup copy failed for %AppName%."
        call :Log "%LogFile%" "!Message!"
    )
) else (
    rem === No backup found, ensure no template remains ===
    if exist "%TargetFile%" del /F /Q "%TargetFile%" >nul 2>&1
    if not exist "%TargetFile%" (
        set "Message=[OK] No backup found; folder left clean for %AppName%."
        call :Log "%LogFile%" "!Message!"
    ) else (
        set "Message=[ERROR] Could not clean template for %AppName%."
        call :Log "%LogFile%" "!Message!"
    )
)

endlocal
exit /b 0

:Finalize
setlocal enabledelayedexpansion
if /I not "%IsInDesignModeEnabled%"=="true" (
    endlocal
    exit /b 0
)

set "ResolvedLogPath=%~1"

>>"%~1" echo [%DATE% %TIME%] --- UNINSTALL COMPLETED ---
call :Log "%~1" ""
call :Log "%~1" "[FINAL] Uninstallation process finished successfully."
call :Log "%~1" "Log saved at: \"!ResolvedLogPath!\""
call :Log "%~1" "--------------------------------------------------------"

if /I "%IsInDesignModeEnabled%"=="true" (
    echo Presiona una tecla para cerrar esta ventana...
    pause
)

endlocal
exit /b 0

:Log
setlocal enabledelayedexpansion
set "LogFile=%~1"
set "LogMessage=%~2"
if /I "%IsInDesignModeEnabled%"=="true" (
    if defined LogFile if not "!LogFile!"=="" (
        if defined LogMessage (
            >>"!LogFile!" echo(!LogMessage!
        ) else (
            >>"!LogFile!" echo.
        )
    )
)
if /I "%IsInDesignModeEnabled%"=="true" (
    if defined LogMessage (
        echo(!LogMessage!
    ) else (
        echo.
    )
)
endlocal
exit /b 0
