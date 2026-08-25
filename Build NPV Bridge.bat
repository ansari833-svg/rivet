@echo off
setlocal EnableExtensions
rem ============================================================
rem  Build NPV Bridge.bat  --  drag-and-drop NPV Bridge builder
rem                            (runner-workbook design)
rem
rem  Drag an Excel workbook (.xlsx / .xlsm) onto this icon. The
rem  dropped workbook gets the NPV Bridge data table and chart, then
rem  stays open and unsaved so you can review and save it yourself.
rem
rem  HOW IT WORKS
rem    The BuildNPVBridge macro lives inside a separate, pre-built
rem    "runner" workbook (NPVBridge_Runner.xlsm) that sits next to this
rem    .bat. This launcher opens the runner read-only and calls its
rem    macro, passing the dropped workbook as the output target. Nothing
rem    is ever written into the runner's VBA project at runtime.
rem
rem    Because the macro already exists inside the runner, this design
rem    needs NO "Trust access to the VBA project object model" setting
rem    (that was the dependency of the old runtime-injection version,
rem    which built the module on the fly). AutomationSecurityLow lets the
rem    runner's macros run without an "Enable Content" prompt.
rem
rem  ONE-TIME SETUP (build the runner workbook once, then forever):
rem    1. Open a blank Excel workbook.
rem    2. Press Alt+F11 to open the VBA editor.
rem    3. File > Import File > select modNPVBridge.bas.
rem    4. File > Save As > name it NPVBridge_Runner.xlsm, type = Excel
rem       Macro-Enabled Workbook (.xlsm).
rem    5. Put NPVBridge_Runner.xlsm in the same folder as this .bat.
rem       Done -- forever.
rem ============================================================

rem --- 1. Require a dropped workbook (path arrives as %1) ---
if "%~1"=="" (
    echo Usage: drag an Excel workbook ^(.xlsx or .xlsm^) onto this .bat,
    echo        or run:  "%~nx0" "C:\path\to\workbook.xlsx"
    echo(
    pause
    exit /b 1
)

rem --- 2. Locate the runner workbook RELATIVE TO THIS .BAT ---
set "RUNNER=%~dp0NPVBridge_Runner.xlsm"
if not exist "%RUNNER%" (
    echo NPVBridge_Runner.xlsm was not found next to this .bat:
    echo     "%RUNNER%"
    echo(
    echo One-time setup to create it:
    echo   1. Open a blank Excel workbook.
    echo   2. Press Alt+F11 to open the VBA editor.
    echo   3. File ^> Import File ^> select modNPVBridge.bas.
    echo   4. File ^> Save As ^> name it NPVBridge_Runner.xlsm, type =
    echo      Excel Macro-Enabled Workbook ^(.xlsm^).
    echo   5. Put NPVBridge_Runner.xlsm in the same folder as this .bat.
    echo      Done -- forever.
    echo(
    pause
    exit /b 1
)

set "WB=%~1"
set "VBS=%TEMP%\npvbridge_%RANDOM%%RANDOM%.vbs"

rem --- 3. Write the temporary .vbs driver ---
if exist "%VBS%" del "%VBS%"
>  "%VBS%" echo Option Explicit
>> "%VBS%" echo Dim xl, runner, wb, runnerPath, wbPath
>> "%VBS%" echo runnerPath = WScript.Arguments(0)
>> "%VBS%" echo wbPath     = WScript.Arguments(1)
>> "%VBS%" echo Set xl = CreateObject("Excel.Application")
>> "%VBS%" echo xl.Visible = True
>> "%VBS%" echo ' msoAutomationSecurityLow = 1: run the runner's macros with no
>> "%VBS%" echo ' "Enable Content" prompt. Must be set BEFORE opening the runner.
>> "%VBS%" echo xl.AutomationSecurity = 1
>> "%VBS%" echo ' Open the runner read-only (we never save it) ...
>> "%VBS%" echo Set runner = xl.Workbooks.Open(runnerPath, , True)
>> "%VBS%" echo ' ... and the dropped workbook (this is the output target).
>> "%VBS%" echo Set wb = xl.Workbooks.Open(wbPath)
>> "%VBS%" echo ' Run the macro inside the runner, passing the dropped workbook.
>> "%VBS%" echo xl.Run "NPVBridge_Runner.xlsm!BuildNPVBridge", wb
>> "%VBS%" echo ' Close the runner without saving; leave the dropped workbook open.
>> "%VBS%" echo runner.Close False
>> "%VBS%" echo ' Do NOT save wb, do NOT close wb, do NOT xl.Quit -- leave it for review.

rem --- 4. Run it (runner path + dropped workbook path, both quoted) ---
cscript //nologo "%VBS%" "%RUNNER%" "%WB%"
set "RC=%ERRORLEVEL%"

rem --- 5. Clean up the temp .vbs ---
del "%VBS%" >nul 2>&1

if not "%RC%"=="0" (
    echo(
    echo The macro did not complete successfully. See any message shown by Excel.
    pause
)

endlocal
exit /b %RC%
