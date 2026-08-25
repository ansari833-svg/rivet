@echo off
setlocal EnableExtensions
rem ============================================================
rem  RunNPVBridge.bat
rem  Imports modNPVBridge.bas into Excel and runs BuildNPVBridge.
rem
rem  Behaviour:
rem    - If an Excel workbook is already open, the module is imported
rem      into the ACTIVE workbook and the macro is run there.
rem    - Otherwise a new workbook is created for you.
rem
rem  Requirements (one-time, in Excel):
rem    File > Options > Trust Center > Trust Center Settings >
rem    Macro Settings > tick "Trust access to the VBA project object
rem    model".  Without this the import step is blocked by Excel.
rem
rem  Usage:   double-click, or   RunNPVBridge.bat  [path\to\module.bas]
rem ============================================================

rem --- Locate the .bas module (arg 1, else next to this .bat) ---
set "BAS=%~1"
if "%BAS%"=="" set "BAS=%~dp0modNPVBridge.bas"

if not exist "%BAS%" (
    echo(
    echo ERROR: Could not find the VBA module:
    echo        "%BAS%"
    echo Place modNPVBridge.bas next to this .bat, or pass its path as an argument.
    echo(
    pause
    exit /b 1
)

rem --- Build a temporary VBScript driver ---
set "VBS=%TEMP%\_npvbridge_%RANDOM%.vbs"

>  "%VBS%" echo Option Explicit
>> "%VBS%" echo Dim xl, wb, basPath, comp, m, alreadyRunning
>> "%VBS%" echo basPath = WScript.Arguments(0)
>> "%VBS%" echo On Error Resume Next
>> "%VBS%" echo Set xl = GetObject(, "Excel.Application")
>> "%VBS%" echo alreadyRunning = (Err.Number = 0 And Not xl Is Nothing)
>> "%VBS%" echo On Error GoTo 0
>> "%VBS%" echo If xl Is Nothing Then Set xl = CreateObject("Excel.Application")
>> "%VBS%" echo xl.Visible = True
>> "%VBS%" echo If xl.Workbooks.Count = 0 Then
>> "%VBS%" echo   Set wb = xl.Workbooks.Add
>> "%VBS%" echo Else
>> "%VBS%" echo   Set wb = xl.ActiveWorkbook
>> "%VBS%" echo End If
>> "%VBS%" echo On Error Resume Next
>> "%VBS%" echo For Each comp In wb.VBProject.VBComponents
>> "%VBS%" echo   If comp.Name = "modNPVBridge" Then wb.VBProject.VBComponents.Remove comp
>> "%VBS%" echo Next
>> "%VBS%" echo If Err.Number ^<^> 0 Then
>> "%VBS%" echo   MsgBox "Cannot access the VBA project." ^& vbCrLf ^& vbCrLf ^& _
>> "%VBS%" echo     "Enable: File ^> Options ^> Trust Center ^> Trust Center Settings ^> " ^& _
>> "%VBS%" echo     "Macro Settings ^> 'Trust access to the VBA project object model', " ^& _
>> "%VBS%" echo     "then run this again.", vbExclamation, "NPV Bridge"
>> "%VBS%" echo   WScript.Quit 1
>> "%VBS%" echo End If
>> "%VBS%" echo On Error GoTo 0
>> "%VBS%" echo Set m = wb.VBProject.VBComponents.Import(basPath)
>> "%VBS%" echo xl.Run "BuildNPVBridge"

rem --- Run it ---
echo Launching Excel and running BuildNPVBridge...
cscript //nologo "%VBS%" "%BAS%"
set "RC=%ERRORLEVEL%"

del "%VBS%" >nul 2>&1

if not "%RC%"=="0" (
    echo(
    echo The macro did not complete successfully. See any message shown by Excel.
    pause
)

endlocal
exit /b %RC%
