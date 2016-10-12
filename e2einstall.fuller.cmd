:: //***************************************************************************
:: //
:: // File:      e2einstall.fuller.cmd
:: //
:: // Additional files required:  None.  Script creates required elevate.cmd and 
:: //                             elevate.vbs in %Temp% when run.
:: //
:: // Purpose:   takes events and emails them
:: //            
:: //
:: // Usage:     e2einstall.fuller.cmd 
:: //
:: // Version:   0.8
:: //
:: // History:
:: // 0.8   16.10.12 First Github version, the backups password is now prompted for, if we retain this script I can encrypt the password to be unlocked with
:: //                a master key but it's probably not worth the effort of redesigning the whole thing for that.
:: // 0.7   15.06.17 Fixed all the passwords so they use !pwd! so that ! can be escaped by using caret(^! = !) added company name to test email subject line.
:: // 0.6   15.03.03 Bah, previous version borked, changed subject line to correct it.
:: // 0.6   15.02.23 Added CloudBackup to monitor Azure backup jobs, added build 6 to subject line of emails
:: // 0.5   14.12.05 Fuller default fields pre-defined
:: // 0.4   14.10.02 Added /RL Highest to schtasks to fix Windows 2008hanged dialogs to explain need to escape some characters
:: // 0.3   14.08.19 Changed dialogs to explain need to escape some characters
:: // 0.2   14.04.25 Fixed Crash Eventlog, removed false positive event id 2.

:: //
:: // ***** End Header *****
:: //***************************************************************************

@echo off
setlocal enabledelayedexpansion

set CmdDir=%~dp0
set CmdDir=%CmdDir:~0,-1%


:: ////////////////////////////////////////////////////////////////////////////
:: Check whether running elevated
:: ////////////////////////////////////////////////////////////////////////////
call :CREATE_ELEVATE_SCRIPTS

:: Check for Mandatory Label\High Mandatory Level
whoami /groups | find "S-1-16-12288" > nul
if "%errorlevel%"=="0" (
    echo Running as elevated user.  Continuing script.
) else (
    echo Not running as elevated user.
    echo Relaunching Elevated: "%~dpnx0" %*

    if exist "%Temp%\elevate.cmd" (
        set ELEVATE_COMMAND="%Temp%\elevate.cmd"
    ) else (
        set ELEVATE_COMMAND=elevate.cmd
    )

    set CARET=^^
    !ELEVATE_COMMAND! cmd /k cd /d "%~dp0" !CARET!^& call "%~dpnx0" %*
    goto :EOF
)

if exist %ELEVATE_CMD% del %ELEVATE_CMD%
if exist %ELEVATE_VBS% del %ELEVATE_VBS%


:: ////////////////////////////////////////////////////////////////////////////
:: Main script code starts here
:: ////////////////////////////////////////////////////////////////////////////
@echo off
SETLOCAL ENABLEEXTENSIONS

echo Phil's Evt2Eml Installer
copy /y blat.exe %windir%
if not exist %windir%\blat.exe goto blatmissing
if not exist "%programfiles(x86)%\stunnel\stunnel.exe" goto stunnelinst
:stunnelcomplete

::Cleanup prior runs
if exist %temp%\xquery.txt del /q %temp%\xquery.txt
if exist %temp%\bquery.txt del /q %temp%\bquery.txt
if exist %temp%\iquery.txt del /q %temp%\iquery.txt
if exist %temp%\fquery.txt del /q %temp%\fquery.txt
if exist %windir%\evt2emlcrash.cmd del /q %windir%\evt2emlcrash.cmd
if exist %windir%\evt2emlcrash.cmd del /q %windir%\evt2emlcrash.bat
if exist %windir%\evt2emlbackup.cmd del /q %windir%\evt2emlbackup.cmd
if exist %windir%\evt2emlbackup.cmd del /q %windir%\evt2emlbackup.bat
if exist %windir%\evt2emlintel.cmd del /q %windir%\evt2emlintel.cmd
if exist %windir%\evt2emlintel.cmd del /q %windir%\evt2emlintel.bat
if exist %windir%\evt2emlfuture.cmd del /q %windir%\evt2emlfuture.cmd
if exist %windir%\evt2emlfuture.cmd del /q %windir%\evt2emlfuture.bat
schtasks /Delete /TN "Evt2Eml Crash" /F
schtasks /Delete /TN "Evt2Eml Backup" /F
schtasks /Delete /TN "Evt2EmlBackup" /F
schtasks /Delete /TN "Evt2EmlRaid" /F
schtasks /Delete /TN "Evt2Eml Intel" /F
schtasks /Delete /TN "Evt2Eml Future" /F
::Debug pause
::pause
cls

::Create the VBS script with an echo statement:
ECHO Wscript.Echo Inputbox("If password contains special characters you must escape the characters with caret. See http://adf.ly/rIYAz", "Enter Windows Password: ")>%TEMP%\~input.vbs
FOR /f "delims=/" %%G IN ('cscript //nologo %TEMP%\~input.vbs') DO set pwd=%%G
DEL %TEMP%\~input.vbs

::Below is the manual override password in case a forward slash or ! breaks the input field
:: % must be %%, ! must be ^!, and ^ must be ^^, 
::set pwd=Example^!Example^^Example^/Example%%Example

::ECHO Wscript.Echo Inputbox("Example: jdoe@contoso.com", "Enter Email")>%TEMP%\~input.vbs
::FOR /f "delims=/" %%G IN ('cscript //nologo %TEMP%\~input.vbs') DO set eml=%%G
::DEL %TEMP%\~input.vbs
set eml=customerbackup@fullercomputer.com

::ECHO Wscript.Echo Inputbox("Enter Email Server")>%TEMP%\~input.vbs
::FOR /f "delims=/" %%G IN ('cscript //nologo %TEMP%\~input.vbs') DO set srv=%%G
::DEL %TEMP%\~input.vbs
set srv=127.0.0.1

::ECHO Wscript.Echo Inputbox("Enter Email Server Port")>%TEMP%\~input.vbs
::FOR /f "delims=/" %%G IN ('cscript //nologo %TEMP%\~input.vbs') DO set port=%%G
::DEL %TEMP%\~input.vbs
set port=2501

ECHO Wscript.Echo Inputbox("If password contains special characters you must escape the characters with caret. See http://adf.ly/rIYAz", "Enter Backups Email Password: ")>%TEMP%\~input.vbs
FOR /f "delims=/" %%G IN ('cscript //nologo %TEMP%\~input.vbs') DO set epass=%%G
DEL %TEMP%\~input.vbs

ECHO Wscript.Echo Inputbox("No spaces allowed. Example: PremiumWidgetCo", "Enter Company name: ")>%TEMP%\~input.vbs
FOR /f "delims=/" %%G IN ('cscript //nologo %TEMP%\~input.vbs') DO set company=%%G
DEL %TEMP%\~input.vbs

:menuLOOP
cls

::This is here to avoid password botching
SETLOCAL ENABLEDELAYEDEXPANSION
echo PWD !pwd!
echo EML %eml%
echo SRV %srv%
echo PORT %port%
echo UID %company%
echo EPASS !epass!

::debug pause
::pause

for /f "tokens=1,2,* delims=_ " %%A in ('"findstr /b /c:":menu_" "%~f0""') do echo.  %%B  %%C
set choice=
echo.&set /p choice=Selection(Q to quit): ||GOTO:EOF
echo.&call:menu_%choice%
GOTO:menuLOOP

:menu_M   Mail Test
ECHO ON
blat.exe -to backups@fullercomputer.com -u %eml% -pw !epass! -f %eml% -server %srv% -port %port% -subject "Test email v8 %company%" -priority 1 -log %windir%\blat.log -body "Email Config Success"
ECHO OFF
pause
GOTO:menuLOOP

:menu_T   Test Events
::eventcreate /ID 
eventcreate /ID 1 /L APPLICATION /T ERROR  /SO "Windows Test Backup" /D "evt2eml Windows Backup Trigger Test Event"
eventcreate /ID 2 /L SYSTEM /T ERROR  /SO "Windows Crash" /D "evt2eml Windows Crash Trigger Test Event"
eventcreate /ID 3 /L SYSTEM /T ERROR  /SO "iastor" /D "evt2eml Windows RAID Trigger Test Event"
eventcreate /ID 4 /L APPLICATION /T ERROR  /SO "Windows Future" /D "evt2eml Windows Future Trigger Test Event"
pause
GOTO:menuLOOP

:menu_X   Crash
goto xskip 
::Setup evt2emlcrash.cmd 
::This section is not processed in e2einstall.cmd but used to generate an evt2emlcrash.cmd 
%XID%@Echo off 
%XID%
%XID%Echo "Please wait, gathering log info to email to Fuller Computer"
%XID%del %temp%\xquery.txt /q
%XID%
%XID%::Crash
%XID%wevtutil qe System "/q:*[System[(EventID=6008) and TimeCreated[timediff(@SystemTime) <= 57600000]]]" /f:text /rd:true > %temp%\xquery.txt 
%XID%wevtutil qe System "/q:*[System[Provider[@Name='Windows Crash'] and (Level=1 or Level=2 or Level=3) and TimeCreated[timediff(@SystemTime) <= 57600000]]]" /f:text /rd:true >> %temp%\xquery.txt
%XID%
%XID%set "dir=%temp%"
%XID%set "file=%dir%\xquery.txt"
%XID%echo %file% %dir%
%XID%for /f %%i in ("%file%") do set size=%%~zi
%XID%if %size% gtr 1 blat.exe %temp%\xquery.txt -to backups@fullercomputer.com -u %1 -pw %4 -f %1 -server %2 -port %3 -subject "Crash Alert v7 %5" -priority 1 -log %windir%\blat.log
:xskip 
::This is the line to create the evt2emlcrash.cmd file, this is a hack of a mess but the only way to do it in a single batchfile. 
type %~dp0e2einstall.fuller.cmd | find "%%XID%%"| find /v "BUT NOT THIS LINE!" > %windir%\evt2emlcrash.cmd 
ECHO ON
schtasks /create /tn "Evt2Eml Crash" /tr "evt2emlcrash.cmd %eml% %srv% %port% !epass! \"\"%company%\"\"" /delay 0010:00 /SC onlogon /RU %userdomain%\%username% /RP !pwd! /F /RL Highest
pause
ECHO OFF
GOTO:menuLOOP

:menu_B   Backup 
schtasks /create /tn "Evt2Eml Backup" /tr evt2emlbackup.cmd /st 08:00  /SC Daily /RU %userdomain%\%username% /RP !pwd! /RL Highest
goto bskip 
::Setup evt2emlbackup.cmd 
::This section is not processed in e2einstall.cmd but used to generate an evt2emlbackup.cmd 
%BID%@Echo off 
%BID%
%BID%Echo "Please wait, gathering log info to email to Fuller Computer"
%BID%del %temp%\bquery.txt /q
%BID%
%BID%::Backup
%BID%wevtutil qe Microsoft-Windows-Backup "/q:*[System[(Level=1 or Level =2 or Level =3) and TimeCreated[timediff(@SystemTime) <= 86400000]]]" /f:text /rd:true > %temp%\bquery.txt
%BID%wevtutil qe CloudBackup "/q:*[System[(Level=1 or Level =2 or Level =3) and TimeCreated[timediff(@SystemTime) <= 86400000]]]" /f:text /rd:true >> %temp%\bquery.txt
%BID%wevtutil qe Application "/q:*[System[Provider[@Name='Windows Backup'] and (Level=1 or Level=2 or Level=3) and TimeCreated[timediff(@SystemTime) <= 86400000]]]" /f:text /rd:true >> %temp%\bquery.txt
%BID%wevtutil qe Application "/q:*[System[Provider[@Name='Windows Test Backup'] and (Level=1 or Level=2 or Level=3) and TimeCreated[timediff(@SystemTime) <= 86400000]]]" /f:text /rd:true >> %temp%\bquery.txt
%BID%
%BID%set "dir=%temp%"
%BID%set "file=%dir%\bquery.txt"
%BID%echo %file% %dir%
%BID%for /f %%i in ("%file%") do set size=%%~zi
%BID%if %size% gtr 1 blat.exe %temp%\bquery.txt -to backups@fullercomputer.com -u %1 -pw %4 -f %1 -server %2 -port %3 -subject "Backup Alert v7 %5" -priority 1 -log %windir%\blat.log
:bskip 
::This is the line to create the evt2emlbackup.cmd file, this is a hack of a mess but the only way to do it in a single batchfile. 
type %~dp0e2einstall.fuller.cmd | find "%%BID%%"| find /v "BUT NOT THIS LINE!" > %windir%\evt2emlbackup.cmd 
ECHO ON
schtasks /create /tn "Evt2Eml Backup" /tr "evt2emlbackup.cmd %eml% %srv% %port% !epass! \"\"%company%\"\"" /st 08:00  /SC Daily /RU %userdomain%\%username% /RP !pwd! /F /RL Highest
pause
ECHO OFF
GOTO:menuLOOP

:menu_I   Intel Matrix/RST

goto iskip 
::Setup evt2emlintel.cmd 
::This section is not processed in e2einstall.cmd but used to generate an evt2emlintel.cmd 
%IID%@Echo off 
%IID%
%IID%Echo "Please wait, gathering log info to email to Fuller Computer"
%IID%del %temp%\iquery.txt /q
%IID%
%IID%::Intel
%IID%wevtutil qe System "/q:*[System[Provider[@Name='iaStor' or @Name='IAStorDataMgrSvc' or @Name='iaStorV'] and (Level=1  or Level=2 or Level=3 or Level=5) and TimeCreated[timediff(@SystemTime) <= 3660000]]]" /f:text /rd:true > %temp%\iquery.txt 
%IID%
%IID%set "dir=%temp%"
%IID%set "file=%dir%\iquery.txt"
%IID%echo %file% %dir%
%IID%for /f %%i in ("%file%") do set size=%%~zi
%IID%if %size% gtr 1 blat.exe %temp%\iquery.txt -to backups@fullercomputer.com -u %1 -pw %4 -f %1 -server %2 -port %3 -subject "Raid Alert v7 %5" -priority 1 -log %windir%\blat.log
:iskip 
::This is the line to create the evt2emlintel.cmd file, this is a hack of a mess but the only way to do it in a single batchfile. 
type %~dp0e2einstall.fuller.cmd | find "%%IID%%"| find /v "BUT NOT THIS LINE!" > %windir%\evt2emlintel.cmd 
ECHO ON
schtasks /create /tn "Evt2Eml Intel" /tr "evt2emlintel.cmd %eml% %srv% %port% !epass! \"\"%company%\"\"" /SC Hourly /RU %userdomain%\%username% /RP !pwd! /F /RL Highest
pause
ECHO OFF
GOTO:menuLOOP

:menu_F   Future Option

goto fskip 
::Setup emaillog.cmd 
::This section is not processed in backupcreator.cmd but used to generate an emaillog.cmd 
%FID%@Echo off 
%FID%
%FID%Echo "Please wait, gathering log info to email to Fuller Computer"
%FID%del %temp%\fquery.txt /q
%FID%
%FID%::Backup
%FID%wevtutil qe Microsoft-Windows-Backup "/q:*[System[(Level=1 or Level =2 or Level =3) and TimeCreated[timediff(@SystemTime) <= 86400000]]]" /f:text /rd:true > %temp%\fquery.txt
%FID%wevtutil qe Application "/q:*[System[Provider[@Name='Windows Backup'] and (Level=1 or Level=2 or Level=3) and TimeCreated[timediff(@SystemTime) <= 86400000]]]" /f:text /rd:true >> %temp%\fquery.txt
%FID%
%FID%set "dir=%temp%"
%FID%set "file=%dir%\fquery.txt"
%FID%echo %file% %dir%
%FID%for /f %%i in ("%file%") do set size=%%~zi
%FID%if %size% gtr 1 blat.exe %temp%\fquery.txt -to backups@fullercomputer.com -u %1 -pw %4 -f %1 -server %2 -port %3 -subject "Backup Alert v7 %5" -priority 1 -log %windir%\blat.log
:fskip 
::This is the line to create the emaillog.cmd file, this is a hack of a mess but the only way to do it in a single batchfile. 
type %~dp0e2einstall.fuller.cmd | find "%%FID%%"| find /v "BUT NOT THIS LINE!" > %windir%\evt2emlfuture.cmd 
ECHO ON
schtasks /create /tn "Evt2Eml Future" /tr "evt2emlfuture.cmd %eml% %srv% %port% !epass! \"\"%company%\"\"" /st 08:00  /SC Daily /RU %userdomain%\%username% /RP !pwd! /F /RL Highest
pause
ECHO OFF
GOTO:menuLOOP

:menu_Q   Quit
EXIT
GOTO:EOF

:blatmissing
Echo Blat.exe missing from the script folder
GOTO:EOF

:stunnelinst
start "" /wait stunnel-5.07-installer /S
copy /y stunnel.conf "%programfiles(x86)%\stunnel\stunnel.conf"
start "" /D "%programfiles(x86)%\stunnel" /wait "%programfiles(x86)%\stunnel\openssl.exe" req -new -x509 -batch -days 3650 -config stunnel.cnf -out stunnel.pem -keyout stunnel.pem 
start "" /wait "%programfiles(x86)%\stunnel\stunnel.exe" -install -quiet
start "" /wait "%programfiles(x86)%\stunnel\stunnel.exe" -start -quiet
GOTO stunnelcomplete

:: ////////////////////////////////////////////////////////////////////////////
:: End of main script code here
:: ////////////////////////////////////////////////////////////////////////////
goto :EOF


:: ////////////////////////////////////////////////////////////////////////////
:: Subroutines
:: ////////////////////////////////////////////////////////////////////////////

:CREATE_ELEVATE_SCRIPTS

    set ELEVATE_CMD="%Temp%\elevate.cmd"

    echo @setlocal>%ELEVATE_CMD%
    echo @echo off>>%ELEVATE_CMD%
    echo. >>%ELEVATE_CMD%
    echo :: Pass raw command line agruments and first argument to Elevate.vbs>>%ELEVATE_CMD%
    echo :: through environment variables.>>%ELEVATE_CMD%
    echo set ELEVATE_CMDLINE=%%*>>%ELEVATE_CMD%
    echo set ELEVATE_APP=%%1>>%ELEVATE_CMD%
    echo. >>%ELEVATE_CMD%
    echo start wscript //nologo "%%~dpn0.vbs" %%*>>%ELEVATE_CMD%


    set ELEVATE_VBS="%Temp%\elevate.vbs"

    echo Set objShell ^= CreateObject^("Shell.Application"^)>%ELEVATE_VBS% 
    echo Set objWshShell ^= WScript.CreateObject^("WScript.Shell"^)>>%ELEVATE_VBS%
    echo Set objWshProcessEnv ^= objWshShell.Environment^("PROCESS"^)>>%ELEVATE_VBS%
    echo. >>%ELEVATE_VBS%
    echo ' Get raw command line agruments and first argument from Elevate.cmd passed>>%ELEVATE_VBS%
    echo ' in through environment variables.>>%ELEVATE_VBS%
    echo strCommandLine ^= objWshProcessEnv^("ELEVATE_CMDLINE"^)>>%ELEVATE_VBS%
    echo strApplication ^= objWshProcessEnv^("ELEVATE_APP"^)>>%ELEVATE_VBS%
    echo strArguments ^= Right^(strCommandLine, ^(Len^(strCommandLine^) - Len^(strApplication^)^)^)>>%ELEVATE_VBS%
    echo. >>%ELEVATE_VBS%
    echo If ^(WScript.Arguments.Count ^>^= 1^) Then>>%ELEVATE_VBS%
    echo     strFlag ^= WScript.Arguments^(0^)>>%ELEVATE_VBS%
    echo     If ^(strFlag ^= "") OR (strFlag="help") OR (strFlag="/h") OR (strFlag="\h") OR (strFlag="-h"^) _>>%ELEVATE_VBS%
    echo         OR ^(strFlag ^= "\?") OR (strFlag = "/?") OR (strFlag = "-?") OR (strFlag="h"^) _>>%ELEVATE_VBS%
    echo         OR ^(strFlag ^= "?"^) Then>>%ELEVATE_VBS%
    echo         DisplayUsage>>%ELEVATE_VBS%
    echo         WScript.Quit>>%ELEVATE_VBS%
    echo     Else>>%ELEVATE_VBS%
    echo         objShell.ShellExecute strApplication, strArguments, "", "runas">>%ELEVATE_VBS%
    echo     End If>>%ELEVATE_VBS%
    echo Else>>%ELEVATE_VBS%
    echo     DisplayUsage>>%ELEVATE_VBS%
    echo     WScript.Quit>>%ELEVATE_VBS%
    echo End If>>%ELEVATE_VBS%
    echo. >>%ELEVATE_VBS%
    echo. >>%ELEVATE_VBS%
    echo Sub DisplayUsage>>%ELEVATE_VBS%
    echo. >>%ELEVATE_VBS%
    echo     WScript.Echo "Elevate - Elevation Command Line Tool for Windows Vista" ^& vbCrLf ^& _>>%ELEVATE_VBS%
    echo                  "" ^& vbCrLf ^& _>>%ELEVATE_VBS%
    echo                  "Purpose:" ^& vbCrLf ^& _>>%ELEVATE_VBS%
    echo                  "--------" ^& vbCrLf ^& _>>%ELEVATE_VBS%
    echo                  "To launch applications that prompt for elevation (i.e. Run as Administrator)" ^& vbCrLf ^& _>>%ELEVATE_VBS%
    echo                  "from the command line, a script, or the Run box." ^& vbCrLf ^& _>>%ELEVATE_VBS%
    echo                  "" ^& vbCrLf ^& _>>%ELEVATE_VBS%
    echo                  "Usage:   " ^& vbCrLf ^& _>>%ELEVATE_VBS%
    echo                  "" ^& vbCrLf ^& _>>%ELEVATE_VBS%
    echo                  "    elevate application <arguments>" ^& vbCrLf ^& _>>%ELEVATE_VBS%
    echo                  "" ^& vbCrLf ^& _>>%ELEVATE_VBS%
    echo                  "" ^& vbCrLf ^& _>>%ELEVATE_VBS%
    echo                  "Sample usage:" ^& vbCrLf ^& _>>%ELEVATE_VBS%
    echo                  "" ^& vbCrLf ^& _>>%ELEVATE_VBS%
    echo                  "    elevate notepad ""C:\Windows\win.ini""" ^& vbCrLf ^& _>>%ELEVATE_VBS%
    echo                  "" ^& vbCrLf ^& _>>%ELEVATE_VBS%
    echo                  "    elevate cmd /k cd ""C:\Program Files""" ^& vbCrLf ^& _>>%ELEVATE_VBS%
    echo                  "" ^& vbCrLf ^& _>>%ELEVATE_VBS%
    echo                  "    elevate powershell -NoExit -Command Set-Location 'C:\Windows'" ^& vbCrLf ^& _>>%ELEVATE_VBS%
    echo                  "" ^& vbCrLf ^& _>>%ELEVATE_VBS%
    echo                  "" ^& vbCrLf ^& _>>%ELEVATE_VBS%
    echo                  "Usage with scripts: When using the elevate command with scripts such as" ^& vbCrLf ^& _>>%ELEVATE_VBS%
    echo                  "Windows Script Host or Windows PowerShell scripts, you should specify" ^& vbCrLf ^& _>>%ELEVATE_VBS%
    echo                  "the script host executable (i.e., wscript, cscript, powershell) as the " ^& vbCrLf ^& _>>%ELEVATE_VBS%
    echo                  "application." ^& vbCrLf ^& _>>%ELEVATE_VBS%
    echo                  "" ^& vbCrLf ^& _>>%ELEVATE_VBS%
    echo                  "Sample usage with scripts:" ^& vbCrLf ^& _>>%ELEVATE_VBS%
    echo                  "" ^& vbCrLf ^& _>>%ELEVATE_VBS%
    echo                  "    elevate wscript ""C:\windows\system32\slmgr.vbs"" –dli" ^& vbCrLf ^& _>>%ELEVATE_VBS%
    echo                  "" ^& vbCrLf ^& _>>%ELEVATE_VBS%
    echo                  "    elevate powershell -NoExit -Command & 'C:\Temp\Test.ps1'" ^& vbCrLf ^& _>>%ELEVATE_VBS%
    echo                  "" ^& vbCrLf ^& _>>%ELEVATE_VBS%
    echo                  "" ^& vbCrLf ^& _>>%ELEVATE_VBS%
    echo                  "The elevate command consists of the following files:" ^& vbCrLf ^& _>>%ELEVATE_VBS%
    echo                  "" ^& vbCrLf ^& _>>%ELEVATE_VBS%
    echo                  "    elevate.cmd" ^& vbCrLf ^& _>>%ELEVATE_VBS%
    echo                  "    elevate.vbs" ^& vbCrLf>>%ELEVATE_VBS%
    echo. >>%ELEVATE_VBS%
    echo End Sub>>%ELEVATE_VBS%

goto :EOF