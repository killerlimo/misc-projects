@echo off
REM DOS Batch File to open files created by BOM Tree in DrawingFinder
Set Version=2
echo.
echo.
echo.FileOpener Version %Version%

Setlocal EnableDelayedExpansion

REM Determine whether network is present

set NetDataPath=\\atle.bombardier.com\data\uk\pl\dos\Drgstate\
set LocalDataPath=c:\Drgstate\
set ResultPath=c:\windows\temp\FileOpenResult.txt

if not exist %NetDataPath% (
	set NetDataPath=%LocalDataPath%
	echo Local working
) else (
	echo Network working
)

REM Determine whether material or drawing
if "%2%"=="Material" (
	set IndexPath=%NetDataPath%PartsCurrentIndex.txt
	echo Material detected
) else (
	set IndexPath=%NetDataPath%CurrentIndex.txt
	echo Drawing detected
)
echo.
echo MENU
echo ----
echo.

REM Search for a match in the index
set SearchStr="%1%"
find /i %SearchStr% %IndexPath%>%ResultPath%

REM Stop at first line containing a path and open the file
set /a line=1
for /F "tokens=*"  %%i IN ('findstr /v "INDEX" %ResultPath%') DO (
	rem echo !line! - %%i
	set FileList[!line!]=%%i
	set /a line+=1
)

REM Check for file not found
if %line%==1 (
	echo File Not Found!
	pause
	goto :Finished
)

set /a i=1
:loop
	For %%A in ("!FileList[%i%]!") do (
		Set Folder=%%~dpA
		Set Name=%%~nxA
	)
	echo %i% - !Name!
	set /a i+=1
if not %i%==%line% goto :loop

echo %line% - Exit
echo.

set /p Selection=Type number and then press ENTER:
rem echo %Selection%
if %Selection%==!Line! (
	goto :Finish
)
set "Link=!FileList[%Selection%]!"

start "" "%Link%"

:Finish

