@echo off
REM DOS Batch File to open files created by BOM Tree in DrawingFinder
Set Version=1

echo FileOpener Version %Version%

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

REM Search for a match in the index
set SearchStr="%1%"
find /i %SearchStr% %IndexPath%>%ResultPath%

REM Stop at first line containing a path and open the file
for /F "tokens=*"  %%i IN ('findstr /v "INDEX" %ResultPath%') DO (
	start "" "%%i"
	goto :Finish
)

:Finish
