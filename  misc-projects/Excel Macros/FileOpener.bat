REM DOS Batch File to open files created by BOM Tree in DrawingFinder
REM Version 1

@echo off

Setlocal EnableDelayedExpansion

REM Determine whether network is present

set NetDataPath=\\atle.bombardier.com\data\uk\pl\dos\Drgstate\
set ResultPath=c:\windows\temp\FileOpenResult.txt

if not exist %NetDataPath% (
	set NetDataPath=c:\Drgstate\
	echo No network
) else (
	echo Network
)

REM Determine whether material or drawing
if "%2%"=="Material" (
	set IndexPath=%NetDataPath%PartsCurrentIndex.txt
	echo Material
) else (
	set IndexPath=%NetDataPath%CurrentIndex.txt
	echo Drawing
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
