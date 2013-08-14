@echo off

Setlocal EnableDelayedExpansion

REM Determine whether network is present

set "NetDataPath=\\atle.bombardier.com\data\uk\pl\dos\Drgstate\"
set "ResultPath= c:\windows\temp\FileOpenResult.txt"

if not exist %NetDataPath% (
	set NetDataPath=c:\Drgstate\
	REM echo No network
) else (
	echo Network
)

REM Determine whether material or drawing
if "%2%"=="material" (
	set IndexPath=" %NetDataPath%PartsCurrentIndex.txt
	echo Material
) else (
	set IndexPath=" %NetDataPath%CurrentIndex.txt
	echo Drawing
)

REM Search for a match in the index
set "Redirect=>"
set "FindFiles=%IndexPath%%%Redirect%%ResultPath%"
echo Index:%IndexPath%%
echo Result:%ResultPath%
echo Find:!FindFiles!


rem find /i "%1% %%FindFiles%%

REM Stop at first line containing a path and open the file
for /F "tokens=*"  %%i IN ('findstr /v "INDEX" %ResultPath%') DO (
	rem"%%i"
	goto :Finish
)

:Finish


