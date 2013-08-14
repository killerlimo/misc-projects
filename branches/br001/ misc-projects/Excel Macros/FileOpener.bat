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
set "Redirect=>"
set FindFiles="%IndexPath% %Redirect% %ResultPath%"
echo Index:%IndexPath%%
echo Result:%ResultPath%
echo Find:%FindFiles%

find /i %1% %FindFiles%


