@echo off
REM Determine whether network is present

set NetDataPath=\\atle.bombardier.com\data\uk\pl\dos\Drgstate\
set ResultPath= ^> c:\windows\temp\FileOpenResult.txt

if not exist "%NetDataPath%\nul" set NetDataPath=c:\Drgstate\

	
REM Determine whether material or drawing
if "%2%"=="material" (
	set IndexPath=" %NetDataPath%PartsCurrentIndex.txt
) else (
	set IndexPath=" %NetDataPath%CurrentIndex.txt
)

REM Search for a match in the index
set FindFiles=%IndexPath%% ^> %ResultPath%
echo %IndexPath%%
echo %ResultPath%
echo %FindFiles%

rem find /i "%1% %%FindFiles%%

REM Stop at first line containing a path and open the file
for /F "tokens=*"  %%i IN ('findstr /v "INDEX" %ResultPath%') DO (
	rem"%%i"
	goto :Finish
)

:Finish


