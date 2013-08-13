@echo off
echo Drawing: %1%
rem find /i "%1%" "c:\Drgstate\PartsCurrentIndex.txt" >  Result.txt
echo Results
for /F "tokens=*"  %%i IN ('findstr /v "INDEX" Result.txt') DO (
	%%i
	goto :Finish
)

:Finish

echo Finished

