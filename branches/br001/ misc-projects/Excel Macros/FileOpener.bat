@echo off
echo Number: %1%
echo Type: %2%

if "%2%" == "material" (
	echo It is a Material
	rem find /i "%1%" "c:\Drgstate\PartsCurrentIndex.txt" >  Result.txt
) else (
	echo It is a Drawing
)

echo Results
for /F "tokens=*"  %%i IN ('findstr /v "INDEX" Result.txt') DO (
	%%i
	goto :Finish
)

:Finish

echo Finished

