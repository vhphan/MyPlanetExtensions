rem Make sure project does not reference development environment
echo %1
find ".wnp\" %1
if ERRORLEVEL 1 goto GoodReference
echo ****** ERROR: Bad Reference detected coming from Marconi.wnp!
exit 1

:GoodReference
echo References OK!

rem set the solution folder
set solution=%2
set solution=%solution:"=%
echo Solution %solution%

rem throw the first two parameters away
shift
shift
set params=%1
:loop
shift
if [%1]==[] goto afterloop
set params=%params% %1
goto loop
:afterloop

if exist "%solution%..\..\References\common\CommonUtilities.dll" set assemblies=%solution%..\..\References\common\

rem Primary assembly source is the Planet installation
if not exist "%assemblies%CommonUtilities.dll" set assemblies=C:\Program Files\InfoVista\Planet 7.6\Assembly\



echo Assemblies %assemblies%

rem Now copy the assemblies needed for the build
for %%x in (%params%) do if exist "%assemblies%%%x" xcopy "%assemblies%%%x" "%solution%References\" /y /D

exit 0
