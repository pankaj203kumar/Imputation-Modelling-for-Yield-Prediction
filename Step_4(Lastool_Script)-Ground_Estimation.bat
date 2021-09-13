REM THIS SCRIPT ESTIMATES GROUND POINTS IN NOISE-FREE LIDAR FILES

@echo off

set Lastools_Directory=C:\LAStools\bin\
set Input_Files_Directory_1="Input File Path" REM INPUT LIDAR FILES IN .LAZ FORMAT
set Output_Files_Directory_1="Output File Path" REM OUTPUT LIDAR FILES IN .LAZ FORMAT

%Lastools_Directory%lasground -i %Input_Files_Directory_1%\*.laz ^
-wilderness ^ REM STEP SIZE OF 3m IS USED TO ESTIMATE THE GROUND POINTS
-olaz ^
-odir %Output_Files_Directory_1%\
pause