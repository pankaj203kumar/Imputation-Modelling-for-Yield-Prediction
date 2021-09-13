REM THIS SCRIPT REPROJECTS THE COORDINATE SYSTEM OF LIDAR FILES

@echo off

set Lastools_Directory=C:\LAStools\bin\
set Input_Files_Directory="Input File Path" REM INPUT LIDAR FILES IN .LAZ FORMAT
set Output_Files_Directory="Output File Path" REM OUTPUT LIDAR FILES IN .LAZ FORMAT

%Lastools_Directory%las2las -i %Input_Files_Directory%\*.laz ^
-epsg 32754 ^ REM ESPG CODE OF INPUT LIDAR FILES
-target_epsg 28354 ^ REM EPSG CODE OF OUTPUT LIDAR FILES
-target_precision 0.01 ^
-olaz ^
-odir %Output_Files_Directory%\
pause