REM THIS SCRIPT REMOVES THE NOISE (OUTLIER) FROM THE CLIPPED LIDAR FILES

@echo off

set Lastools_Directory=C:\LAStools\bin\
set Input_Files_Directory_1="Input File Path 1" REM INPUT LIDAR FILES IN .LAZ FORMAT
set Output_Files_Directory_1="Output File Path 1" REM OUTPUT LIDAR FILES IN .LAZ FORMAT

%Lastools_Directory%lasnoise -i %Input_Files_Directory_1%\*.laz ^
-step_xy 4 ^ REM GRID SIZE IN METRES ALONG X-Y AXIS
-step_z 1 ^ REM GRID SIZE IN METRES ALONG Z-AXIS
-remove_noise ^
-olaz ^
-odir %Output_Files_Directory_1%\
pause

set Input_Files_Directory_2="Output File Path 1" REM INPUT LIDAR FILES IN .LAZ FORMAT AS OUTPUT FROM ABOVE PROCESSING
set Output_Files_Directory_2="Output File Path 2" REM OUTPUT LIDAR FILES IN .LAZ FORMAT

%Lastools_Directory%las2las -i %Input_Files_Directory_2%\*.laz ^
-drop_z_below 0 ^ REM LOWER HEIGHT THRESHOLD IN METRES
-drop_z_above 50 ^ REM HIGHER HEIGHT THRESHOLD IN METRES
-olaz ^
-odir %Output_Files_Directory_2%\
pause