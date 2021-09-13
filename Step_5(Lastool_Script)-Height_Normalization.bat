REM THIS SCRIPT NORMALIZES HEIGHT IN GROUND ESTIMATED LIDAR FILES

@echo off

set Lastools_Directory=C:\LAStools\bin\
set Input_Files_Directory_2="Input File Path" REM INPUT LIDAR FILES IN .LAZ FORMAT
set Output_Files_Directory_2="Output File Path" REM OUTPUT LIDAR FILES IN .LAZ FORMAT

%Lastools_Directory%lasheight -i %Input_Files_Directory_2%\*.laz ^
-replace_z ^
-drop_below 0 ^ REM LOWER HEIGHT THRESHOLD IN METRES
-drop_above 50 ^ REM HIGHER HEIGHT THRESHOLD IN METRES
-olaz ^
-odir %Output_Files_Directory_2%\
pause