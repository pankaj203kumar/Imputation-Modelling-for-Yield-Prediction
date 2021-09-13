REM THIS SCRIPT CLIPS THE HEIGHT NORMALIZED LiDAR FILES BASED ON INDIVIDUAL TREE BOUNDARY FILES

@echo off

set Lastools_Directory=C:\LAStools\bin\
set Input_Files_Directory="Input File Path" REM INPUT LIDAR FILES IN .LAZ FORMAT
set Input_Polygon_File="Input Polygon Shapefile Name" REM INPUT BOUNDARY FILE OF INDIVIDUAL TREES IN .SHP FORMAT
set Output_File_Directory="Output File Path" REM OUTPUT LIDAR FILES IN .LAZ FORMAT

%Lastools_Directory%lasclip -i %Input_Files_Directory%\*.laz ^
-merged ^
-split Tree_ID ^ REM INDIVIDUAL TREE IDs
-poly %Input_Polygon_File% ^
-olaz ^
-odir %Output_File_Directory%\
pause