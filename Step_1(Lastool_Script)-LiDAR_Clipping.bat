REM THIS SCRIPT CLIPS THE LIDAR FILES BASED ON PLOT BOUNDARY FILES

@echo off

set Lastools_Directory=C:\LAStools\bin\
set Input_Files_Directory="Input File Path" REM INPUT LIDAR FILES IN .LAZ FORMAT
set Input_Polygon_File="Input Polygon Shapefile Name" REM INPUT INVENTORY BOUNDARY FILE IN .SHP FORMAT
set Output_File_Directory="Output File Path" REM OUTPUT LIDAR FILES IN .LAZ FORMAT

%Lastools_Directory%lasclip -i %Input_Files_Directory%\*.laz ^
-merged ^
-split plot_ID ^ REM PLOT IDS OF THE PLANTATION PLOTS
-poly %Input_Polygon_File% ^
-olaz ^
-odir %Output_File_Directory%\
pause