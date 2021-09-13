REM THIS SCRIPT EXTRACT INFORMATION FROM CLIPPED LIDAR FILES (POINT DENSITY, HISTOGRAM)

@echo off

set Lastools_Directory=C:\LAStools\bin\
set Input_Files_Directory="Input File Path" REM INPUT LIDAR FILES IN .LAZ FORMAT
set Output_File=Output File Name REM OUTPUT FILE NAME IN .TEXT FORMAT
set Output_File_Directory="Output File Path"

%Lastools_Directory%lasinfo -i %Input_Files_Directory%\*.laz ^
-merged ^
-cd ^ REM COMPUTE DENSITY
-histo z 1 ^ REM ELEVATION ATTRIBUTE
-histo intensity 1 ^ REM INTENSITY ATTRIBUTE
-otxt ^
-odir %Output_File_Directory%\ ^
-o %Output_File%
pause