REM THIS SCRIPT ESTIMATES THE BOUNDARY OF LIDAR FILES

@echo off

set Lastools_Directory=C:\LAStools\bin\
set Input_Files_Directory="Input File Path" REM INPUT LIDAR FILES IN .LAZ FORMAT
set Output_Files_Directory="Output File Path" REM OUTPUT FILE IN .SHP FORMAT

%Lastools_Directory%lasboundary -i %Input_Files_Directory%\*.laz ^
-oshp ^
-odir %Output_Files_Directory%\
pause