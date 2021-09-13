REM THIS SCRIPT GENERATES SPIKE-FREE CANOPY HEIGHT MODEL (CHM) FROM HEIGHT NORMALIZED LIDAR FILES

@echo off

set Lastools_Directory=C:\LAStools\bin\
set Input_Files_Directory_2="Input File Path" REM INPUT LIDAR FILES IN .LAZ FORMAT
set Output_Files_Directory_2="Output File Path" REM OUTPUT CHM FILES IN .IMG FORMAT

%Lastools_Directory%las2dem -i %Input_Files_Directory_2%\*.laz ^
-step 0.25 ^ REM CHM WITH 25cm RESOLUTION
-elevation ^ REM CHM BASED ON ELEVATION ATTRIBUTE
-spike_free 0.81 ^ REM GENERATES SPIKE-FREE CHM
-oimg ^
-odir %Output_Files_Directory_2%\
pause
