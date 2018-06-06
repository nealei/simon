@echo off
:: 
:: Usage: Simon [Roadfile (without extension) [Trafficfile(without extension)]] [^> SimonSays.bat]
::
:: By Neale Irons Version 05/06/2018 (CC BY-SA 4.0)

if not exist Trarr.exe echo Error - Must run %0 from Trarr.exe program directory. Exiting. && goto :End
rem.>FAIL
for %%R in ("%1*.ROD") do (
  echo echo ::
  echo ::*** %%~pdnR
  echo copy "%%~pdnR.ROD" ROAD
  echo copy "%%~pdnR.MLT" MULTIP
  echo copy "%%~pdnR.OBS" OBS
  echo echo ::
  for %%T in ("%2*.TRF") do (
    echo echo ::
    echo ::** %%~pdnT
    echo copy "%%~pdnT.TRF" TRAF
    echo del PASS FAIL
    echo TRARR 2^>nul
    echo if exist PASS move /y OUT "%%~pdnR_%%~nT.OUT"
    echo if exist FAIL if not exist "%%~pdRBADOUTS"\. md "%%~pdRBADOUTS"
    echo if exist FAIL move /y OUT "%%~pdRBADOUTS\%%~nR_%%~nT.OUT"
   )
  )
  echo echo ::
  echo echo ::
)
:End