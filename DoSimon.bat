@echo off
:: 
:: Usage: GoSimon [Roadfile (without extension) [Trafficfile(without extension) [RoadSec]]] [^> SimonSays.bat]
::
:: By Neale Irons Version 03/06/2018 (CC BY-SA 4.0)

for %%a in ("%cd%") do set "CurDir=%%~nxa"
if exist "Trarr trf Values.xlsx" python ..\make_trf.py "Trarr trf Values%3.xlsx"
if exist ..\pypath.bat call ..\pypath
call MakeSimonSetup > MakeSimonSetup.log
cd ..
echo .
echo Please wait - simulating all traffic on all road options ...
call simon "%CurDir%\%1*" "%CurDir%\%2*" > %CurDir%\SimonSays%3.bat
call %CurDir%\SimonSays%3 > %CurDir%\SimonSays%3.log
if exist cleanup.bat call cleanup
cd %CurDir%
python ..\avg_seeds.py "%1*%2*" > avg_seeds%3.log
echo  Done.