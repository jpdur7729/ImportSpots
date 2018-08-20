REM Original script created by Chadi Hammed with assistance from Bachir Allouch
REM Updated by C Bowser 26/09/2017

@echo off

REM Set the Working Directory to a suitable place on the server
set workingdir=C:\Users\CBOWS\Desktop\ECB
REM Set the location of FrontCmd
set efrontdir=C:\batch
REM Set the website URL including https
set website=https://uktraining.frontsrv.com
REM Set the username & password to access the website, remember to use an ENC password!
set username=CBOWS
set password=enc:hLDsNpiXnkKkVgcFw6xvIvO1yD3s4ZAZp+nsy0mOPKI=

REM This needs the GnuWin32 utilty WGET to work, download from sourceforge 
cd C:\Program Files (x86)\GnuWin32\bin
wget http://www.ecb.europa.eu/stats/eurofxref/eurofxref-daily.xml -P %workingdir%

cd %workingdir%
 
set fichierxml="%workingdir%\eurofxref-daily.xml"
 
if exist eurofxref-daily-import.csv (
del eurofxref-daily-import.csv
)
 
echo Reference date;Source Currency;Destination Currency;Rate;Description > eurofxref-daily-import.csv
 
for /f "skip=7 tokens=2 delims='" %%i in ('type "%fichierxml%"') do set refdate=%%i & goto suit
:suit
REM echo %refdate%
 
set newrefdate=%refdate:~8,2%/%refdate:~5,2%/%refdate:~0,4%
REM echo %newrefdate%
set refccy=EUR
REM echo %refccy%
 
set compt=
setlocal enableDelayedExpansion
for /f " skip=8 tokens=2,4 delims='" %%j in ('type "%fichierxml%"') do  (
set /A compt +=1
If !compt! LEQ 33 echo %newrefdate%;%refccy%;%%j;%%k;; >> eurofxref-daily-import.csv
)
endlocal


cd %efrontdir%
FrontCmd ExecImport /server:%website% /userid:%username% /password:%password% /params:"Standard Currency Rates" /files:%workingdir%\eurofxref-daily-import.csv

del %workingdir%\eurofxref-daily.xml
