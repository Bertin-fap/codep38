:: Bertin F. 26-05-2024

::curl https://github.com/sindresorhus/recycle-bin/releases/download/v2.0.0/recycle-bin.exe
@echo off
Title create COPEP.exe
::color 17

mkdir %TEMP%\CODEP_exe
@echo ceate %TEMP%\CODEP_exe successfully

:: create a venv with python 3.9.7
:: adapted from https://stackoverflow.com/questions/45833736/how-to-store-python-version-in-a-variable-inside-bat-file-in-an-easy-way?noredirect=1
for /F "tokens=* USEBACKQ" %%F in (`python --version`) do (set var=%%F)
echo create a virtual environment with %var%
cd %TEMP%\CODEP_exe
python -m venv venv

:: activate the venv
set VIRTUAL_ENV=%TEMP%\CODEP_exe\venv
call %VIRTUAL_ENV%\Scripts\activate.bat

:: install packages
::pip install git+https://github.com/Bertin-fap/ctg.git#egg=ctg
pip install --user --upgrade pip setuptools wheel
pip install %userprofile%\pyvenv\codep38
::pip install ctg
pip install greenlet==3.0.0rc3
pip install pyinstaller

:: set the default directories
:: ICON contains the icon file with the format.ico
:: PGM contain the application lauch python program

set "ICON=%TEMP%/CODEP_exe/venv/Lib/site-packages/codep38/codep38functs/CTG_RefFiles/logo_ffvelo.ico"
set "DATA=%TEMP%/CODEP_exe/venv/Lib/site-packages/codep38;codep38/"
set "PGM=%TEMP%/CODEP_exe/venv/Lib/site-packages/codep38/codep38functs/CTG_RefFiles/CODEP_METER.py"

:: make the executable 
pyinstaller --noconfirm --onefile --console^
 --icon "%ICON%"^
 --add-data "%DATA%"^
 --collect-data "docxcompose"^
 "%PGM%"
 
:: remove the directories build
set "BUILD=%TEMP%\CTG_exe\build"
rmdir /s /q %BUILD%

for /f "tokens=1-3 delims=/ " %%a in ('date /T') do (set mydate=%%c-%%b-%%a)
set dirname="%mydate% CODEP_Meter"
rename dist %dirname%
 
set "new_file_name=%dirname%.exe"
ren %TEMP%\CODEP_exe\%dirname%\CODEP_meter.exe %new_file_name%
echo %new_file_name% is located in %TEMP%\CODEP_exe\%dirname% 

:: Copy Exe
set input_file=%TEMP%\CODEP_exe\%dirname%\%dirname%.exe%
set /p "rep=do you want to copy this file in a folder (y/n) : "
if NOT %rep%==y GOTO FIN
set /p "rep1=do you use %userprofile%\CTG (y/n) : "
if NOT %rep%==y GOTO A
copy  %input_file% %userprofile%\CTG
GOTO FIN
A: set /p "new_dir=enter the full path of the folder : "
set output_file=%new_dir%\%dirname%.exe%
copy  %input_file% %output_file%
:FIN

pause