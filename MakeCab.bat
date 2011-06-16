@echo off

set RPath=Redist
set OutFile=ClearCab.bat

title User Redist package [make]

echo [Start] at %time% %date%
echo.

Misc\CompileVB6 check::vb6check.bat
call vb6check.bat
if exist vb6check.bat del vb6check.bat

if not "%VBInstalled%" == "1" (
  echo Visual Basic 6 not detected. Exiting...
  pause>nul
  exit
)

call %OutFile% skip

echo Compressing source...

Misc\MsCabFC .\::script.txt::Src.cab>nul
Misc\MakeCab /F script.txt>nul

if exist script.001 del script.001
if exist script.002 del script.002
if exist script.txt del script.txt

echo Making directory structure...

if not exist %RPath% MkDir %RPath%
if not exist %RPath%\BCompiler MkDir %RPath%\BCompiler

echo Compiling programs...
echo - CBlaster
Misc\CompileVB6 make::CharBlaster\CBlaster.vbp
echo - VBlaster
Misc\CompileVB6 make::VaultBlaster\VBlaster.vbp
echo - BCompiler
Misc\CompileVB6 make::BCompiler\BCompiler.vbp

echo Compiling library...
BCompiler\Bcompiler data.src::mudata.lib::silence

echo Copying files...
copy VaultBlaster\VBlaster.exe %RPath%>nul
copy VaultBlaster\1046.vbl %RPath%>nul
copy CharBlaster\CBlaster.exe %RPath%>nul
copy CharBlaster\1046.cbl %RPath%>nul

copy BCompiler\mudata.lib %RPath%>nul
if exist BCompiler\mudata.lib del BCompiler\mudata.lib

copy BCompiler\BCompiler.exe %RPath%\BCompiler>nul
copy BCompiler\Data.src %RPath%\BCompiler>nul
copy Src.cab %RPath%>nul
if exist Src.cab del Src.cab

copy Misc\Redist.ReadMe.htm %RPath%\ReadMe.htm>nul
Misc\KillMetaName %RPath%\ReadMe.htm

copy Misc\Redist.LeiaMe.htm %RPath%\LeiaMe.htm>nul
Misc\KillMetaName %RPath%\LeiaMe.htm

copy Misc\BCompiler.Datasrc.htm %RPath%\BCompiler\Datasrc.htm>nul
Misc\KillMetaName %RPath%\BCompiler\Datasrc.htm

copy Misc\BCompiler.DefCompile.bat %RPath%\BCompiler\DefCompile.bat>nul

echo Compressing redist...

if not exist Temp MkDir Temp

Misc\MsCabFC %RPath%::Temp\script.txt::Temp\MuBlasters.cab>nul
Misc\MakeCab /F Temp\script.txt>nul
if exist Temp\script.001 del Temp\script.001
if exist Temp\script.002 del Temp\script.002
if exist Temp\script.txt del Temp\script.txt

copy Temp\MuBlasters.cab .\>nul

del /s /q Temp\*.*>nul
RmDir Temp

call %OutFile%  skip nocab

echo.
echo [End] at %time% %date%
pause>nul