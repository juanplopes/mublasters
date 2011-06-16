@echo off
 
set RPath=Redist
 
if not "%1" == "skip" (
  title User Redist package [clear]
)
 
if not "%2" == "nocab" (
   echo Removing old redist files...
) ELSE (
   echo Removing temp redist files...
)
 
if not "%3" == "keep" (
  if exist %RPath% rmdir /s /q Redist>nul
) 
if exist VaultBlaster\VBlaster.exe del VaultBlaster\VBlaster.exe 
if exist VaultBlaster\blaster.vws del VaultBlaster\blaster.vws
 
if exist CharBlaster\CBlaster.exe del CharBlaster\CBlaster.exe 
if exist CharBlaster\blaster.cws del CharBlaster\blaster.cws
 
if exist BCompiler\BCompiler.exe del BCompiler\BCompiler.exe 
if exist BCompiler\mudata.lib del BCompiler\mudata.lib
 
if not "%2" == "nocab" if exist MuBlasters.cab del MuBlasters.cab
 
if not "%1" == "skip" (
  echo.
  echo Done.
  pause>nul
)
