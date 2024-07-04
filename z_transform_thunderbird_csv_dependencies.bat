@echo on
@setlocal ENABLEDELAYEDEXPANSION
@setlocal enableextensions

REM https://github.com/vapoursynth/vapoursynth/releases/download/R67/Install-Portable-VapourSynth-R67.ps1
REM https://github.com/vapoursynth/vapoursynth/releases/download/R67/VapourSynth64-Portable-R67.zip
REM per https://github.com/vapoursynth/vapoursynth/issues/1037#issuecomment-1996980426
REM powershell -executionpolicy bypass -File .\Install-Portable-VapourSynth-RXX.ps1 -TargetFolder <wherever you want> -Unattended

set "wget=c:\software\wget\wget.exe"
set "wunzip=C:\Program Files\WinZip\WZUNZIP.EXE"
set "z7=c:\software\7zip\7za.exe"

REM set the root for the newe VapourSynth to be installed in
REM set "root=G:\TEST\Vapoursynth-x64\"
set "root=C:\SOFTWARE\Vapoursynth-x64\"

REM Set the VapourSynth version to download and install
set "VSVersion=68"
set "VSVersionR=R!VSVersion!"
set "VSFile=VapourSynth64-Portable-!VSVersionR!.zip"
set "VSps1=Install-Portable-VapourSynth-!VSVersionR!.ps1"

REM Set the MediaInfo version to download and install
REM set "MIv=24.04"
set "MIv=24.05"

REM Set the dgdecnv version to download and install
Set "DGv=255"

REM Set the 7zip version to download and install
set "version_7zip=2407"

REM ------------------------------------------------
set "vs_path=!root!"
REM ensure trailing slash exists
if /I NOT "!vs_path:~-1!" == "\" (set "vs_path=!vs_path!\")
set "vspath_nobackslash=%vs_path:~0,-1%

if not exist "!vs_path!" mkdir "!vs_path!"
set "vs_path_drive=!vs_path:~,2!"
set "vs_scripts_path=!vs_path!vs-scripts"
set "vs_plugins_path=!vs_path!vs-plugins"
set "vs_coreplugins_path=!vs_path!vs-coreplugins"
set "py_exe=!vs_path!python.exe"

set "vs_temp=!vs_path!TEMP\"
if /I NOT "!vs_temp:~-1!" == "\" (set "vs_temp=!vs_temp!\")
if not exist "!vs_temp!" (mkdir "!vs_temp!")

REM ------------------------------------------------

REM CD into the target path and do the work
!vs_path_drive!
cd "!vs_path!"
cd

REM "!py_exe!" -m pip install pip --target=%vs_path% --no-cache-dir --upgrade --check-build-dependencies --upgrade-strategy eager --verbose

"!py_exe!" -m pip uninstall pywin32 --no-cache-dir --verbose

pause
goto :eof
