@echo off
echo.
echo Simple script to reset Windows Update components.
echo.
pause

setlocal

:stopService
for %%s in (bits wuauserv appidsvc cryptsvc) do (
    set attempts=0
    :tryStop
    set /a attempts+=1
    net stop %%s
    sc query %%s | findstr /I /C:"STOPPED" >nul && (
        echo Checking the %%s service status...
        goto :serviceStopped
    )
    if %attempts% lss 3 goto :tryStop
    echo Failed to reset Windows Update due to %%s service failing to stop.
    pause
    goto :Start
    :serviceStopped
)

:resetUpdate
Ipconfig /flushdns
del /s /q /f "%ALLUSERSPROFILE%\Application Data\Microsoft\Network\Downloader\qmgr*.dat" 
del /s /q /f "%ALLUSERSPROFILE%\Microsoft\Network\Downloader\qmgr*.dat"
cd /d %windir%\system32

for %%f in (pending.xml SoftwareDistribution Catroot2 WindowsUpdate.log) do (
    if exist "%SYSTEMROOT%\winsxs\%%f.bak" del /s /q /f "%SYSTEMROOT%\winsxs\%%f.bak" 
    if exist "%SYSTEMROOT%\winsxs\%%f" (
        takeown /f "%SYSTEMROOT%\winsxs\%%f" 
        attrib -r -s -h /s /d "%SYSTEMROOT%\winsxs\%%f" 
        ren "%SYSTEMROOT%\winsxs\%%f" %%f.bak 
    )
)

for %%d in (SoftwareDistribution Catroot2) do (
    if exist "%SYSTEMROOT%\%%d.bak" rmdir /s /q "%SYSTEMROOT%\%%d.bak"
    if exist "%SYSTEMROOT%\%%d" (
        attrib -r -s -h /s /d "%SYSTEMROOT%\%%d" 
        ren "%SYSTEMROOT%\%%d" %%d.bak 
    )
)

if exist "%SYSTEMROOT%\WindowsUpdate.log.bak" del /s /q /f "%SYSTEMROOT%\WindowsUpdate.log.bak" 
if exist "%SYSTEMROOT%\WindowsUpdate.log" (
    attrib -r -s -h /s /d "%SYSTEMROOT%\WindowsUpdate.log" 
    ren "%SYSTEMROOT%\WindowsUpdate.log" WindowsUpdate.log.bak 
)

sc.exe sdset bits D:(A;;CCLCSWRPWPDTLOCRRC;;;SY)(A;;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;BA)(A;;CCLCSWLOCRRC;;;AU)(A;;CCLCSWRPWPDTLOCRRC;;;PU)
sc.exe sdset wuauserv D:(A;;CCLCSWRPWPDTLOCRRC;;;SY)(A;;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;BA)(A;;CCLCSWLOCRRC;;;AU)(A;;CCLCSWRPWPDTLOCRRC;;;PU)

for %%dll in (
    atl urlmon mshtml shdocvw browseui jscript vbscript scrrun msxml msxml3 msxml6 actxprxy softpub wintrust dssenh rsaenh gpkcsp sccbase slbcsp cryptdlg oleaut32 ole32 shell32 initpki wuapi wuaueng wuaueng1 wucltui wups wups2 wuweb qmgr qmgrprxy wucltux muweb wuwebv wudriver
) do regsvr32.exe /s "%%dll.dll"

netsh winsock reset
netsh winsock reset proxy

:startServices
for %%s in (bits wuauserv appidsvc cryptsvc) do (
    net start %%s
)

echo Task completed successfully! Please restart your computer and check for updates again.
:end
endlocal
