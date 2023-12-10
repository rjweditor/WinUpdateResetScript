# Windows Update Reset Script

# Stop Services Function
function Stop-ServiceAndWait {
    param(
        [string]$serviceName
    )

    $attempts = 0
    do {
        $attempts++
        Stop-Service -Name $serviceName
        $serviceStatus = Get-Service -Name $serviceName | Select-Object -ExpandProperty Status
        if ($serviceStatus -eq "Stopped") {
            Write-Host "Checking the $serviceName service status..."
            return $true
        }
    } while ($attempts -lt 3)

    Write-Host "Failed to reset Windows Update due to $serviceName service failing to stop."
    return $false
}

# Reset Windows Update Components Function
function Reset-WindowsUpdate {
    ipconfig /flushdns

    Get-ChildItem "$env:ALLUSERSPROFILE\Application Data\Microsoft\Network\Downloader\qmgr*.dat" -Recurse | Remove-Item -Force
    Get-ChildItem "$env:ALLUSERSPROFILE\Microsoft\Network\Downloader\qmgr*.dat" -Recurse | Remove-Item -Force

    $filesToRename = @("pending.xml", "SoftwareDistribution", "Catroot2", "WindowsUpdate.log")
    foreach ($file in $filesToRename) {
        $fileBak = "$env:SYSTEMROOT\winsxs\$file.bak"
        if (Test-Path $fileBak) { Remove-Item $fileBak -Force }
        if (Test-Path "$env:SYSTEMROOT\winsxs\$file") {
            Take-Own -Path "$env:SYSTEMROOT\winsxs\$file"
            Get-Item -Path "$env:SYSTEMROOT\winsxs\$file" | ForEach-Object { $_.Attributes = 'Archive' }
            Rename-Item -Path "$env:SYSTEMROOT\winsxs\$file" -NewName "$file.bak"
        }
    }

    $dirsToRename = @("SoftwareDistribution", "Catroot2")
    foreach ($dir in $dirsToRename) {
        $dirBak = "$env:SYSTEMROOT\$dir.bak"
        if (Test-Path $dirBak) { Remove-Item $dirBak -Recurse -Force }
        if (Test-Path "$env:SYSTEMROOT\$dir") {
            Get-Item -Path "$env:SYSTEMROOT\$dir" | ForEach-Object { $_.Attributes = 'Archive' }
            Rename-Item -Path "$env:SYSTEMROOT\$dir" -NewName "$dir.bak"
        }
    }

    $logsToDelete = @("$env:SYSTEMROOT\WindowsUpdate.log", "$env:SYSTEMROOT\WindowsUpdate.log.bak")
    foreach ($log in $logsToDelete) {
        if (Test-Path $log) { Remove-Item $log -Force }
    }

    $acl = "D:(A;;CCLCSWRPWPDTLOCRRC;;;SY)(A;;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;BA)(A;;CCLCSWLOCRRC;;;AU)(A;;CCLCSWRPWPDTLOCRRC;;;PU)"
    sc.exe sdset bits $acl
    sc.exe sdset wuauserv $acl

    $dlls = @(
        "atl", "urlmon", "mshtml", "shdocvw", "browseui", "jscript", "vbscript",
        "scrrun", "msxml", "msxml3", "msxml6", "actxprxy", "softpub", "wintrust",
        "dssenh", "rsaenh", "gpkcsp", "sccbase", "slbcsp", "cryptdlg", "oleaut32",
        "ole32", "shell32", "initpki", "wuapi", "wuaueng", "wuaueng1", "wucltui",
        "wups", "wups2", "wuweb", "qmgr", "qmgrprxy", "wucltux", "muweb", "wuwebv",
        "wudriver"
    )
    foreach ($dll in $dlls) {
        regsvr32.exe /s "$dll.dll"
    }

    netsh winsock reset
    netsh winsock reset proxy
}

# Start Services Function
function Start-ServiceIfStopped {
    param(
        [string]$serviceName
    )

    Start-Service -Name $serviceName -ErrorAction SilentlyContinue
}

# Script Execution
Write-Host "Windows Update Reset Script"

# Stop Services
$servicesToStop = @("bits", "wuauserv", "appidsvc", "cryptsvc")
foreach ($service in $servicesToStop) {
    if (-not (Stop-ServiceAndWait -serviceName $service)) {
        exit 1
    }
}

# Reset Windows Update Components
Reset-WindowsUpdate

# Start Services
$servicesToStart = @("bits", "wuauserv", "appidsvc", "cryptsvc")
foreach ($service in $servicesToStart) {
    Start-ServiceIfStopped -serviceName $service
}

Write-Host "Task completed successfully! Please restart your computer and check for updates again."
