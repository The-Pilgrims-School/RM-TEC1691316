#
# Automated steps for "How to restart the RM Service Host on a CC4 network"

$ErrorActionPreference="Continue"
$VerbosePreference="Continue"

if (Get-Process | Where-Object { $_.Name -eq "RMMC"}) {
    Write-Error "Please close the RMMC and run the script again."
    exit
}

$ErrorActionPreference="Stop"

$TempDir = ($env:TEMP + "\TEC-1691316_" + (Get-Date -Format "yyyy-MM-dd_HH_mm_ss"))
New-Item -ItemType Directory -Path $TempDir


Write-Progress -Activity "Stopping services..." -Status "Stopping RM Service Host..." -PercentComplete 10

Stop-Service "RM Service Host" -Force

Write-Progress -Activity "Stopping services..." -Status "Stopping RM Unify Network Agent..." -PercentComplete 20

Stop-Service "RM Unify Network Agent Service" -Force # tends to lock some files

# HostedProxies
Write-Progress -Activity "Backing up and deleting files" -Status "HostedProxies..." -PercentComplete 35
New-Item -ItemType Directory -Path "$TempDir\HostedProxies"
Copy-Item -Recurse -Force -Path "C:\Program Files (x86)\RM\Connect\Comms\HostedProxies\*" -Destination "$TempDir\HostedProxies" -Verbose
Remove-Item -Recurse -Force -Path "C:\Program Files (x86)\RM\Connect\Comms\HostedProxies\*" -Verbose

# RemotingHost
Write-Progress -Activity "Backing up and deleting files" -Status "RemotingHost..." -PercentComplete 50
New-Item -ItemType Directory -Path "$TempDir\RemotingHost"
Copy-Item -Recurse -Force -Path "C:\Program Files (x86)\RM\Connect\Comms\RemotingHost\*" -Destination "$TempDir\RemotingHost" -Verbose
Remove-Item -Recurse -Force -Path "C:\Program Files (x86)\RM\Connect\Comms\RemotingHost\*" -Verbose

# RMCache/Cache
Write-Progress -Activity "Backing up and deleting files" -Status "RM Cache" -PercentComplete 65
Remove-Item "C:\ProgramData\RM\Connect\Comms\*.rmcache" -Force -Verbose
Remove-Item "C:\ProgramData\RM\Connect\Comms\*.cache" -Force -Verbose

# ProgramData ClientProxies
Write-Progress -Activity "Backing up and deleting files" -Status "ClientProxies" -PercentComplete 75
New-Item -ItemType Directory -Path "$TempDir\ClientProxies"
Copy-Item -Recurse -Force -Path "C:\ProgramData\RM\Connect\Comms\ClientProxies\*" -Destination "$TempDir\ClientProxies" -Verbose
Remove-Item -Recurse -Force -Path "C:\ProgramData\RM\Connect\Comms\ClientProxies\*" -Verbose

Write-Progress -Activity "Restarting services..." -Status "RM Service Host..." -PercentComplete 85
Start-Service "RM Service Host"

Write-Progress -Activity "Restarting services..." -Status "RM Unify Network Agent Service..." -PercentComplete 90
Start-Service "RM Unify Network Agent Service"

Write-Host "Complete. Backups are in $TempDir" -ForegroundColor White -BackgroundColor Green
