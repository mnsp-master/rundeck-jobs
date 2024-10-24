$mnspver = "0.0.8"


Write-Host $(Get-Date)
Write-Host "MNSP Version" $mnspver
Start-Sleep 10
$ErrorActionPreference="Continue"
Set-Location $GamDir

Write-host "--------------------------------------`n"

#Write-Host "Current workspace source..."
#Invoke-Expression "$GamDir\gam.exe"

Write-host "--------------------------------------`n"

Write-Host "Setting workspace source: $GoogleWorkSpaceSource"
Invoke-Expression "$GamDir\gam.exe select $GoogleWorkSpaceSource save" # swap/set google workspace
Invoke-Expression "$GamDir\gam.exe"

Start-sleep 5
Write-host "--------------------------------------`n"

Write-Host "Setting workspace destination: $GoogleWorkSpaceDestination"
Invoke-Expression "$GamDir\gam.exe select $GoogleWorkSpaceDestination save" # swap/set google workspace
Invoke-Expression "$GamDir\gam.exe"

Write-host "--------------------------------------`n"

Start-sleep 5
Write-Host $(Get-Date)
