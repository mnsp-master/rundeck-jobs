$mnspver = "0.0.5"

function DashedLine {
Write-host "-----------------------------------------------------------`n"
}

Write-Host $(Get-Date)
Write-Host "MNSP Version" $mnspver
#Start-Sleep 10
$ErrorActionPreference="Continue"
Set-Location $GamDir

#Set Google Instance: Destination...
Write-Host "###### Set Google instance: Destination... ######"

Write-Host "Setting workspace Destination: $GoogleWorkSpaceDestination"
Invoke-Expression "$GamDir\gam.exe select $GoogleWorkSpaceDestination save" # swap/set google workspace
Invoke-Expression "$GamDir\gam.exe info domain"
start-sleep 3
DashedLine

