$mnspver = "0.0.12"


Write-Host $(Get-Date)
Write-Host "MNSP Version" $mnspver
Start-Sleep 10
$ErrorActionPreference="Continue"
Set-Location $GamDir

Write-host "--------------------------------------`n"

Write-Host "Setting workspace source: $GoogleWorkSpaceSource"
Invoke-Expression "$GamDir\gam.exe select $GoogleWorkSpaceSource save" # swap/set google workspace
Invoke-Expression "$GamDir\gam.exe info domain"

Write-Host "Getting members of users to process source group $GoogleWorkspaceSourceGroup"
Invoke-Expression "$GamDir\gam.exe print group-members group_ns $GoogleWorkspaceSourceGroup > $tempcsv"

#<
Start-sleep 5
Write-host "--------------------------------------`n"

#Write-Host "Setting workspace destination: $GoogleWorkSpaceDestination"
#Invoke-Expression "$GamDir\gam.exe select $GoogleWorkSpaceDestination save" # swap/set google workspace
#Invoke-Expression "$GamDir\gam.exe"

Write-host "--------------------------------------`n"

#Start-sleep 5
Write-Host $(Get-Date)
#>

