$mnspver = "0.0.7"

Write-Host $(Get-Date)
Write-Host "MNSP Version" $mnspver

$ScriptName = Split-Path $PSCommandPath -Leaf
Write-Host "Executing Main Rundeck Job..."
Write-Host "MNSP script: $scriptName version: $mnspver"

#Start-Sleep 10
$ErrorActionPreference="Continue"
Set-Location $GamDir

#Set Google Instance: Destination...
Write-Host "###### Set Google instance: Destination... ######"

Write-Host "Setting workspace Destination: $GoogleWorkSpaceDestination"
Invoke-Expression "$GamDir\gam.exe select $GoogleWorkSpaceDestination save" # swap/set google workspace
Invoke-Expression "$GamDir\gam.exe info domain"
start-sleep 3

Write-Host "School prefix: $SchoolCode"

Write-Host "Getting all users and OUs from supplied source: $GoogleSourceBaseOU"
$SourceGoogleOus = @()
$SourceGoogleOUs = Invoke-Expression "$GamDir\Gam.exe print orgs from parent $GoogleSourceBaseOU"

foreach ($SourceGoogleOU in $GoogleOUs) {
    Write-Host "Processing: $($SourceGoogleOU)"
}


DashedLine

