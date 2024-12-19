$mnspver = "0.0.10"

Write-Host $(Get-Date)
Write-Host "MNSP Version" $mnspver
Start-Sleep 10
$ErrorActionPreference="Continue"
Set-Location $GamDir

function DashedLine {
Write-host "-----------------------------------------------------------`n"
}

DashedLine

#Set Google Instance: Destination...
Write-Host "###### Set Google instance: Destination... ######"

$GoogleSvcAccount = $GoogleWorkspaceMNSPsvcAccount
Write-Host "Google Destination Service Account: $GoogleSvcAccount"

Write-Host "Setting workspace Destination: $GoogleWorkSpaceDestination"
Invoke-Expression "$GamDir\gam.exe select $GoogleWorkSpaceDestination save" # swap/set google workspace
Invoke-Expression "$GamDir\gam.exe"
start-sleep 3
DashedLine

#get data
if exist check & remove $tempcsv1
if (test-path $tempcsv1) { remove-item $tempcsv1 -force -verbose }

Write-Host "downloading gsheet ID: $GoogleSheetID"
Write-Host "$GamDir\gam.exe user $GoogleWorkspaceMNSPsvcAccount get drivefile $GoogleSheetID format csv targetfolder $DataDir targetname $tempcsv1" 
Invoke-Expression "$GamDir\gam.exe user $GoogleWorkspaceMNSPsvcAccount get drivefile $GoogleSheetID format csv targetfolder $DataDir targetname $tempcsv1"

Start-sleep 2

$gsheetsData = import-csv $tempcsv1
Write Host "Number of records matching selection criteria:" $gsheetsData.count



