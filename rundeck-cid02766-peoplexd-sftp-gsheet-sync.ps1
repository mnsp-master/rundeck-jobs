$mnspver = "0.0.25"

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
#if exist check & remove $tempcsv1
#if (test-path $tempcsv1) { remove-item $tempcsv1 -force -verbose }
remove-item $DataDir\*.csv -force -verbose

Write-Host "downloading gsheet ID: $GoogleSheetID"
Write-Host "$GamDir\gam.exe user $GoogleWorkspaceMNSPsvcAccount get drivefile $GoogleSheetID format csv targetfolder $DataDir targetname $tempcsv1" 
Invoke-Expression "$GamDir\gam.exe user $GoogleWorkspaceMNSPsvcAccount get drivefile $GoogleSheetID format csv targetfolder $DataDir targetname $tempcsv1"

Start-sleep 1

Write-Host "scp all csv's from SFTP server to local data folder..."
Invoke-Expression "scp.exe -s $SecureCopyCmd $DataDir"
start-sleep 1
Get-ChildItem $DataDir -filter *.csv -recurse

$gsheetsData = import-csv $tempcsv1
Write Host "Number of rows to process:" $gsheetsData.count

foreach ( $report in $gsheetsData) {
    $GoogleSheetID = $report."Google sheet ID"
    $SourceSFTPfilename = $report."Source SFTP filename"
    $Environment = $report."Production/UAT"
    $GoogleSheetReportName = $report."Google Sheet Report name"
    $SourceSFTPFileNameComplete = "$Datadir\$SourceSFTPfilename"
    Write-Host "Google Sheet ID:" $GoogleSheetID
    write-host "Source SFTP filename:" $SourceSFTPfilename
    write-host "Environment:" $Environment
    write-host "Google sheet Report name:" $GoogleSheetReportName

    start-sleep 20

        #confirm file exists and is > 0 bytes
        if ((Get-Item $SourcesSFTPFileNameCompete).length -gt 0){
            Invoke-Expression "$GamDir\gam.exe user $GoogleWorkspaceMNSPsvcAccount update drivefile id $GoogleSheetID localfile $SourceSFTPFileNameComplete newfilename '$GoogleSheetReportName as of $(get-date)' "
        } else {
            Write-Warning "File: $SourcesSFTPFileNameCompete size 0 Bytes not proceeding with replacement of gsheet: $GoogleSheetID"
        }
    DashedLine
}
