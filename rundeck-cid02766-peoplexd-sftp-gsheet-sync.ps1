$mnspver = "0.0.42"

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

#get source data

remove-item $DataDir\*.csv -force -verbose

Write-Host "downloading gsheet ID: $GoogleSheetID"
#Write-Host "$GamDir\gam.exe user $GoogleWorkspaceMNSPsvcAccount get drivefile $GoogleSheetID format csv targetfolder $DataDir targetname $tempcsv1" 
Invoke-Expression "$GamDir\gam.exe user $GoogleWorkspaceMNSPsvcAccount get drivefile $GoogleSheetID format csv targetfolder $DataDir targetname $tempcsv1"

Start-sleep 1

Write-Host "scp all csv's from SFTP server to local data folder..."
Invoke-Expression "scp.exe -s $SecureCopyCmd $DataDir"
start-sleep 1
Get-ChildItem $DataDir -filter *.csv -recurse #list all csv's in $DataDir...

$gsheetsData = import-csv $tempcsv1 #create array of csv's/gsheets to process...
#Write Host "Number of rows to process:" $gsheetsData.count

foreach ( $report in $gsheetsData) {
    $GoogleSheetID = $report."Google sheet ID"
    $SourceSFTPfilename = $report."Source SFTP filename"
    $Environment = $report."Production/UAT"
    $GoogleSheetReportName = $report."Google Sheet Report name"
    $SourceSFTPFileNameComplete = "$Datadir\$SourceSFTPfilename"
    DashedLine
    Write-Host "Google Sheet ID:" $GoogleSheetID
    write-host "Source SFTP filename:" $SourceSFTPfilename
    write-host "Environment:" $Environment
    write-host "Google sheet Report name:" $GoogleSheetReportName

        #confirm desired downloaded csv exists and is > 0 bytes
        if ((Get-Item $SourceSFTPFileNameComplete).length -gt 0){
            Write-Host "File: $SourceSFTPFileNameComplete size: $((Get-Item $SourceSFTPFileNameComplete).length) bytes - proceeding with gsheet replacement"

                $SEL = get-content $SourceSFTPFileNameComplete
                if ( $SEL -imatch ";") {
                    DashedLine
                    Write-Warning ";'s present"
                    get-content $SourceSFTPFileNameComplete | % { if($_ -match ";") {write-host $_}}
                    #get-content $SourceSFTPFileNameComplete | % { if($_ -match ";") {write-host $_ | out-file $GmailAttachment }}
                    Invoke-Expression ".\gam.exe sendemail $GmailRecipient subject '$GmailSubject as of $(get-date)' attach $transcriptlog"

                    #ENHANCEMENT - send offending line(s) to nominated mail recpient(s)
                } else {
                   DashedLine
                   Write-Host "No ;'s found in csv, replacing with all ÿ with ; in gsheet: $GoogleSheetID"
                   (get-content $SourceSFTPFileNameComplete) | ForEach-Object {$_ -replace 'ÿ',';'} | Out-File $SourceSFTPFileNameComplete
                   Invoke-Expression ".\gam.exe user $GoogleWorkspaceMNSPsvcAccount update drivefile id $GoogleSheetID localfile $SourceSFTPFileNameComplete newfilename '$GoogleSheetReportName as of $(get-date)' columndelimiter ';'"

                }

            #Invoke-Expression "$GamDir\gam.exe user $GoogleWorkspaceMNSPsvcAccount update drivefile id $GoogleSheetID localfile $SourceSFTPFileNameComplete newfilename '$GoogleSheetReportName as of $(get-date)' "
            #alternate delimeter as souce data has commas within fields...
            #Invoke-Expression "$GamDir\gam.exe user $GoogleWorkspaceMNSPsvcAccount update drivefile id $GoogleSheetID localfile $SourceSFTPFileNameComplete newfilename '$GoogleSheetReportName as of $(get-date)' columndelimiter 'ÿ'"
        } else {
            Write-Warning "File: $SourceSFTPFileNameComplete size 0 Bytes not proceeding with replacement of gsheet: $GoogleSheetID"
        }

    DashedLine
}
