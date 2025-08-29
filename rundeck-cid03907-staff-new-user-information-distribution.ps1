$mnspver = "0.0.15"

Write-Host $(Get-Date)
Write-Host "MNSP Version" $mnspver
Start-Sleep 10
$ErrorActionPreference="Continue"
Set-Location $GamDir

function DashedLine {
Write-host "-----------------------------------------------------------`n"
}

#Set Google Instance: Destination...
Write-Host "###### Set Google instance: Destination... ######"

$GoogleSvcAccount = $GoogleWorkspaceMNSPsvcAccount
Write-Host "Google Destination Service Account: $GoogleSvcAccount"

Write-Host "Setting workspace Destination: $GoogleWorkSpaceDestination"
Invoke-Expression "$GamDir\gam.exe select $GoogleWorkSpaceDestination save" # swap/set google workspace
Invoke-Expression "$GamDir\gam.exe"
start-sleep 3
DashedLine



#get verified user data
#if exist check & remove $tempcsv4
if (test-path $tempcsv4) { remove-item $tempcsv4 -force -verbose }

Write-Host "downloading gsheet ID: $GoogleSheetID tab: $GoogleSheetTab01"
Invoke-Expression "$GamDir\gam.exe user $GoogleSvcAccount get drivefile $GoogleSheetID format csv targetfolder $DataDir targetname $tempcsv4"
#Write-Host "Invoke-Expression $GamDir\gam.exe user $GoogleSourceSvcAccount get drivefile $GoogleSheetID format csv gsheet ""$GoogleSheetTab01"" targetfolder $DataDir targetname $tempcsv4"

Start-sleep 2

#$VerifiedUserData = Get-Content -path $tempcsv4 | select-object -skip 1 | convertFrom-csv | where { $_.$FieldMatch01 -like $FieldString } #import where field like $FieldMatch01, and skip 1st line
$VerifiedUserData = import-csv -path $tempcsv4
Write-Host "Number of records matching selection criteria:" $VerifiedUserData.count
#TODO - if count 0 break out of script...
#$VerifiedUserData

#download html template
#if exist check & remove $tempcsv4
#if (test-path $temphtml1) { remove-item $temphtml1 -force -verbose }
#Invoke-Expression "$GamDir\gam.exe user $GoogleSvcAccount get drivefile $GoogleDocID format html targetfolder $DataDir targetname $temphtml1"

Start-Sleep 2

Write-Host "Processing users in destination..."
foreach ($user in $VerifiedUserData) {
    DashedLine
    $LegacyUserMail = $user."legacy email" #current mail address
    $HRid = $user."HR ID" # HR id
    $FirstName = $user."First Name" #prefered firstname
    $lastname = $user."Last Name"
    $ReplacementUserMail = $user."email"
    $Password = $User."password"

    Write-Host "Processing: $LegacyUserMail"
    Write-Host "HR ID: $HRid"
    Write-Host "Firstname: $FirstName"
    Write-Host "Lastname: $lastname"

    #generate MFA backup codes
    #$userBackupCodes = invoke-expression "$GamDir\gam.exe user $ReplacementUserMail update backupcodes"
    Write-host "$GamDir\gam.exe user $ReplacementUserMail update backupcodes | ForEach-Object { $_ -replace '^\s*\d+:\s*', '' }"
    $userBackupCodes = Invoke-expression "$GamDir\gam.exe user $ReplacementUserMail update backupcodes" | ForEach-Object { $_ -replace '^\s*\d+:\s*', '' } #cleanup output

    #send mail(s)
    ##backup codes...
    Write-Host "$GamDir\gam.exe sendemail $legacyUserMail from $GoogleWorkspaceSenderMail subject 'MNSP - MFA Backup Codes as of $(get-Date)' message '$userBackupCodes'"
    Invoke-expression "$GamDir\gam.exe sendemail $legacyUserMail from $GoogleWorkspaceSenderMail subject 'MNSP - MFA Backup Codes as of: $(get-Date)' message '$userBackupCodes'"
    
    ##credentials...
    Write-Host "$GamDir\gam.exe sendemail $legacyUserMail from $GoogleWorkspaceSenderMail subject 'MNSP - As Advised $(get-Date)' message '$password'"
    Invoke-Expression "$GamDir\gam.exe sendemail $legacyUserMail from $GoogleWorkspaceSenderMail subject 'MNSP - As Advised as of: $(get-Date)' message '$password'"

    ##account information...
    Write-Host "$GamDir\gam.exe sendemail $legacyUserMail from $GoogleWorkspaceSenderMail newuser $ReplacementUserMail firstname $FirstName LastName $LastName password 'Sent in another email'"
    Invoke-Expression "$GamDir\gam.exe sendemail $legacyUserMail from $GoogleWorkspaceSenderMail newuser $ReplacementUserMail firstname $FirstName LastName $LastName password 'Sent in another email'"

    ##General Information
    $htmlContent = get-content $temphtml1
    $htmlContent = $htmlContent.replace('REPLACE_FIRSTNAME',$FirstName)
    $htmlContent = $htmlContent.replace('REPLACE_DOMAIN',$GoogleWorkspaceDestinationMailDomain)
    $htmlContent = $htmlContent.replace('REPLACE_EMAIL',$ReplacementUserMail)

    #Write-Host "$GamDir\gam.exe sendemail $legacyUserMail from $GoogleWorkspaceSenderMail subject 'Your New MNSP email account Information as of: $(Get-date)' htmlmessage '$htmlContent'"
    invoke-expression "$GamDir\gam.exe sendemail $legacyUserMail from $GoogleWorkspaceSenderMail subject 'Your New MNSP email account Information summary as of: $(Get-date)' htmlmessage '$htmlContent'"

    #>

    DashedLine
}

#>