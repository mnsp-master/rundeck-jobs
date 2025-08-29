$mnspver = "0.0.1"

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
Invoke-Expression "$GamDir\gam.exe user $GoogleSvcAccount get drivefile $GoogleSheetID format csv gsheet ""$GoogleSheetTab01"" targetfolder $DataDir targetname $tempcsv4"
#Write-Host "Invoke-Expression $GamDir\gam.exe user $GoogleSourceSvcAccount get drivefile $GoogleSheetID format csv gsheet ""$GoogleSheetTab01"" targetfolder $DataDir targetname $tempcsv4"

Start-sleep 2

$VerifiedUserData = Get-Content -path $tempcsv4 | select-object -skip 1 | convertFrom-csv | where { $_.$FieldMatch01 -like $FieldString } #import where field like $FieldMatch01, and skip 1st line
Write-Host "Number of records matching selection criteria:" $VerifiedUserData.count
#TODO - if count 0 break out of script...
#$VerifiedUserData

Start-Sleep 10


Write-Host "Processing users in destination..."
foreach ($user in $VerifiedUserData) {
    DashedLine
    $LegacyUserMail = $user."Existing Email Address" #current mail address
    $HRid = $user."Person ID" # HR id
    $FirstName = $user."Staff first name" #prefered firstname
    $lastname = $user."Staff Surname"
    $ReplacementUserMail = $user."new email"

    Write-Host "Processing: $LegacyUserMail"
    Write-Host "HR ID: $HRid"
    Write-Host "Firstname: $FirstName"
    Write-Host "Lastname: $lastname"

    #password generator...
        $pwd = @()
        try {
        $pwd = $(Invoke-WebRequest -Uri $PwdWebRequestURI -UseBasicParsing)
        $Password = $($pwd.content)
        }
            catch {
            Write-Error "No Webserver, or pwd received"
            $password = $pwdFailsafe
        }

    Write-Host "process account..."
    Write-Host "$GamDir\gam.exe create user $ReplacementUserMail firstname '$FirstName' lastname '$LastName' org '$GoogleWorkspaceDestinationUserOU' $GoogleCustomAttribute01 $HRid password $password gal $GoogleIncludeInGal changepasswordatnextlogin True"
    #Invoke-Expression "$GamDir\gam.exe create user $ReplacementUserMail firstname '$FirstName' lastname '$LastName' org '$GoogleWorkspaceDestinationUserOU' $GoogleCustomAttribute01 $HRid password '$password' gal $GoogleIncludeInGal changepasswordatnextlogin True" ###create user #CID00#### dry run
    #Start-Sleep 5 #allow time for user creation

    #capture initial credentials
    #"$firstname,$lastname,$legacyUserMail,$ReplacementUserMail,$password,$HRid" | out-file -filepath $tempcsv2 -Append

    #generate MFA backup codes
    #$userBackupCodes = invoke-expression "$GamDir\gam.exe user $ReplacementUserMail update backupcodes"
    Write-host "$GamDir\gam.exe user $ReplacementUserMail update backupcodes | ForEach-Object { $_ -replace '^\s*\d+:\s*', '' }"
    $userBackupCodes = Invoke-expression "$GamDir\gam.exe user $ReplacementUserMail update backupcodes" | ForEach-Object { $_ -replace '^\s*\d+:\s*', '' } #cleanup output

   
    #send mail(s)
    ##backup codes...
    Write-Host "$GamDir\gam.exe sendemail $legacyUserMail from $GoogleWorkspaceSenderMail subject 'MFA Backup Codes as of $(get-Date)' message '$userBackupCodes'"
    Invoke-expression "$GamDir\gam.exe sendemail $legacyUserMail from $GoogleWorkspaceSenderMail subject 'MFA Backup Codes as of: $(get-Date)' message '$userBackupCodes'"
    
    ##credentials...
    Write-Host "$GamDir\gam.exe sendemail $legacyUserMail from $GoogleWorkspaceSenderMail subject 'As Advised $(get-Date)' message '$password'"
    Invoke-Expression "$GamDir\gam.exe sendemail $legacyUserMail from $GoogleWorkspaceSenderMail subject 'As Advised as of: $(get-Date)' message '$password'"

    ##account information...
    Write-Host "$GamDir\gam.exe sendemail $legacyUserMail from $GoogleWorkspaceSenderMail newuser $ReplacementUserMail firstname $FirstName LastName $LastName password 'Sent in another email'"
    Invoke-Expression "$GamDir\gam.exe sendemail $legacyUserMail from $GoogleWorkspaceSenderMail newuser $ReplacementUserMail firstname $FirstName LastName $LastName password 'Sent in another email'"
    #>

    DashedLine
}

#>