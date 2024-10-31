$mnspver = "0.0.77"

Write-Host $(Get-Date)
Write-Host "MNSP Version" $mnspver
Start-Sleep 10
$ErrorActionPreference="Continue"
Set-Location $GamDir

function DashedLine {
Write-host "-----------------------------------------------------------`n"
}

#Write-Host "creating destination gfolder for all existing shared drive association user reports..."
#$GfolderReportsID = @()
#$GfolderReportsID = $(Invoke-Expression "$GamDir\gam.exe user $GoogleSourceSvcAccount create drivefile drivefilename '$ReportsFolderName' mimetype gfolder parentid $ReportsFolderParentID returnidonly")
$GfolderReportsID = "1X4xdjK5fJLnXn5Q1sqqhBpKYJ1KJHVc9" #update/enable once in production

$GoogleSourceSvcAccount = ("$GoogleServiceAccountPrefix" + "$GoogleWorkSpaceSource" + "@" + "$GGoogleWorkspaceSourceMailDomain")
Write-Host "Google Source Service Account: $GoogleSourceSvcAccount"

#get verified user data
#if exist check & remove $tempcsv4
if (test-path $tempcsv4) { remove-item $tempcsv4 -force -verbose }

Write-Host "downloading gsheet ID: $GoogleSheetID tab: $GoogleSheetTab01"
Invoke-Expression "$GamDir\gam.exe user $GoogleSourceSvcAccount get drivefile $GoogleSheetID format csv gsheet ""$GoogleSheetTab01"" targetfolder $DataDir targetname $tempcsv4"

Start-sleep 2

$VerifiedUserData = Get-Content -path $tempcsv4 | select-object -skip 1 | convertFrom-csv | where { $_.$FieldMatch01 -like $FieldString } #import where field like $FieldMatch01, and skip 1st line
Write Host "Number of records matching selection criteria:" $VerifiedUserData.count

#if ($uuids.Contains($uuid)) { } # if var is in array

#Replacement google instance...
Write-Host "Setting workspace Destination: $GoogleWorkSpaceDestination"
Invoke-Expression "$GamDir\gam.exe select $GoogleWorkSpaceDestination save" # swap/set google workspace
Invoke-Expression "$GamDir\gam.exe"
start-sleep 3
DashedLine

foreach ($user in $VerifiedUserData) {
    DashedLine
    $LegacyUserMail= $user."Existing Email Address" #current mail address
    $HRid = $user."Staff full name" # HR id
    $FirstName = $user."Staff first name" #prefered firstname
    $LastName = $user."Staff Surname"
    $ReplacementUserMail = $user."new email"

    #dev enviornment...
    if ( $RunDeckDev -eq "true" ) {
        Write-Host "Setting random dev mail address..."
        $RundeckDevMail = ("SNO-" + $([int64](Get-Date -UFormat %s)) + "@" + "$GoogleWorkspaceDestinationMailDomain")
        $ReplacementUserMail = $RundeckDevMail
        }

    Write-Host "Processing: $ReplacementUserMail"
    Write-Host "HR ID: $HRid"
    Write-Host "Firstname: $FirstName"
    Write-Host "Lastname: $LastName"

    Write-Host "create destination account..."
    #Write-Host "$GamDir\gam.exe create user $ReplacementUserMail firstname $FirstName lastname $LastName password random 16 org '$GoogleWorkspaceDestinationUserOU'"
    Invoke-Expression "$GamDir\gam.exe create user $ReplacementUserMail firstname $FirstName lastname $LastName password random 16 org '$GoogleWorkspaceDestinationUserOU'"

    Write-Host "hide account from GAL.."
    Write-Host "$GamDir\gam.exe update user $ReplacementUserMail gal false"
    Invoke-Expression "$GamDir\gam.exe update user $ReplacementUserMail gal false"

    Write-Host "generate MFA backup codes..."
    Invoke-Expression "$GamDir\gam.exe user $ReplacementUserMail update backupcodes"

    Write-Host "update Replacement account..."
    Write-Host "Invoke-Expression $GamDir\gam.exe update user $ReplacementUserMail $GoogleCustomAttribute01 $HRid" #set HR ID

    DashedLine
}


#legacy google instance...
Write-Host "Setting workspace source: $GoogleWorkSpaceSource"
Invoke-Expression "$GamDir\gam.exe select $GoogleWorkSpaceSource save" # swap/set google workspace
Invoke-Expression "$GamDir\gam.exe"
DashedLine

#$GoogleWorkspaceSourceGroup = ("$GoogleWorkspaceSourceGroupPrefix" + "@" + "$GGoogleWorkspaceSourceMailDomain")
#Write-Host "Getting members of users to process source group $GoogleWorkspaceSourceGroup"
#Invoke-Expression "$GamDir\gam.exe print group-members group_ns $GoogleWorkspaceSourceGroup > $tempcsv"

foreach ($user in $VerifiedUserData) {
    DashedLine
    $LegacyUserMail = $user."Existing Email Address" #current mail address
    $HRid = $user."Staff full name" # HR id
    $FirstName = $user."Staff first name" #prefered firstname
    $LastName = $user."Staff Surname"
    $ReplacementUserMail = $user."new email"

    if ( $RunDeckDev -eq "true" ) { $ReplacementUserMail = $RundeckDevMail } # Dev environment
    
    Write-Host "Processing: $LegacyUserMail"
    Write-Host "HR ID: $HRid"
    Write-Host "Firstname: $FirstName"
    Write-Host "Lastname: $LastName"

    Write-Host "update legacy accounts..."
    Invoke-Expression "$GamDir\gam.exe update user $LegacyUserMail $GoogleCustomAttribute01 $HRid" #set HR ID

    Write-Host "send current calendar invite..."
    Invoke-Expression "$GamDir\gam.exe calendar $LegacyUserMail add acls reader $ReplacementUserMail sendnotifications false"

    Write-Host "report current shared drive folder associations for: $LegacyUserMail ..."
    Invoke-expression "$GamDir\gam.exe user $legacyUserMail print teamdrives todrive tdparent id:$GfolderReportsID tdnobrowser tdtitle '$LegacyUserMail shared drives summary as of $(get-date)'"

    Write-Host "shared drive creation (Legacy Source to Destination user)..."
    $TeamDriveName = "DRA DEV $(Get-Date)" #convention needs confirming
    $LegacyUserTeamDriveID = "$( Invoke-expression $GamDir\gam.exe user $GoogleSourceSvcAccount create teamdrive '$TeamDriveName' adminmanagedrestrictions true asadmin returnidonly )"

    Write-Host "move to data move enabled OU..."
    #Write-Host "Invoke-expression "$GamDir\gam.exe update teamdrive $LegacyUserTeamDriveID asadmin ou $LegacyUserTeamDriveOU"" #location needs confirming

    Write-Host "Allow outside sharing..."
    Invoke-expression "$GamDir\gam.exe update teamdrive $LegacyUserTeamDriveID asadmin domainUsersOnly False"

    Write-Host "Allow people who aren't shared drive members to access files - false..."
    Invoke-expression "$GamDir\gam.exe update teamdrive $LegacyUserTeamDriveID asadmin sharingFoldersRequiresOrganizerPermission True"

    Write-Host "Add internal user as manager..."
    Invoke-expression "$GamDir\gam.exe add drivefileacl $LegacyUserTeamDriveID user $LegacyUserMail role organizer"
    
    Write-Host "Add external user as manager..."
    Invoke-expression "$GamDir\gam.exe add drivefileacl $LegacyUserTeamDriveID user $ReplacementUserMail role organizer"

    DashedLine
}

#Replacement google instance...
Write-Host "Setting workspace Destination: $GoogleWorkSpaceDestination"
Invoke-Expression "$GamDir\gam.exe select $GoogleWorkSpaceDestination save" # swap/set google workspace
Invoke-Expression "$GamDir\gam.exe"
start-sleep 3
DashedLine

foreach ($user in $VerifiedUserData) {
    DashedLine
    $LegacyUserMail = $user."Existing Email Address" #current mail address
    $HRid = $user."Staff full name" # HR id
    $FirstName = $user."Staff first name" #prefered firstname
    $LastName = $user."Staff Surname"
    $ReplacementUserMail = $user."new email"

    if ( $RunDeckDev -eq "true" ) { $ReplacementUserMail = $RundeckDevMail } # Dev environment
    
    Write-Host "Processing: $LegacyUserMail"
    Write-Host "HR ID: $HRid"
    Write-Host "Firstname: $FirstName"
    Write-Host "Lastname: $LastName"

    Write-Host "Accept calendar invite..."
    Invoke-Expression "$GamDir\gam.exe user $LegacyUserMail add calendar $ReplacementUserMail selected true"

}
<#


#>