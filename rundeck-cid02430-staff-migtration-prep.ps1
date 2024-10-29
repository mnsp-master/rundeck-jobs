$mnspver = "0.0.58"

Write-Host $(Get-Date)
Write-Host "MNSP Version" $mnspver
Start-Sleep 10
$ErrorActionPreference="Continue"
Set-Location $GamDir

function DashedLine {
Write-host "-----------------------------------------------------------`n"
}

Write-Host "Setting workspace source: $GoogleWorkSpaceSource"
Invoke-Expression "$GamDir\gam.exe select $GoogleWorkSpaceSource save" # swap/set google workspace
Invoke-Expression "$GamDir\gam.exe"
DashedLine

$GoogleSourceSvcAccount = ("$GoogleServiceAccountPrefix" + "$GoogleWorkSpaceSource" + "@" + "$GGoogleWorkspaceSourceMailDomain")
Write-Host "Google Source Service Account: $GoogleSourceSvcAccount"

$GoogleWorkspaceSourceGroup = ("$GoogleWorkspaceSourceGroupPrefix" + "@" + "$GGoogleWorkspaceSourceMailDomain")
Write-Host "Getting members of users to process source group $GoogleWorkspaceSourceGroup"
Invoke-Expression "$GamDir\gam.exe print group-members group_ns $GoogleWorkspaceSourceGroup > $tempcsv"

#get verified user data
#if exist check & remove $tempcsv4
if (test-path $tempcsv4) { remove-item $tempcsv4 -force -verbose }

Start-sleep 2

#Write-Host "creating destination gfolder for all existing shared drive association user reports..."
#$GfolderReportsID = @()
#$GfolderReportsID = $(Invoke-Expression "$GamDir\gam.exe user $GoogleSourceSvcAccount create drivefile drivefilename '$ReportsFolderName' mimetype gfolder parentid $ReportsFolderParentID returnidonly")
$GfolderReportsID = "1X4xdjK5fJLnXn5Q1sqqhBpKYJ1KJHVc9" #update once in production

Write-Host "downloading gsheet ID: $GoogleSheetID tab: $GoogleSheetTab01"
Invoke-Expression "$GamDir\gam.exe user $GoogleSourceSvcAccount get drivefile $GoogleSheetID format csv gsheet ""$GoogleSheetTab01"" targetfolder $DataDir targetname $tempcsv4"

$VerifiedUserData = Get-Content -path $tempcsv4 | select-object -skip 1 | convertFrom-csv | where { $_.$FieldMatch01 -like $FieldString } #import where field like $FieldMatch01, and skip 1st line
Write Host "Number of records matching selection criteria:" $VerifiedUserData.count

#if ($uuids.Contains($uuid)) { } # if var is in array
#legacy google instance...
foreach ($user in $VerifiedUserData) {
    DashedLine
    $LegacyUserMail = $user."Existing Email Address" #current mail address
    $HRid = $user."Staff full name" # HR id
    $FirstName = $user."Staff first name" #prefered firstname
    $LastName = $user."Staff Surname"
    $ReplacementUserMail = $user."new email"

    Write-Host "Processing: $LegacyUserMail"
    Write-Host "HR ID: $HRid"
    Write-Host "Firstname: $FirstName"
    Write-Host "Lastname: $LastName"

    #update legacy accounts...
    Write-Host "Invoke-Expression $GamDir\gam.exe user $LegacyUserMail $GoogleCustomAttribute01 $HRid" #set HR ID

    #send current calendar invite...
    Write-Host "Invoke-Expression $GamDir\gam.exe calendar $LegacyUserMail add acls reader $ReplacementUserMail sendnotifications false"

    #report current shared drive folder associations...
    #Invoke-expression "$GamDir\gam.exe user $legacyUserMail print teamdrives todrive tdparent id:$GfolderReportsID tdnobrowser tdtitle '$LegacyUserMail shared drives summary as of $(get-date)'"

    #shared drive creation (Legacy Source to Destination user)
    $TeamDriveName = "HEM DEV $(Get-Date)" #convention needs confirming
    Write-Host "$LegacyUserTeamDriveID = $( Invoke-expression "$GamDir\gam.exe user $GoogleSourceSvcAccount create teamdrive '$TeamDriveName' adminmanagedrestrictions true asadmin returnidonly")"

    #move to data move enabled OU...
    Write-Host "Invoke-expression "$GamDir\gam.exe update teamdrive $LegacyUserTeamDriveID asadmin ou $LegacyUserTeamDriveOU""

    #Allow outside sharing...
    Write-Host "Invoke-expression "$GamDir\gam.exe update teamdrive $LegacyUserTeamDriveID asadmin domainUsersOnly False""

    #Allow people who aren't shared drive members to access files - false:
    Write-Host "Invoke-expression "$GamDir\gam.exe update teamdrive $LegacyUserTeamDriveID asadmin sharingFoldersRequiresOrganizerPermission True""

    #Add internal user as manager
    Write-Host "Invoke-expression "$GamDir\gam.exe add drivefileacl $LegacyUserTeamDriveID user $LegacyUserMail role organizer""
    
    #Add external user as manager
    Write-Host "Invoke-expression "$GamDir\gam.exe add drivefileacl $LegacyUserTeamDriveID user $ReplacementUserMail role organizer""

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
    $LegacyUserMail= $user."Existing Email Address" #current mail address
    $HRid = $user."Staff full name" # HR id
    $FirstName = $user."Staff first name" #prefered firstname
    $LastName = $user."Staff Surname"
    $ReplacementUserMail = $user."new email"

    Write-Host "Processing: $ReplacementUserMail"
    Write-Host "HR ID: $HRid"
    Write-Host "Firstname: $FirstName"
    Write-Host "Lastname: $LastName"

    #create destination accounts...
    Write-Host "Invoke-Expression $GamDir\gam.exe create user $ReplacementUserMail firstname $FirstName lastname $LastName password random 16 org $GoogleWorkspaceDestinationUserOU"

    #hide accounts from GAL..
    Write-Host "Invoke-Expression $GamDir\gam.exe update user $ReplacementUserMail gal false"

    #generate MFA backup codes...
    Write-Host "Invoke-Expression $GamDir\gam.exe user $ReplacementUserMail update backupcodes"
    
    #Accept calendar invite...
    Write-Host "Invoke-Expression $GamDir\gam.exe user $LegacyUserMail add calendar $ReplacementUserMail selected true"

    #update Replacement accounts...
    Write-Host "Invoke-Expression $GamDir\gam.exe user $ReplacementUserMail $GoogleCustomAttribute01 $HRid" #set HR ID

    DashedLine
}


<#


#>