$mnspver = "0.0.32"

Write-Host $(Get-Date)
Write-Host "MNSP Version" $mnspver
Start-Sleep 10
$ErrorActionPreference="Continue"
Set-Location $GamDir

function DashedLine {
Write-host "-----------------------------------------------------------`n"
}

Write-Host "gsheet Student number column heading:" $FieldMatch01
Start-sleep 10

#prepare user details csv
Write-Host "emptying $tempcsv2 of any existing data..."
Clear-Content $tempcsv2
sleep 1
$UserInfoCSVheader | out-file -filepath $tempcsv2 -Append #create blank csv with simple header

#Set Google Instance: Destination...
Write-Host "###### Set Google instance: Destination... ######"

$GoogleSvcAccount = $GoogleWorkspaceMNSPsvcAccount
Write-Host "Google Destination Service Account: $GoogleSvcAccount"

Write-Host "Setting workspace Destination: $GoogleWorkSpaceDestination"
Invoke-Expression "$GamDir\gam.exe select $GoogleWorkSpaceDestination save" # swap/set google workspace
Invoke-Expression "$GamDir\gam.exe"
start-sleep 3
DashedLine

Write-Host "Report on all current users from base OU: $GoogleWorkspaceSourceUserOU"
Invoke-expression "$GamDir\gam.exe ou_and_children '$GoogleWorkspaceSourceUserOU' print allfields todrive tdparent id:$GfolderReportsID tdtitle 'User info - Pre-Migration for domain: $GoogleWorkspaceSourceMailDomain as of: $(Get-date)'"

#get verified user data
#if exist check & remove $tempcsv4
if (test-path $tempcsv4) { remove-item $tempcsv4 -force -verbose }

Write-Host "downloading gsheet ID: $GoogleSheetID tab: $GoogleSheetTab01"
#Write-Host "$GamDir\gam.exe user $GoogleSvcAccount get drivefile $GoogleSheetID format csv gsheet ""$GoogleSheetTab01"" targetfolder $DataDir targetname $tempcsv4"
Invoke-Expression "$GamDir\gam.exe user $GoogleSvcAccount get drivefile $GoogleSheetID format csv gsheet ""$GoogleSheetTab01"" targetfolder $DataDir targetname $tempcsv4"
#Write-Host "Invoke-Expression $GamDir\gam.exe user $GoogleSourceSvcAccount get drivefile $GoogleSheetID format csv gsheet ""$GoogleSheetTab01"" targetfolder $DataDir targetname $tempcsv4"

Start-sleep 2

Write-Host "Field match:  " $FieldMatch01
Write-Host "Field String: " $FieldString
#Write-Host "$VerifiedUserData = Get-Content -path $tempcsv4 | select-object -skip 1 | convertFrom-csv | where { $_.$FieldMatch01 -like $FieldString } "
#$VerifiedUserData = Get-Content -path $tempcsv4 | convertFrom-csv | where { $_.$FieldMatch01 -like $FieldString } #import where field like $FieldMatch01

$VerifiedUserData = Get-Content -path $tempcsv4 | convertFrom-csv | where-object { 
    $_.$FieldMatch01 -like $FieldString -and 
    $_.$Fieldmatch02 -match '^[0-9]+$' #Numeric values only - excludes - R N1 N2 etc
    } #import where field like $FieldMatch01

#$VerifiedUserData = Get-Content -path $tempcsv4 | select-object -skip 1 | convertFrom-csv | where { $_.$FieldMatch01 -like $FieldString } #import where field like $FieldMatch01, and skip 1st line
Write Host "Number of records matching selection criteria:" $VerifiedUserData.count
#TODO - if count 0 break out of script...


#create user info destination gsheet
$UserInfoGsheetID = $(Invoke-Expression "$GamDir\gam.exe user $GoogleSvcAccount create drivefile drivefilename '$GoogleWorkspaceDestinationMailDomain User Info' mimetype gsheet parentid $GfolderReportsID returnidonly")

if (test-path $tempcsv6) { remove-item $tempcsv6 -force -verbose }
if (test-path $tempcsv8) { remove-item $tempcsv8 -force -verbose }

start-sleep 2

Write-Host "Updating users in destination..."
foreach ($user in $VerifiedUserData) {
    DashedLine
    $LegacyUserMail = $user."Email Address (Main)" #current mail address
    $UPN = $user."UPN" # student UID (encrypted UPN) ## UPDATE NEEDED ##
    $FirstName = $user."Modified_Preferred_firstname" #prefered firstname ## UPDATE NEEDED ##
    $LastName = $user."Modified_Preferred_Lastname" ## UPDATE NEEDED ##
    $ReplacementUserMail = $user."new email"
    $DestOU = [int] $user."NC Year(s) for today" #set var as interger

    #add leading zero if required: to create consitent OUs YEAR07 not YEAR7: 
        if ( $DestOU -le 9) {
            Write-host "Target Year group less than or equal to 9..."
            $UpdatedDestOU = @()
            $UpdatedDestOU = $("Year" + "0" + $DestOU)
            } else {
            $UpdatedDestOU = $("Year" + $DestOU)
            }
        

    #script dev check...
    #if ( $RunDeckDev -eq "true" ) {
    #    Write-Host "Setting random dev mail address..."
    #    $RundeckDevMail = ("SNO-" + $([int64](Get-Date -UFormat %s)) + "@" + "$GoogleWorkspaceDestinationMailDomain")
    #    $ReplacementUserMail = $RundeckDevMail
    #    }

    Write-Host "Processing (SRC email): $LegacyUserMail"
    Write-Host "Processing (DEST email): $ReplacementUserMail"
    Write-Host "UPN: $UPN"
    Write-Host "Firstname: $FirstName"
    Write-Host "Lastname: $LastName"
    Write-Host "Source Year: $DestOU" 
    Write-Host "Destination OU name: $UpdatedDestOU"

    <#
    Write-Host "Generating Random Password..." 
        $pwd = $(Invoke-WebRequest -Uri $PwdWebRequestURI -UseBasicParsing)
        #    $pwd.Content
            #$pwd.StatusCode
                if ($pwd.StatusCode -eq 200) {
                Write-Host "proceed with pwd reservation"
                $password = $($pwd.Content)
                #Write-Host "Password: " $password
                } else {
                Write-Error "No Webserver, or pwd received"
                $password = $PwdFailsafe
                }

        start-sleep 1
    #>
#>
  

  Write-Host "Checking if legacy mail: $LegacyUserMail  like: $GoogleWorkspaceSourceMailDomain"
    if ( $LegacyUserMail -like "*$GoogleWorkspaceSourceMailDomain" ) {
        Write-Host "modify existing legacy account to reflect replacement target domain..."
        Write-Host "Invoke-Expression $GamDir\gam.exe update user $LegacyUserMail email $ReplacementUserMail firstname '$FirstName' lastname '$LastName' org '$GoogleWorkspaceDestinationUserOU/$UpdatedDestOU' $GoogleCustomAttribute01 $UPN gal false"
        $password = "N/A - unchanged"

        } else {

            Write-Host "Generating Random Password..." 
                $pwd = $(Invoke-WebRequest -Uri $PwdWebRequestURI -UseBasicParsing)
                #    $pwd.Content
                    #$pwd.StatusCode
                        if ($pwd.StatusCode -eq 200) {
                        Write-Host "proceed with pwd reservation"
                        $password = $($pwd.Content)
                        #Write-Host "Password: " $password
                        } else {
                        Write-Error "No Webserver, or pwd received"
                        $password = $PwdFailsafe
                        }

                start-sleep 1

            Write-Warning "Creating desired target mail domain email address..."
            Write-Host "Invoke-Expression $GamDir\gam.exe create user $ReplacementUserMail firstname '$FirstName' lastname '$LastName' org '$GoogleWorkspaceDestinationUserOU/$UpdatedDestOU' $GoogleCustomAttribute01 $UPN password $password gal false"

    }
    
    #capture initial credentials
    "$firstname,$lastname,$legacyUserMail,$ReplacementUserMail,$password" | out-file -filepath $tempcsv2 -Append 
    
    #Invoke-Expression "$GamDir\gam.exe create user $ReplacementUserMail firstname $FirstName lastname $LastName password $password org '$GoogleWorkspaceDestinationUserOU' changepassword on" ### ## UPDATE NEEDED ##
    #write-Host "Invoke-Expression $GamDir\gam.exe update user $ReplacementUserMail firstname $FirstName lastname $LastName password $password org '$GoogleWorkspaceDestinationUserOU/$UpdatedDestOU' " ### ## UPDATE NEEDED ##
    #write-Host "Invoke-Expression $GamDir\gam.exe update user $ReplacementUserMail firstname $FirstName lastname $LastName org '$GoogleWorkspaceDestinationUserOU/$UpdatedDestOU' " ### ## UPDATE NEEDED ##


    <#
    #start-sleep 3 # 
    Write-Host "update Replacement account UPN..."
    #Invoke-Expression "$GamDir\gam.exe update user $ReplacementUserMail $GoogleCustomAttribute01 $UPN" #set UPN - 
    Write-Host "Invoke-Expression $GamDir\gam.exe update user $ReplacementUserMail $GoogleCustomAttribute01 $UPN" #set UPN - 

    

    Write-Host "hide account from GAL.."
    Write-Host "Invoke-Expression $GamDir\gam.exe update user $ReplacementUserMail gal false" ## UPDATE NEEDED##
    #>
    DashedLine
    #>
}

#upload initial credentials to gsheet source $tempcsv2
#Write-Host "replacing content of existing google sheet with upto date data..."
#Invoke-Expression "$GamDir\gam.exe user $GoogleSvcAccount update drivefile id $UserInfoGsheetID localfile $tempcsv2 newfilename '$GoogleWorkspaceDestinationMailDomain User Information' " ##UPDATE NEEDED## student folder/filename

<#
#Set Google instance: legacy...
Write-Host "###### set google instance: legacy... ######"
$GoogleSourceSvcAccount = ("$GoogleServiceAccountPrefix" + "$GoogleWorkSpaceSource" + "@" + "$GGoogleWorkspaceSourceMailDomain")
Write-Host "Google Source Service Account: $GoogleSourceSvcAccount"
Write-Host "Setting workspace source: $GoogleWorkSpaceSource"
Invoke-Expression "$GamDir\gam.exe select $GoogleWorkSpaceSource save" # swap/set google workspace
Invoke-Expression "$GamDir\gam.exe"
DashedLine

foreach ($user in $VerifiedUserData) {
    DashedLine
    $LegacyUserMail = $user."Existing Email Address" #current mail address
    $UPN = $user."Staff full name" # UPN ##UPDATE NEEDED##
    $FirstName = $user."Staff first name" #prefered firstname ##UPDATE NEEDED##
    $LastName = $user."Staff Surname" ##UPDATE NEEDED##
    $ReplacementUserMail = $user."new email"

        
    Write-Host "Processing: $LegacyUserMail"
    Write-Host "UPN: $UPN"
    Write-Host "Firstname: $FirstName"
    Write-Host "Lastname: $LastName"

    Write-Host "update legacy accounts..."
    Invoke-Expression "$GamDir\gam.exe update user $LegacyUserMail $GoogleCustomAttribute01 $UPN" #set LDAP adminNumber ##UPDATE NEEDED##

        DashedLine
}
#>

<#
#Set Google Instance: Destination...
Write-Host "###### Set Google instance: Destination... ######"
$GoogleSvcAccount = $GoogleWorkspaceMNSPsvcAccount
Write-Host "Google Destination Service Account: $GoogleSvcAccount"
Write-Host "Setting workspace Destination: $GoogleWorkSpaceDestination"
Invoke-Expression "$GamDir\gam.exe select $GoogleWorkSpaceDestination save" # swap/set google workspace
Invoke-Expression "$GamDir\gam.exe"
start-sleep 3
DashedLine

foreach ($user in $VerifiedUserData) {
    DashedLine
    $LegacyUserMail = $user."Existing Email Address" #current mail address
    $UPN = $user."Staff full name" # UPN
    $FirstName = $user."Staff first name" #prefered firstname
    $LastName = $user."Staff Surname"
    $ReplacementUserMail = $user."new email"

    #if ( $RunDeckDev -eq "true" ) { $ReplacementUserMail = $RundeckDevMail } #script dev check...
    
    Write-Host "Processing: $LegacyUserMail"
    Write-Host "UPN: $UPN"
    Write-Host "Firstname: $FirstName"
    Write-Host "Lastname: $LastName"

    Write-Host "Accept calendar invite: user $LegacyUserMail add calendar $ReplacementUserMail selected true ..."
    Invoke-Expression "$GamDir\gam.exe user $ReplacementUserMail add calendar $LegacyUserMail selected true"

    
    start-sleep 3 # #TODO - update HRID
    Write-Host "update Replacement account UPN..."
    Invoke-Expression "$GamDir\gam.exe update user $ReplacementUserMail $GoogleCustomAttribute01 $UPN" #set UPN - 
}
#>

<#
    Write-Host "Add members to security groups ..."
        if (test-path $DataDir\*.lst) { remove-item $DataDir\*.lst -force -verbose } #force delete any .lst files if exist...

        $GoogleGroupMembership = @()
        $GroupMembershipHeader = @()
        $member = @()

        $GoogleGroupMembership = Import-csv -path $tempcsv6
        $GroupMembershipHeader = $($GoogleGroupMembership | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name)

        foreach ($member in $GroupMembershipHeader) {

        Write-Host "----------- $member ----------"`n
        $GoogleGroupMembership.$member | where { $_ -notlike "#N/A" } | out-file -Encoding utf8 "$DataDir\$member.lst"

        $GoogleGroupFQDN = ($member + "@" + $GoogleWorkspaceDestinationMailDomain).ToLower()
        Invoke-expression "$GamDir\gam.exe update group $GoogleGroupFQDN add members file $DataDir\$member.lst"

    }

Write-Host "Add members to mail dist groups ..."
        if (test-path $DataDir\*.lst) { remove-item $DataDir\*.lst -force -verbose } #force delete any .lst files if exist...

        $GoogleGroupMembership = @()
        $GroupMembershipHeader = @()
        $member = @()

        $GoogleGroupMembership = Import-csv -path $tempcsv7
        $GroupMembershipHeader = $($GoogleGroupMembership | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name)

        foreach ($member in $GroupMembershipHeader) {

        Write-Host "----------- $member ----------"`n
        $GoogleGroupMembership.$member | where { $_ -notlike "#N/A" } | out-file -Encoding utf8 "$DataDir\$member.lst"

        $GoogleGroupFQDN = ($member + "@" + $GoogleWorkspaceDestinationMailDomain).toLower()
        Invoke-expression "$GamDir\gam.exe update group $GoogleGroupFQDN add members file $DataDir\$member.lst"

    }
#>

<#


#set google instance: legacy
Write-Host "###### set google instance: legacy... ######"
$GoogleSourceSvcAccount = ("$GoogleServiceAccountPrefix" + "$GoogleWorkSpaceSource" + "@" + "$GGoogleWorkspaceSourceMailDomain") # set service account to use to download gsheets
Write-Host "Google Source Service Account: $GoogleSourceSvcAccount"
Write-Host "Setting workspace source: $GoogleWorkSpaceSource"
Invoke-Expression "$GamDir\gam.exe select $GoogleWorkSpaceSource save" # swap/set google workspace
Invoke-Expression "$GamDir\gam.exe" #get current google workspace


#Set local sysadmins group mail address... # any members of this group can see content of all local shared drives
#$GoogleWorkspaceSourceSysadminGroupFQDN = ("$GoogleWorkspaceSourceSysadminGroup" + "@" + "$GoogleWorkspaceSourceMailDomain")

#create $GoogleWorkspaceSourceSysadminGroupFQDN security group...
#Write-Host "Creating local sysadmins security group: $GoogleWorkspaceSourceSysadminGroupFQDN"
#Invoke-expression "$GamDir\gam.exe create group $GoogleWorkspaceSourceSysadminGroupFQDN" # create group
#Start-Sleep 2
#Invoke-Expression "$GamDir\gam.exe update cigroup $GoogleWorkspaceSourceSysadminGroupFQDN makesecuritygroup" # set group label/type to security

#create source users calendar info gsheet
#Write-Host "Source users current calendar info..."
#Invoke-Expression "$GamDir\gam.exe ou_and_children_ns ""$GoogleWorkspaceSourceUserOU"" print calendars showhidden todrive tdparent id:$GfolderReportsID tdnobrowser"

<#
Write-Host "downloading gsheet ID: $GoogleSheetID tab: $GoogleSheetTab06"
Invoke-Expression "$GamDir\gam.exe user $GoogleSvcAccount get drivefile $GoogleSheetID format csv gsheet ""$GoogleSheetTab06"" targetfolder $DataDir targetname $tempcsv6"

$GoogleGroups = @()
$GoogleGroupsHeader = @()
$member = @()
$GroupexistCheck =@()

$GoogleGroups = Import-csv -path $tempcsv6
$GoogleGroupsHeader = $($GoogleGroups | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name) # get all column headings
Write-Host "CSV header: $GoogleGroupsHeader"
$GroupNameSearchString = $($GoogleGroupsHeader[0].substring(0,11)) #first element of array, first 11 chars - TODO - better logic - also do not include names containing "security Group #"
Write-Host "Group Search String: $GroupNameSearchString"

#
Invoke-Expression "$GamDir\gam.exe print groups query ""email:$GroupNameSearchString*"" > $tempcsv8" #check if group already exists...
$GroupexistCheck = import-csv -Path $tempcsv8 #check if group already exists...
Write-Host "Existing group(s):"
$GroupexistCheck.email

    foreach ($member in $GoogleGroupsHeader) {
    
    ##### TODO - ########### at least one group matching criteria MUST exist ###### or ELSE loop fails - rewite required #######
     if ($GroupexistCheck.email.Contains($member)) { #TODO logic not working group creation still being started, although GAM process does spot duplicate and skips actual creation...
        Write-Warning "Group: $member already exists skiping creation ..."
        } else { 

    Write-Host "-----------Creating Security group:$member----------"`n
    $GoogleGroupFQDN = ($member + "@" + $GoogleWorkspaceDestinationMailDomain).ToLower()
    Invoke-expression "$GamDir\gam.exe create group $GoogleGroupFQDN" # create group
    Start-Sleep 2
    Invoke-Expression "$GamDir\gam.exe update cigroup $GoogleGroupFQDN makesecuritygroup" # set group label/type to security

    Start-sleep 2
        #set access controls for group from action array... ENHANCEMENT - migrate this to a single JSON control file route
        foreach ($action in $GoogleWorkspaceSecGroupSettings) { 
        Invoke-expression "$GamDir\gam.exe update group $GoogleGroupFQDN $action" #set access controls for group from action array
        
        }

    }
    }
#>

<#
Write-Host "Create email dist groups (Destination instance)..."
$GoogleWorkspaceGroupSettings = ("isArchived true","whoCanContactOwner ALL_MEMBERS_CAN_CONTACT","whoCanMarkFavoriteReplyOnOwnTopic OWNERS_AND_MANAGERS","whoCanPostMessage ALL_MEMBERS_CAN_POST","whoCanTakeTopics OWNERS_AND_MANAGERS","whoCanViewGroup ALL_MEMBERS_CAN_VIEW","whoCanViewMembership ALL_MEMBERS_CAN_VIEW","whoCanJoin INVITED_CAN_JOIN") #ENHANCEMENT convert to json updating 

if (test-path $tempcsv7) { remove-item $tempcsv7 -force -verbose }
if (test-path $tempcsv8) { remove-item $tempcsv8 -force -verbose }
start-sleep 2

Write-Host "downloading gsheet ID: $GoogleSheetID tab: $GoogleSheetTab07"
Invoke-Expression "$GamDir\gam.exe user $GoogleSvcAccount get drivefile $GoogleSheetID format csv gsheet ""$GoogleSheetTab07"" targetfolder $DataDir targetname $tempcsv7"

$GoogleGroups = @()
$GoogleGroupsHeader = @()
$member = @()
$GroupexistCheck = @()

$GoogleGroups = Import-csv -path $tempcsv7
$GoogleGroupsHeader = $($GoogleGroups | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name)
Write-Host "CSV header: $GoogleGroupsHeader"
$GroupNameSearchString = $($GoogleGroupsHeader[0].substring(0,3)) #first element of array, first 3 chars
Write-Host "Group Search String: $GroupNameSearchString"

Invoke-Expression "$GamDir\gam.exe print groups query ""email:$GroupNameSearchString*"" > $tempcsv8" #check if group already exists...
$GroupexistCheck = import-csv -Path $tempcsv8 #check if group already exists...
Write-Host "Existing group(s):"
$GroupexistCheck.email

    foreach ($member in $GoogleGroupsHeader) {

             if ($GroupexistCheck.email.Contains($member)) {
        Write-Warning "Group: $member already exists skiping creation ..."
        } else { 

    Write-Host "-----------Creating Dist group:$member----------"`n
    $GoogleGroupFQDN = ($member + "@" + $GoogleWorkspaceDestinationMailDomain).ToLower()
    Invoke-expression "$GamDir\gam.exe create group $GoogleGroupFQDN"

    Start-sleep 2

        foreach ($action in $GoogleWorkspaceGroupSettings) { 
        Invoke-expression "$GamDir\gam.exe update group $GoogleGroupFQDN $action"
        
        }

    }
    }
#>



#Write-Host "generate MFA backup codes..." # Agreed  not to enforce imediate MFA - grace period of 2 days instead
    #Invoke-Expression "$GamDir\gam.exe user $ReplacementUserMail update backupcodes"

    #start-sleep 3 # - ENHANCEMENT (confirm) updating of attribute is NOT consistent, may need a few seconds delay after account is created beforeready to accept custom attribute setting: Update Failed: Invalid Schema Value 
    #Write-Host "update Replacement account..."
    #Invoke-Expression "$GamDir\gam.exe update user $ReplacementUserMail $GoogleCustomAttribute01 $UPN" #set UPN - 
    
    #ENHANCEMENT - update Job Title and Department from exported peopleXD data (Job Title Description and Division Description fields)

<#
Write-Host "Create common shared drives security groups (Destination instance)..."
$GoogleWorkspaceSecGroupSettings = ("whoCanContactOwner ALL_MANAGERS_CAN_CONTACT","isArchived true","whoCanContactOwner ALL_MANAGERS_CAN_CONTACT","whoCanMarkFavoriteReplyOnOwnTopic OWNERS_AND_MANAGERS","whoCanPostMessage ALL_MANAGERS_CAN_POST","whoCanTakeTopics OWNERS_AND_MANAGERS","whoCanViewGroup ALL_MANAGERS_CAN_VIEW","whoCanViewMembership ALL_MANAGERS_CAN_VIEW","whoCanJoin INVITED_CAN_JOIN") #ENHANCEMENT - convert to json updating 
#>

<#
Write-Host "send current calendar invite: $LegacyUserMail add acls reader $ReplacementUserMail ..."
    Invoke-Expression "$GamDir\gam.exe calendar $LegacyUserMail add acls reader $ReplacementUserMail sendnotifications false"

    Write-Host "shared drive creation (Legacy Source to Destination user)..."
    $TeamDriveName = "Migration $LegacyUserMail $(Get-Date)"
    $LegacyUserTeamDriveID = $(Invoke-expression "$GamDir\gam.exe user $GoogleSourceSvcAccount create teamdrive '$TeamDriveName' adminmanagedrestrictions true asadmin returnidonly")
    Write-Host "Shared Drive ID: $LegacyUserTeamDriveID "

    Write-Host "Allow outside sharing..."
    Invoke-expression "$GamDir\gam.exe update teamdrive $LegacyUserTeamDriveID asadmin domainUsersOnly False"

    Write-Host "Allow people who aren't shared drive members to access files - false..."
    Invoke-expression "$GamDir\gam.exe update teamdrive $LegacyUserTeamDriveID asadmin sharingFoldersRequiresOrganizerPermission True"

    #Write-Host "move shared drive to move enabled OU..."
    #Invoke-expression "$GamDir\gam.exe update teamdrive $LegacyUserTeamDriveID asadmin ou '$LegacyUserTeamDriveOU'" #location needs confirming

    Write-Host "Add internal sysadmins group as manager: add drivefileacl $LegacyUserTeamDriveID user $GoogleWorkspaceSourceSysadminGroupFQDN role organizer"
    Invoke-expression "$GamDir\gam.exe add drivefileacl $LegacyUserTeamDriveID user $GoogleWorkspaceSourceSysadminGroupFQDN role organizer" 

    Write-Host "Add internal user as manager: add drivefileacl $LegacyUserTeamDriveID user $LegacyUserMail role organizer ..."
    Invoke-expression "$GamDir\gam.exe add drivefileacl $LegacyUserTeamDriveID user $LegacyUserMail role organizer"
    
    Write-Host "Add external user as manager: add drivefileacl $LegacyUserTeamDriveID user $ReplacementUserMail role organizer..."
    Invoke-expression "$GamDir\gam.exe add drivefileacl $LegacyUserTeamDriveID user $ReplacementUserMail role organizer"
    
    Write-Host "report current shared drive folder associations for: $LegacyUserMail ..."
    Invoke-expression "$GamDir\gam.exe user $legacyUserMail print teamdrives todrive tdparent id:$GfolderReportsID tdnobrowser tdtitle '$LegacyUserMail shared drives summary as of $(get-date)'"


#>