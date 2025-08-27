$mnspver = "0.0.42"

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

#prepare user details csv
Write-Host "emptying $tempcsv2 of any existing data..."
Clear-Content $tempcsv2
sleep 1
$UserInfoCSVheader | out-file -filepath $tempcsv2 -Append #create blank csv with simple header

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

#create user info destination gsheet
$UserInfoGsheetID = $(Invoke-Expression "$GamDir\gam.exe user $GoogleSvcAccount create drivefile drivefilename '$GoogleWorkspaceDestinationMailDomain User Info' mimetype gsheet parentid $GfolderReportsID returnidonly")

<# no security groups required: START #

Write-Host "Create/update common shared drives security groups (Destination instance)..."
$GoogleWorkspaceSecGroupSettings = ("whoCanContactOwner ALL_MANAGERS_CAN_CONTACT","isArchived true","whoCanContactOwner ALL_MANAGERS_CAN_CONTACT","whoCanMarkFavoriteReplyOnOwnTopic OWNERS_AND_MANAGERS","whoCanPostMessage ALL_MANAGERS_CAN_POST","whoCanTakeTopics OWNERS_AND_MANAGERS","whoCanViewGroup ALL_MANAGERS_CAN_VIEW","whoCanViewMembership ALL_MANAGERS_CAN_VIEW","whoCanJoin INVITED_CAN_JOIN") #ENHANCEMENT - convert to json updating 

if (test-path $tempcsv6) { remove-item $tempcsv6 -force -verbose }
if (test-path $tempcsv8) { remove-item $tempcsv8 -force -verbose }

start-sleep 2

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

    Write-Host "-----------Creating Security group: $member ----------"`n
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

# no security groups required: END #>

<# no distribution groups required: START #
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

    Write-Host "-----------Creating Dist group: $member ----------"`n
    $GoogleGroupFQDN = ($member + "@" + $GoogleWorkspaceDestinationMailDomain).ToLower()
    Invoke-expression "$GamDir\gam.exe create group $GoogleGroupFQDN"

    Start-sleep 2

        foreach ($action in $GoogleWorkspaceGroupSettings) { 
        Invoke-expression "$GamDir\gam.exe update group $GoogleGroupFQDN $action"
        
        }

    }
    }
#># no distribution groups required: END #

<##>

Write-Host "Creating users in destination..."
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

    Write-Host "create account..."
    #Write-Host "$GamDir\gam.exe create user email $ReplacementUserMail firstname '$FirstName' lastname '$lastname' org '$GoogleWorkspaceDestinationUserOU' "
    Write-Host "$GamDir\gam.exe create user $ReplacementUserMail firstname '$FirstName' lastname '$LastName' org '$GoogleWorkspaceDestinationUserOU' $GoogleCustomAttribute01 $HRid password $password gal $GoogleIncludeInGal"
    #Invoke-Expression "$GamDir\gam.exe update user $LegacyUserMail email $ReplacementUserMail firstname '$FirstName' lastname '$lastname' org '$GoogleWorkspaceDestinationUserOU' " ###move/update existing user #CID00#### dry run

    #capture initial credentials
    "$firstname,$lastname,$legacyUserMail,$ReplacementUserMail,$password,$HRid" | out-file -filepath $tempcsv2 -Append

    #generate MFA backup codes
    Write-host "$GamDir\gam.exe user $ReplacementUserMail update backupcodes"
    #$userBackupCodes = invoke-expression "$GamDir\gam.exe user $ReplacementUserMail update backupcodes"

    #send mail(s)
    ##backup codes...

    ##credentials...

    ##account information...
    Write-Host "$GamDir\gam.exe sendemail $legacyUserMail from $GoogleWorkspaceSenderMail newuser $ReplacementUserMail firstname $FirstName LastName $LastName password $password"

    DashedLine
}

<# no security groups required: START #
    Write-Host "sync members of security groups ..."
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
        #Invoke-expression "$GamDir\gam.exe update group $GoogleGroupFQDN add members file $DataDir\$member.lst"
        Invoke-expression "$GamDir\gam.exe update group $GoogleGroupFQDN sync members file $DataDir\$member.lst"

    }
# no security groups required: END #>

<# no distribution groups required: START #
Write-Host "sync members of mail dist groups ..."
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
        #Invoke-expression "$GamDir\gam.exe update group $GoogleGroupFQDN add members file $DataDir\$member.lst"
        Invoke-expression "$GamDir\gam.exe update group $GoogleGroupFQDN sync members file $DataDir\$member.lst"


    }
    # no distribution groups required: END #



<#
#>