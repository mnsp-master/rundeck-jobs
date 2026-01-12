$mnspver = "0.0.177.17"

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

Start-sleep 2

if (test-path $tempcsv6) { remove-item $tempcsv6 -force -verbose }
if (test-path $tempcsv7) { remove-item $tempcsv7 -force -verbose }
if (test-path $tempcsv8) { remove-item $tempcsv8 -force -verbose }

start-sleep 2

Write-Host "downloading gsheet ID: $GoogleSheetID tab: $GoogleSheetTab06"
Invoke-Expression "$GamDir\gam.exe user $GoogleSvcAccount get drivefile $GoogleSheetID format csv gsheet ""$GoogleSheetTab06"" targetfolder $DataDir targetname $tempcsv6"

Write-Host "downloading gsheet ID: $GoogleSheetID tab: $GoogleSheetTab07"
Invoke-Expression "$GamDir\gam.exe user $GoogleSvcAccount get drivefile $GoogleSheetID format csv gsheet ""$GoogleSheetTab07"" targetfolder $DataDir targetname $tempcsv7"

#create groups...
$GoogleWorkspaceGroupSettings = ("isArchived true","whoCanContactOwner ALL_MEMBERS_CAN_CONTACT","whoCanMarkFavoriteReplyOnOwnTopic OWNERS_AND_MANAGERS","whoCanPostMessage ALL_MEMBERS_CAN_POST","whoCanTakeTopics OWNERS_AND_MANAGERS","whoCanViewGroup ALL_MEMBERS_CAN_VIEW","whoCanViewMembership ALL_MEMBERS_CAN_VIEW","whoCanJoin INVITED_CAN_JOIN") #ENHANCEMENT convert to json updating 
$GoogleWorkspaceSecGroupSettings = ("isArchived true","whoCanContactOwner ALL_MEMBERS_CAN_CONTACT","whoCanMarkFavoriteReplyOnOwnTopic OWNERS_AND_MANAGERS","whoCanPostMessage ALL_MEMBERS_CAN_POST","whoCanTakeTopics OWNERS_AND_MANAGERS","whoCanViewGroup ALL_MEMBERS_CAN_VIEW","whoCanViewMembership ALL_MEMBERS_CAN_VIEW","whoCanJoin INVITED_CAN_JOIN") #ENHANCEMENT convert to json updating 

#<#
Write-Host "Create email security groups (Destination instance)..."

if (test-path $tempcsv8) { remove-item $tempcsv8 -force -verbose }
start-sleep 2

$GoogleGroups = @()
$GoogleGroupsHeader = @()
$member = @()
$GroupexistCheck =@()

$GoogleGroups = Import-csv -path $tempcsv6 ####ENHANCEMENT#### try/catch - dupliacte header columns - exit if duplicates found
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
    Invoke-expression "$GamDir\gam.exe create group $GoogleGroupFQDN" #CID00#### dry run # create group 
    Write-Host "$GamDir\gam.exe create group $GoogleGroupFQDN"
    Start-Sleep 2
    
    #no longer setting groups as security groups Oct 2025
    #Invoke-Expression "$GamDir\gam.exe update cigroup $GoogleGroupFQDN makesecuritygroup" #CID00#### dry run # set group label/type to security
    #Write-Host "$GamDir\gam.exe update cigroup $GoogleGroupFQDN makesecuritygroup"

    Start-sleep 2
        #set access controls for group from action array... ENHANCEMENT - migrate this to a single JSON control file route
        foreach ($action in $GoogleWorkspaceSecGroupSettings) { 
        Invoke-expression "$GamDir\gam.exe update group $GoogleGroupFQDN $action" #CID00#### dry run #set access controls for group from action array
        Write-Host "$GamDir\gam.exe update group $GoogleGroupFQDN $action"

        }

    }
    }

Write-Host "Create email dist groups (Destination instance)..."

#if (test-path $tempcsv7) { remove-item $tempcsv7 -force -verbose }
if (test-path $tempcsv8) { remove-item $tempcsv8 -force -verbose }
start-sleep 2

#Write-Host "downloading gsheet ID: $GoogleSheetID tab: $GoogleSheetTab07"
#Invoke-Expression "$GamDir\gam.exe user $GoogleSvcAccount get drivefile $GoogleSheetID format csv gsheet ""$GoogleSheetTab07"" targetfolder $DataDir targetname $tempcsv7"

$GoogleGroups = @()
$GoogleGroupsHeader = @()
$member = @()
$GroupexistCheck = @()

$GoogleGroups = Import-csv -path $tempcsv7 ####ENHANCEMENT#### try/catch - dupliacte header columns - exit if duplicates found
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
    Invoke-expression "$GamDir\gam.exe create group $GoogleGroupFQDN" #CID00#### dry run
    Write-Host "$GamDir\gam.exe create group $GoogleGroupFQDN"

    Start-sleep 2

        foreach ($action in $GoogleWorkspaceGroupSettings) { 
        Invoke-expression "$GamDir\gam.exe update group $GoogleGroupFQDN $action" #CID00#### dry run
        Write-Host "$GamDir\gam.exe update group $GoogleGroupFQDN $action"
        
        }

    }
    }


    Write-Host "Add members to security groups ..."
        #if (test-path $DataDir\*.lst) { remove-item $DataDir\*.lst -force -verbose } #force delete any .lst files if exist...

        $GoogleGroupMembership = @()
        $GroupMembershipHeader = @()
        $member = @()

        $GoogleGroupMembership = Import-csv -path $tempcsv6
        $GroupMembershipHeader = $($GoogleGroupMembership | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name)

        foreach ($member in $GroupMembershipHeader) {

        Write-Host "----------- $member ----------"`n
        $GoogleGroupMembership.$member | where { $_ -notlike "#N/A" } | out-file -Encoding utf8 "$DataDir\$member.lst"

        $GoogleGroupFQDN = ($member + "@" + $GoogleWorkspaceDestinationMailDomain).ToLower()
        Write-host "Invoke-expression $GamDir\gam.exe update group $GoogleGroupFQDN add members file $DataDir\$member.lst"
        Invoke-expression "$GamDir\gam.exe update group $GoogleGroupFQDN add members file $DataDir\$member.lst" #adds members setting sync will sync add/remove #CID00#### dry runv
        

    }
#>

Write-Host "Add members to mail dist groups ..."
        #if (test-path $DataDir\*.lst) { remove-item $DataDir\*.lst -force -verbose } #force delete any .lst files if exist...

        $GoogleGroupMembership = @()
        $GroupMembershipHeader = @()
        $member = @()

        $GoogleGroupMembership = Import-csv -path $tempcsv7
        $GroupMembershipHeader = $($GoogleGroupMembership | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name)

        foreach ($member in $GroupMembershipHeader) {

        Write-Host "----------- $member ----------"`n
        $GoogleGroupMembership.$member | where { $_ -notlike "#N/A" } | out-file -Encoding utf8 "$DataDir\$member.lst"

        $GoogleGroupFQDN = ($member + "@" + $GoogleWorkspaceDestinationMailDomain).toLower()
        Write-Host "Invoke-expression $GamDir\gam.exe update group $GoogleGroupFQDN add members file $DataDir\$member.lst"
        Invoke-expression "$GamDir\gam.exe update group $GoogleGroupFQDN add members file $DataDir\$member.lst" #adds members setting sync will sync add/remove #CID00#### dry run
        

    }


<#
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
#Invoke-Expression "$GamDir\gam.exe print groups query ""email:$GroupNameSearchString*"" > $tempcsv8" #check if group already exists...
#$GroupexistCheck = import-csv -Path $tempcsv8 #check if group already exists...
#Write-Host "Existing group(s):"
#$GroupexistCheck.email


$GoogleGroups = @()
$GoogleGroupsHeader = @()
$member = @()
$GroupexistCheck = @()

$GoogleGroups = Import-csv -path $tempcsv7
$GoogleGroupsHeader = $($GoogleGroups | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name)
Write-Host "CSV header: $GoogleGroupsHeader"
$GroupNameSearchString = $($GoogleGroupsHeader[0].substring(0,3)) #first element of array, first 3 chars
Write-Host "Group Search String: $GroupNameSearchString"

#>