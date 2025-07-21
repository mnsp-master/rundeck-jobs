$mnspver = "0.0.95"

Write-Host $(Get-Date)
Write-Host "MNSP Version" $mnspver
Start-Sleep 10
$ErrorActionPreference="Continue"
Set-Location $GamDir

#$FormatEnumerationLimit=-1
#$FormatEnumerationLimit

function DashedLine {
Write-host "-----------------------------------------------------------`n"
}

#Get-Variable

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

#OUS to create:
if (test-path $tempcsv10) { remove-item $tempcsv10 -force -verbose }
Write-Host "Getting OU's from parent: $GoogleworkspaceDestinationUserOU"
$CurrentOUsCSV =@()
$CurrentOUsCSV = $(Invoke-expression "$GamDir\gam.exe print orgs fromparent '$GoogleworkspaceDestinationUserOU' | Out-file $tempcsv10" )

$CurrentOUs =@()
DashedLine
$OUplaceHolder | out-file $tempcsv10 -Append
$CurrentOUs = Import-Csv -Path $tempcsv10


foreach ($OUtoCreate in $OUsToCreate) {
    if ($CurrentOUs.name.contains($OUtoCreate)) { #logic does not work if NO sub OU's currently exist, hence $OUplaceHolder fix... ####ENHANCEMENT#### fix the logic properly
    Write-host "OU: $OUtoCreate already exists"
    DashedLine
    } else {
    Write-Warning "OU: $OUtoCreate does not exist, creating..."
    Write-Host "$GamDir\gam.exe create org '$OutoCreate' parent '$GoogleWorkspaceDestinationUserOU'"

    #invoke-expression "$GamDir\gam.exe create org '$OutoCreate' parent '$GoogleWorkspaceDestinationUserOU'" #CID00#### dry run

    DashedLine
    }
}

#exit ###SNO DEBUG###

if (test-path $tempcsv9) { remove-item $tempcsv9 -force -verbose }
Write-Host "Report on all current users from base OU: $GoogleWorkspaceSourceUserOU"
Invoke-expression "$GamDir\gam.exe ou_and_children '$GoogleWorkspaceSourceUserOU' print allfields >> $tempcsv9" 
Invoke-expression "$GamDir\gam.exe ou_and_children '$GoogleWorkspaceSourceUserOU' print allfields todrive tdparent id:$GfolderReportsID tdtitle 'User info - Pre Migration for domain: $GoogleWorkspaceSourceMailDomain as of: $(Get-date)'"

$GoogleWorkspaceSourceUsers = import-csv -path $tempcsv9

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

<#
$VerifiedUserData = Get-Content -path $tempcsv4 | convertFrom-csv | where-object { 
    $_.$FieldMatch01 -like $FieldString -and 
    $_.$Fieldmatch02 -match '^[0-9]+$' #Numeric values only - excludes - R N1 N2 etc ####ENHANCEMENT #### include only specific year groups
    #$_.$Fieldmatch02 -like "12" #limit to year group(s)
    } #import where field like $FieldMatch01
#>

###CID003776#### Specific year groups only
$VerifiedUserData = Get-Content -path $tempcsv4 | convertFrom-csv | where-object { 
    ($_.$FieldMatch01 -like $FieldString) -and 
        (
            ($_.$Fieldmatch02 -like "9") -or 
            ($_.$Fieldmatch02 -like "10") -or
            ($_.$Fieldmatch02 -like "12")
        )
    }

    $VerifiedUserData.count
    $VerifiedUserData.School_Name | Get-Unique
    $VerifiedUserData.$Fieldmatch02 | Get-Unique
###CID003776####

#$VerifiedUserData = Get-Content -path $tempcsv4 | select-object -skip 1 | convertFrom-csv | where { $_.$FieldMatch01 -like $FieldString } #import where field like $FieldMatch01, and skip 1st line
Write Host "Number of records matching selection criteria:" $VerifiedUserData.count
#TODO - if count 0 break out of script...

#create user info destination gsheet
#$UserInfoGsheetID = $(Invoke-Expression "$GamDir\gam.exe user $GoogleSvcAccount create drivefile drivefilename '$GoogleWorkspaceDestinationMailDomain User Info' mimetype gsheet parentid $GfolderReportsID returnidonly")

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
    #$ReplacementUserMail = $user."new email"
    $ReplacementUserMail = $user."Email20Chars"
    $DestOU = [int] $user."NC Year(s) for today" #set var as interger

    #add leading zero if required: to create consitent OUs YEAR07 not YEAR7: 
        if ( $DestOU -le 9) {
            Write-host "Target Year group less than or equal to 9..."
            $UpdatedDestOU = @()
            $UpdatedDestOU = $($GoogleWorkSpaceDest + "-Year" + "0" + $DestOU)
            } else {
            $UpdatedDestOU = $($GoogleWorkSpaceDest + "-Year" + $DestOU)
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


  Write-Host "Checking if legacy mail: $LegacyUserMail  like: $GoogleWorkspaceSourceMailDomain"
    if ( $LegacyUserMail -like "*$GoogleWorkspaceSourceMailDomain" ) {
        Write-Host "modifying existing legacy account to reflect replacement target domain..."
        
        Write-host "confirm MIS email data email address actually exists..."
        if ($GoogleWorkspaceSourceUsers.primaryEmail.Contains($LegacyUserMail)) {
            Write-Host "$LegacyUserMail exists..."
        } else {
            Write-Warning "!!WARNING $LegacyUserMail DOES NOT EXIST!! - update MIS data"
        }
        
        
        #Invoke-Expression "$GamDir\gam.exe update user $LegacyUserMail email $ReplacementUserMail firstname '$FirstName' lastname '$LastName' org '$GoogleWorkspaceDestinationUserOU/$UpdatedDestOU' $GoogleCustomAttribute01 $UPN gal $GoogleIncludeInGal" #CID00#### dry run
        Write-Host "$GamDir\gam.exe update user $LegacyUserMail email $ReplacementUserMail firstname '$FirstName' lastname '$LastName' org '$GoogleWorkspaceDestinationUserOU/$UpdatedDestOU' $GoogleCustomAttribute01 $UPN gal $GoogleIncludeInGal"
        $password = "N/A - unchanged"
        $AccountHistory = "Migrated"

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
            #Invoke-Expression "$GamDir\gam.exe create user $ReplacementUserMail firstname '$FirstName' lastname '$LastName' org '$GoogleWorkspaceDestinationUserOU/$UpdatedDestOU' $GoogleCustomAttribute01 $UPN password $password gal $GoogleIncludeInGal" ##CID00#### dry run
            Write-Host "$GamDir\gam.exe create user $ReplacementUserMail firstname '$FirstName' lastname '$LastName' org '$GoogleWorkspaceDestinationUserOU/$UpdatedDestOU' $GoogleCustomAttribute01 $UPN password $password gal $GoogleIncludeInGal"
            $LegacyUserMail = "N/A"
            $AccountHistory = "New"

    }
    
    #capture initial credentials
    "$firstname,$lastname,$legacyUserMail,$ReplacementUserMail,$password,$AccountHistory,$UPN" | out-file -filepath $tempcsv2 -Append
    
    DashedLine
    

}
    #upload post migtation data in gsheet...
    $UpdatedUsersInfoGsheetID = $(Invoke-Expression "$GamDir\gam.exe user $GoogleSvcAccount create drivefile drivefilename '$GoogleWorkspaceDestinationMailDomain Migrated User Info' mimetype gsheet parentid $GfolderReportsID returnidonly")
    Invoke-Expression "$GamDir\gam.exe user $GoogleSvcAccount update drivefile id $UpdatedUsersInfoGsheetID localfile $tempcsv2 newfilename 'User info - Post Migration for domain: $GoogleWorkspaceDestinationMailDomain as of: $(Get-date)'" #-ErrorAction SilentlyContinue 


#######################



