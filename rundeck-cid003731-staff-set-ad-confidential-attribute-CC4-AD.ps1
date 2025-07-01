$mnspver = "0.0.19"

<#
Overall process to:
- Begin transcript logging
- Download google sheet - or use local/manually saved CSV, containing all necessary information (MAT wide export from central arbor), data will be filtered based on desired school
- Download google sheet - or use local/manually saved CSV, containing all necessary information (user migration information containing HR ID (M00######))
- Create a user array (from downloaded/filtered csv data)
    - Loop through the user array, finding the AD object that matches Arbor ID (employeeNumber AD attribute previously synchronised by salamander)
    - Return matched AD user and attributes
    - match legacy email with returned AD search result to cross refernce with HR ID source data - return HR ID for email address
    - Set custom confidential attribute (mnspAdminNumber)
    - Repeat for as many users that are in the array
- Stop transcript logging
- Exit PS script
#>

##### ENHANCEMENT ##### general environment/vars area required to enable PS execution outside of rundeck environment

Write-Host $(Get-Date)
Write-Host "MNSP Version" $mnspver
Start-Sleep 10
$ErrorActionPreference="Continue"
Set-Location $GamDir

function DashedLine01 {
Write-host "-----------------------------------------------------------`n"
}

function DashedLine02 {
Write-host "---------`n"
}

#Get-Variable

#Set Google Instance: Destination...
Write-Host "###### Set Google instance: Destination... ######"

$GoogleSvcAccount = $GoogleWorkspaceMNSPsvcAccount
Write-Host "Google Destination Service Account: $GoogleSvcAccount"

Write-Host "Setting workspace Destination: $GoogleWorkSpaceDestination"
Invoke-Expression "$GamDir\gam.exe select $GoogleWorkSpaceDestination save" # swap/set google workspace
Invoke-Expression "$GamDir\gam.exe info domain"
start-sleep 3

Write-Host "gsheet Student number column heading:" $FieldMatch01
Start-sleep 10

<#
#prepare user details csv
Write-Host "emptying $tempcsv2 of any existing data..."
Clear-Content $tempcsv2
sleep 1
$UserInfoCSVheader | out-file -filepath $tempcsv2 -Append #create blank csv with simple header

##### ENHANCEMENT ##### toggle between local previously manually downloaded gsheet/csv, and downloading gsheet each time
#>

DashedLine01

#get verified user data
#if exist check & remove $tempcsv4
if (test-path $tempcsv4) { remove-item $tempcsv4 -force -verbose }

##### ENHANCEMENT ##### ERROR TO ADDRESS - Duplicate fileds Value with production gsheet #####
Write-Host "downloading gsheet ID: $GoogleSheetID tab: $GoogleSheetTab01"
Write-Host "Invoke-Expression $GamDir\gam.exe user $GoogleSvcAccount get drivefile $GoogleSheetID format csv gsheet ""$GoogleSheetTab01"" targetfolder $DataDir targetname $tempcsv4"
Invoke-Expression "$GamDir\gam.exe user $GoogleSvcAccount get drivefile $GoogleSheetID format csv gsheet ""$GoogleSheetTab01"" targetfolder $DataDir targetname $tempcsv4"
#Write-Host "Invoke-Expression $GamDir\gam.exe user $GoogleSourceSvcAccount get drivefile $GoogleSheetID format csv gsheet ""$GoogleSheetTab01"" targetfolder $DataDir targetname $tempcsv4"

#get verified user data (HR + mail)
#if exist check & remove $tempcsv6
if (test-path $tempcsv6) { remove-item $tempcsv6 -force -verbose }
Write-Host "downloading gsheet ID: $GoogleSheetID02 tab: $GoogleSheetTab02"
Write-Host "Invoke-Expression $GamDir\gam.exe user $GoogleSvcAccount get drivefile $GoogleSheetID02 format csv gsheet ""$GoogleSheetTab02"" targetfolder $DataDir targetname $tempcsv6"
Invoke-Expression "$GamDir\gam.exe user $GoogleSvcAccount get drivefile $GoogleSheetID02 format csv gsheet ""$GoogleSheetTab02"" targetfolder $DataDir targetname $tempcsv6"


Start-sleep 2

Write-Host "Field match:  " $FieldMatch01
Write-Host "Field String: " $FieldString
#$VerifiedUserData = Get-Content -path $tempcsv4 | where { $_.$FieldMatch01 -like $FieldString }
$VerifiedUserData = Get-Content -path $tempcsv4 | convertFrom-csv | where { $_.$FieldMatch01 -like $FieldString } #import where field like $FieldMatch01


Write-Host "Field match 02:" $Fieldmatch02
Write-Host "Field String 02:" $FiledString02
$VerifiedUserData2 = Get-Content -path $tempcsv6 | select-object -skip 1 | convertFrom-csv | where { $_.$FieldMatch02 -like $FieldString02 } #import where field like $FieldMatch01, and skip 1st line
Write-Host "Number of records matching selection criteria:" $VerifiedUserData2.count
#TODO - if count 0 break out of script...
$VerifiedUserData2


#$VerifiedUserData = Get-Content -path $tempcsv4 | convertFrom-csv | where-object { 
#    $_.$FieldMatch01 -like $FieldString -and 
#    $_.$Fieldmatch02 -match '^[0-9]+$' #Numeric values only - excludes - R N1 N2 etc
#    } #import where field like $FieldMatch01


#$VerifiedUserData = Get-Content -path $tempcsv4 | select-object -skip 1 | convertFrom-csv | where { $_.$FieldMatch01 -like $FieldString } #import where field like $FieldMatch01, and skip 1st line
Write-Host "Number of records matching selection criteria:" $VerifiedUserData.count
##### ENHANCEMENT ##### - if count 0 break out of script...



if (test-path $tempcsv8) { remove-item $tempcsv8 -force -verbose }

start-sleep 2

#OU information:
Write-Host "AD search base DN: $OUBaseDn"

<#
$OUS = $(Get-AdOrganizationalUnit -searchbase $OUBaseDn -Filter *) # get all OU's from baseDN
#>

Write-Host "Updating users..."
foreach ($user in $VerifiedUserData) {
    DashedLine01
    Write-Host "PROCESSING next user..."
    $LegacyUserMail = $user."Email Address (Main)" #current mail address
    #$UPN = $user."UPN" # student MIS UPN
    $FirstName = $user."First Name" #firstname 
    $LastName = $user."Last Name" #lastname
    #$ReplacementUserMail = $user."new email"
    #$ReplacementUserMail = $user."Email20Chars" #UPDATE NEEDED ### Column heading needs agreeing
    #$DestOU = [int] $user."NC Year(s) for today" #set var as interger
    $MISid = $user."Arbor Staff ID" # DEV 
    $MISidComplete = $MISsitePrefix + '-' + $MISrolePrefix + '_' + $MISid #concatenate sitename hyphen and MIS id number e.g: SCH-Stf_34
    #$Yearprefix = $user."New prefix"
    #$MISid = $user."Arbor Student ID" #Production

    
    #user info/attributes:
    Write-Host "HR stored email: $LegacyUserMail"
    #Write-Host "Replacement email: $ReplacementUserMail"
    #Write-Host "UPN: $UPN"
    Write-Host "Firstname: $FirstName"
    Write-Host "Lastname: $LastName"
    #Write-Host "Source Year: $DestOU" 
    #Write-Host "Destination OU name: $UpdatedDestOU"
    Write-Host "MIS ID: $MISid"
    Write-Host "MIS Role:" $MISrolePrefix
    Write-Host "Complete MIS ID: $MISidComplete"
    DashedLine02

    if ($MISid) { #check value is not null...
        $UserToProcess = @()
        $UserToProcess = $(Get-ADUser -Filter "EmployeeNumber -like '$MISidComplete'" -Properties * | select-object $ADattribs) ####ENHACEMENT#### control group needed if user member skip
            if ($UserToProcess.count -gt 1) {
                Write-Warning "Not an singular match..."
                $UserToProcess
                DashedLine02
                    } else {
                        Write-Host "AD attributes found by searching for user with MIS ID: $MISidComplete"
                        $UserToProcess

                        #user lookup using legacy mail address...
                            $UserLookup = @()
                            $UserLookup = $($VerifiedUserData2 | where-object ({ $_.$CSVheaderObject -eq $LegacyUserMail })) ####ENHANCEMENT#### deifne 
                            $HRid = $UserLookup.'Staff full name'
                        
                                        DashedLine02                             
                                        Write-Host "updating AD user: "
                                                                                
                                        # update mnspAdminNumber attribute...
                                        Write-Host "PS to process: Set-ADUser -Identity $($UserToProcess.ObjectGUID) -Add @{mnspAdminNumber="$HRid"} -verbose`n"
                                        Write-host "`n---`n"
                                        Set-ADUser -Identity $($UserToProcess.ObjectGUID) -Add @{mnspAdminNumber="$MISidComplete"} -verbose -whatif ## Comment Whatif to Action
                                        Write-host "`n---`n"
                                        
                                        Write-Host "Updated AD users attributes using GUID:"
                                        $UserToProcessPostupdate = $(Get-ADUser -id $($UserToProcess.ObjectGUID) -Properties * | select-object $ADattribs)
                                        $UserToProcessPostupdate
                                        DashedLine01
                        }
                        DashedLine02

    } else { #if MIS id is NULL...
                        Write-Warning "No MIS ID found for:"
                        Write-host "Legacy email: $LegacyUserMail"
                        Write-Host "Firstname: $FirstName"
                        Write-Host "Lastname: $LastName"
}

}

<#
Write-Host "Closing all remote PSSessions..."
Get-PSSession | Remove-PSSession
#>

<#

Write-Host "creating remote PSSEssions for all Fileservers: $FileServers"
    foreach ($fileServer in $fileServers) {
    Write-host "Creating PSSession to $Fileserver"
    try {
    $PSsessionFileserver = New-PSSession -computer $fileserver -verbose }
        catch {
            Write-Warning "Unable to create remote PS session to: $fileserver"
            Get-PSSession | Remove-PSSession #cleanup any open sessions
            throw #exiting script
        }
    }
    Write-host "Current remote PSsessions:"
    Get-PSSession
    DashedLine01


#add leading zero if required: to create consitent OUs YEAR07 not YEAR7: 
        if ( $DestOU -le 9) {
            Write-host "Target Year group less than or equal to 9..."
            $UpdatedDestOU = @()
            $UpdatedDestOU = $($GoogleWorkSpaceDest + "-Year" + "0" + $DestOU)
            } else {
            $UpdatedDestOU = $($GoogleWorkSpaceDest + "-Year" + $DestOU)
            }



#Write-Host "invoke-command -computername $UsersFileServer -Scriptblock {Get-SmbOpenFile | where-object {$_.Path -like "*$LegacyShareNoDollar*"} }"
                    #$SMBopenfilesChk = $(invoke-command -computername $UsersFileServer -Scriptblock {Get-SmbOpenFile | where-object {$_.Path -like "*$LegacyShareNoDollar*"} })

#>

