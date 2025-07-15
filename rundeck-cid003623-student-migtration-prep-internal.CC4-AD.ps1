$mnspver = "0.0.133"

<#
Overall process to:
- Begin transcript logging
- Download google sheet - or use local/manually saved CSV, containing all necessary information (MAT wide export from central arbor), data will be filtered based on desired school
- Create any necessary remote Powershell sessions on any file servers where required
- Create a user array (from downloaded/filtered csv data)
    - Loop through the user array, finding the AD object that matches Arbor ID (employeeNumber AD attribute previously synchronised by salamander)
    - Return matched AD user and attributes
    - Use these to determine share hosting server
    - Remotely connect to that file server
        - Update local registry setting representing existing share to reflect any updated username (year number firstname.lastname,renamed homedrive local path (H:\Rmusers\.....\old username etc))
        - Rename local filesystem path to reflect/sync username change
        - Exit PS remote session
    - Update existing users AD attributes to set firstname,lastname,homedir path (renamed share), displayname, usePrincipalName,replacement email address
    - Set custom confidential attribute (mnspAdminNumber)
    - Rename AD object to reflect desired name
    - Repeat for as many users that are in the array
- Close all remote PS sessions
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

Write-Host "gsheet Student number column heading:" $FieldMatch01
Start-sleep 10

#prepare user details csv
Write-Host "emptying $tempcsv2 of any existing data..."
Clear-Content $tempcsv2
sleep 1
$UserInfoCSVheader | out-file -filepath $tempcsv2 -Append #create blank csv with simple header

##### ENHANCEMENT ##### toggle between local previously manually downloaded gsheet/csv, and downloading gsheet each time
#Set Google Instance: Destination...
Write-Host "###### Set Google instance: Destination... ######"

$GoogleSvcAccount = $GoogleWorkspaceMNSPsvcAccount
Write-Host "Google Destination Service Account: $GoogleSvcAccount"

Write-Host "Setting workspace Destination: $GoogleWorkSpaceDestination"
Invoke-Expression "$GamDir\gam.exe select $GoogleWorkSpaceDestination save" # swap/set google workspace
Invoke-Expression "$GamDir\gam.exe"
start-sleep 3
DashedLine01

#get verified user data
#if exist check & remove $tempcsv4
if (test-path $tempcsv4) { remove-item $tempcsv4 -force -verbose }

##### ENHANCEMENT ##### ERROR TO ADDRESS - Duplicate fileds Value with production gsheet #####
Write-Host "downloading gsheet ID: $GoogleSheetID tab: $GoogleSheetTab01"
#Write-Host "$GamDir\gam.exe user $GoogleSvcAccount get drivefile $GoogleSheetID format csv gsheet ""$GoogleSheetTab01"" targetfolder $DataDir targetname $tempcsv4"
Invoke-Expression "$GamDir\gam.exe user $GoogleSvcAccount get drivefile $GoogleSheetID format csv gsheet ""$GoogleSheetTab01"" targetfolder $DataDir targetname $tempcsv4"
#Write-Host "Invoke-Expression $GamDir\gam.exe user $GoogleSourceSvcAccount get drivefile $GoogleSheetID format csv gsheet ""$GoogleSheetTab01"" targetfolder $DataDir targetname $tempcsv4"

Start-sleep 2

Write-Host "Field match:  " $FieldMatch01
Write-Host "Field String: " $FieldString
#$VerifiedUserData = Get-Content -path $tempcsv4 | where { $_.$FieldMatch01 -like $FieldString }
$VerifiedUserData = Get-Content -path $tempcsv4 | convertFrom-csv | where { $_.$FieldMatch01 -like $FieldString } #import where field like $FieldMatch01

#$VerifiedUserData = Get-Content -path $tempcsv4 | convertFrom-csv | where-object { 
#    $_.$FieldMatch01 -like $FieldString -and 
#    $_.$Fieldmatch02 -match '^[0-9]+$' #Numeric values only - excludes - R N1 N2 etc
#    } #import where field like $FieldMatch01


#$VerifiedUserData = Get-Content -path $tempcsv4 | select-object -skip 1 | convertFrom-csv | where { $_.$FieldMatch01 -like $FieldString } #import where field like $FieldMatch01, and skip 1st line
Write-Host "Number of records matching selection criteria:" $VerifiedUserData.count
##### ENHANCEMENT ##### - if count 0 break out of script...


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

if (test-path $tempcsv6) { remove-item $tempcsv6 -force -verbose }
if (test-path $tempcsv8) { remove-item $tempcsv8 -force -verbose }

start-sleep 2

#OU information to ultimetly move updated user:
Write-Host "AD search base DN: $OUBaseDn"
$OUS = $(Get-AdOrganizationalUnit -searchbase $OUBaseDn -Filter *) # get all OU's from baseDN

Write-Host "Updating users..."
foreach ($user in $VerifiedUserData) {
    DashedLine01
    Write-Host "PROCESSING next user..."
    $LegacyUserMail = $user."Email Address (Main)" #current mail address
    $UPN = $user."UPN" # student MIS UPN
    $FirstName = $user."Modified_Preferred_firstname" #prefered firstname 
    $LastName = $user."Modified_Preferred_Lastname" #prefered lastname
    #$ReplacementUserMail = $user."new email"
    $ReplacementUserMail = $user."Email20Chars" #UPDATE NEEDED ### Column heading needs agreeing
    $DestOU = [int] $user."NC Year(s) for today" #set var as interger
    $MISid = $user."Arbor ID" # DEV 
    $MISidComplete = "$MISsitePrefix-$MISid" #concatenate sitename hyphen and MIS id number e.g: SCH-292 students SCH-Stf_34 ##### ENHANCEMENT ##### if staff/student var needed as MISID's are of dirrent construction
    $Yearprefix = $user."New prefix"
    #$MISid = $user."Arbor Student ID" #Production

    #add leading zero if required: to create consitent OUs YEAR07 not YEAR7: 
        if ( $DestOU -le 9) {
            Write-host "Target Year group less than or equal to 9..."
            $UpdatedDestOU = @()
            $UpdatedDestOU = $($GoogleWorkSpaceDest + "-Year" + "0" + $DestOU)
            } else {
            $UpdatedDestOU = $($GoogleWorkSpaceDest + "-Year" + $DestOU)
            }

    #user info/attributes:
    Write-Host "Legacy email: $LegacyUserMail"
    Write-Host "Replacement email: $ReplacementUserMail"
    Write-Host "UPN: $UPN"
    Write-Host "Firstname: $FirstName"
    Write-Host "Lastname: $LastName"
    Write-Host "Source Year: $DestOU" 
    Write-Host "Destination OU name: $UpdatedDestOU"
    Write-Host "MIS ID: $MISid"
    Write-Host "Complete MIS ID: $MISidComplete"
    DashedLine02

    if ($MISid) { #check value is not null...
        $UserToProcess = @()
        $UserToProcess = $(Get-ADUser -Filter "EmployeeNumber -like '$MISidComplete'" -Properties * | select-object $ADattribs) ####ENHANCEMENT#### control group needed if user member skip
        if ($UserToProcess.count -gt 1) {
            Write-Warning "Not an singular match for employeeNumber $MISidComplete..." #WARNING if not a singular value, do not process current user as uniqueness cannot be confirmed...
            $UserToProcess
            DashedLine02
                } else {
                    Write-Host "AD attributes found by searching for user with MIS ID: $MISidComplete"
                    $UserToProcess
                    
                    #Build Vars from current $UserToProcess object...
                    $UsersFileServer = @()
                    $UsersFileServer = $UserToProcess.HomeDirectory.Split("\")[2]
                    $Legacyshare = $UserToProcess.HomeDirectory.Split("\")[3]
                    $LegacyShareNoDollar = $Legacyshare.split('$')[0]
                    $ReplacementShare = $ReplacementUserMail.split("@")[0] + "$"
                    $ReplacementShareNoDollar = $ReplacementShare.split('$')[0]

                    Write-host "Legacy Share: $Legacyshare"
                    Write-Host "Legacy Share no Dollar:" $LegacyShareNoDollar
                    Write-host "Replacement Share: $ReplacementShare"
                    Write-Host "Replacement Share no Dollar: $ReplacementShareNoDollar"

                    #smb openfile check...
                    Write-Host "Checking for open files from share: $($UserToProcess.HomeDirectory)"
                    $SMBopenfilesChk = @()
                    $sessn = @()
                    $sessn = new-cimsession -ComputerName $UsersFileServer
                    Write-Host "CIM session details:"
                    Get-CimSession
                    $SMBopenFilesChk =$( Get-SmbOpenFile -CimSession $sessn | where {$_.Path -like "*$LegacyShareNoDollar*"})
                    Remove-CimSession $sessn
                    Write-Host "Current share: $($UserToProcess.HomeDirectory) open file count:" $SMBopenfilesChk.count

                            if (!$SMBopenfilesChk.count -ge 1 ) {
                            Write-Host "no files currently open from share, proceeding...."

                                    DashedLine02
                                    
                                        #create remote PS session to users file share host...
                                        Write-Host "Remote Session Progress..."
                                        Invoke-Command -computer  $UsersFileServer -ScriptBlock { #remote share rename scriptblock
                                        Write-Host "Local hostname:" $env:COMPUTERNAME
                                        
                                        #Get-ItemProperty $using:RegPath
                                        Write-Host "Checking for share: $using:Legacyshare"
                                        #Get-smbshare -name $using:Legacyshare

                                        $test = Get-ItemProperty $using:RegPath -Name $using:Legacyshare #get current multi value reg key/values
                                        
                                        Write-Host "Existing Multi vale registry key:"
                                        $test.$using:Legacyshare
                                        Write-host "---------`n"

                                        $PathToAlter = $test.$using:Legacyshare[3] #local path of share
                                        $PathToAlterVar1 = $PathToAlter.Substring(0, $PathToAlter.lastIndexOf('\')) #split using \ upto last delimeter
                                        $PathToAlterVar2 = $PathToAlter.split("\")[-1] #split using \ return last element (username)
                                        $PathToAlterOS = $PathToAlter.split("=")[-1] #remove $ from sharename
                                        
                                        $LegacyPathOS = @()
                                        $LegacyPathOStemp = @()
                                        $LegacyPathOStemp = $test.$using:LegacyShare[3]
                                        $LegacyPathOS = $LegacyPathOStemp.split("=")[1]
                                        
                                        
                                        Write-Host "LegacyPathOS: " $LegacyPathOS
                                        Write-Host "PathToAlterOS:" $PathToAlterOS

                                        #build new registry path items...
                                        $PathToAlterRegItem = $PathToAlterVar1 + "\" + $using:ReplacementShareNoDollar
                                        $test.$using:LegacyShare[6] = "ShareName=$using:ReplacementShare"
                                        $test.$using:LegacyShare[3] = "$PathToAlterRegItem"

                                        Write-host "---------`n"
                                        Write-Host "updated multivalue registry key:"
                                        $test.$using:LegacyShare

                                        #build new path item
                                        $PathToAlterRegItem = $PathToAlterVar1 + "\" + $using:ReplacementShareNoDollar

                                        $test.$using:LegacyShare[6] = "ShareName=$using:ReplacementShare"
                                        $test.$using:LegacyShare[3] = "$PathToAlterRegItem"

                                        Write-host "`n---------`n"

                                        #Write-Host "updated multivalue registry key"
                                        #$test.$using:LegacyShare
                                        #Write-host "`n---------`n"

                                        Write-Host "PS to process: Rename-ItemProperty -Path $using:RegPath -Name $using:Legacyshare -NewName $using:ReplacementShare -verbose"
                                        Rename-ItemProperty -Path $using:RegPath -Name $using:Legacyshare -NewName $using:ReplacementShare -verbose ##-whatif ## Comment Whatif to Action
                                        
                                        Write-Host "PS to process: Set-ItemProperty -path $test.PSPath -name $using:ReplacementShare -Value $test.$using:LegacyShare -verbose"
                                        Set-ItemProperty -path $test.PSPath -name $using:ReplacementShare -Value $test.$using:LegacyShare -verbose ##-whatif ## Comment Whatif to Action

                                        #rename existing user home drive folder:
                                        Write-host "rename existing folder: $LegacyPathOS to $using:ReplacementShareNoDollar"
                                        Write-Host "PS to process: rename-item -path $LegacyPathOS -NewName $using:ReplacementShareNoDollar -verbose"
                                        rename-item -path $LegacyPathOS -NewName $using:ReplacementShareNoDollar -verbose ##-whatif ## Comment Whatif to Action

                                        #restart service to reflect updated registry keys/values to present renamed share ##### ENHANCEMENT ##### highly inefficient consider restart of sevice once post mods per server
                                        restart-service LanmanServer -verbose ##-whatif ## Comment Whatif to Action 
                                        
                                        Write-host "`n---------`n"
                                        Write-Host "Replacement Share info: $using:ReplacementShare (NOTE: will not report/find as expected if in Whatif Mode...)"
                                        $ReplacementShareInfo =@()
                                        $ReplacementShareInfo = Get-smbshare -name $using:ReplacementShare
                                        $ReplacementShareInfo

                                        <#
                                        $ReplacementShareInfo =@()
                                        $ReplacementShareInfo = Get-smbshare -name $using:ReplacementShare 2> $Null

                                            if ($ReplacementShareInfo) {
                                                Write-Host "Share Information: "$ReplacementShareInfo
                                                } else {
                                                Write-Warning "share: $using:ReplacementShare does NOT exist..."
                                            }
                                        #>    
                                        Write-host "`n---------`n"

                                        } #end of remote pssession

                                    Write-Host "updating AD user...: "
                                    Write-host "`n---`n"
                                    $ReplacementUserPrincipalNameDomain = $UserToProcess.userPrincipalName.split("@")[1] #split using @ select 2nd element
                                    $ReplacementUserPrincipalName = $ReplacementShareNoDollar + "@" + $ReplacementUserPrincipalNameDomain #rebuild replacement userPrincipalName
                                    $ReplacementShareFull = "\\" + $UsersFileServer + "\" + $ReplacementShare 
                                    
                                    # update multiple existing user attributes...
                                    Write-Host "PS to process: set-aduser -Identity $($UserToProcess.ObjectGUID) -GivenName "$FirstName" -surname "$LastName" -email "$ReplacementUserMail" -SamAccountName "$ReplacementShareNoDollar" -DisplayName "$FirstName $LastName" -homeDirectory "$ReplacementShareFull" -userPrincipalName "$ReplacementUserPrincipalName" -verbose"
                                    #Write-host "`n---`n"
                                    set-aduser -Identity $($UserToProcess.ObjectGUID) -GivenName "$FirstName" -surname "$LastName" -email "$ReplacementUserMail" -SamAccountName "$ReplacementShareNoDollar" -DisplayName "$FirstName $LastName" -homeDirectory "$ReplacementShareFull" -userPrincipalName "$ReplacementUserPrincipalName" -verbose ##-whatif ## Comment Whatif to Action
                                    Write-host "`n---`n"
                                    
                                    # update mnspAdminNumber attribute...
                                    Write-Host "PS to process: Set-ADUser -Identity $($UserToProcess.ObjectGUID) -Add @{mnspAdminNumber="$UPN"} -verbose`n"
                                    #Write-host "`n---`n"
                                    Set-ADUser -Identity $($UserToProcess.ObjectGUID) -Add @{mnspAdminNumber="$UPN"} -verbose ##-whatif ## Comment Whatif to Action
                                    Write-host "`n---`n"
                                    
                                    # rename existing AD user object...
                                    $NewName = ($Yearprefix + $FirstName + "." + $LastName).ToLower()
                                    Write-Host "PS to process: get-aduser -Identity $($UserToProcess.ObjectGUID) | rename-ADobject -NewName $NewName -verbose`n"
                                    #Write-host "`n---`n"
                                    get-aduser -Identity $($UserToProcess.ObjectGUID) | rename-ADobject -NewName "$NewName" -verbose ##-whatif ## Comment Whatif to Action
                                    Write-host "`n---`n"

                                    ####ENHANCEMENT#### move user to replacement AD OU
                                    $DestADOU = $OUS | where-object {$_ -like "*$UpdatedDestOU*"}
                                    Write-Host "Moving user to Destination OU: $($DestADOU.DistinguishedName)"
                                    Write-Host "PS to process: Move-ADobject -id $($UserToProcess.ObjectGUID) -TargetPath $($DestADOU.DistinguishedName) -verbose"
                                    Move-ADobject -id $($UserToProcess.ObjectGUID) -TargetPath $($DestADOU.DistinguishedName) -verbose ##-whatif ## Comment Whatif to Action
                                    Write-host "`n---`n"

                                    Write-Host "Updated AD users attributes using GUID: (NOTE: will not change if in Whatif Mode...)"
                                    $UserToProcessPostupdate = $(Get-ADUser -id $($UserToProcess.ObjectGUID) -Properties * | select-object $ADattribs)
                                    $UserToProcessPostupdate
                                    DashedLine01
                                    

                                DashedLine01
                            } else {
                                Write-Warning "user: $($UserToProcess.samAccountName) has $($SMBopenfilesChk.count) files Open from share: $($UserToProcess.HomeDirectory) `nABANDONING any processing of account, users MUST be fully logged out to sucessfully rename/update share configuration..."
                                
                                #$SMBopenfilesChk
                    }
                    DashedLine02

                }
    } else { #if MIS id is NULL...
        Write-Warning "No MIS ID found for:"
        Write-host "Legacy email: $LegacyUserMail"
        Write-Host "Replacement email: $ReplacementUserMail"
        Write-Host "UPN: $UPN"
        Write-Host "Firstname: $FirstName"
        Write-Host "Lastname: $LastName"
    }
}



Write-Host "Closing all remote PSSessions..."
Get-PSSession | Remove-PSSession

<#
#Write-Host "invoke-command -computername $UsersFileServer -Scriptblock {Get-SmbOpenFile | where-object {$_.Path -like "*$LegacyShareNoDollar*"} }"
                    #$SMBopenfilesChk = $(invoke-command -computername $UsersFileServer -Scriptblock {Get-SmbOpenFile | where-object {$_.Path -like "*$LegacyShareNoDollar*"} })

##### ENHANCEMENT ##### - appears to be a duplicate of lines 189 - 192
                                        #$PathToAlter = $test.$using:Legacyshare[3] #local path of share 
                                        #$PathToAlterVar1 = $PathToAlter.Substring(0, $PathToAlter.lastIndexOf('\')) #split using \ upto last delimeter
                                        #$PathToAlterVar2 = $PathToAlter.split("\")[-1] #split using \ return last element (username)
                                        #$PathToAlterOS = $PathToAlter.split("=")[-1] #remove $ from sharename


#>

