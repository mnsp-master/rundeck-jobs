$mnspver = "0.0.76"

<#
Overall process to:
- Begin trarscript logging
- Download google sheet containing all necessary information (MAT wide export from central arbor), filtered based on desired school or use local CSV
- Create any necessary remote Powershell sessions on any file servers where required
- Create a user array (from downloaded/filtered csv data)
    - Loop through the user array, finding the AD object that matches Arbor ID (employeeNumber AD attribute previously synchronised by salamander)
    - Return matched AD user and attributes
    - Use these to determine share hosting server
    - Remotely connect to that file server
        - Update local registry setting representing existing share to reflect any updated username (year number firstname.lastname,renamed homedrive local path (H:\Rmusers\.....\old username etc))
        - Rename local filesystem path to reflect/sync username change
        - Exit remote session
    - Update existing user’s AD attributes to set firstname,lastname,homedir path (renamed share), displayname, usePrincipalName,replacement email address
    - Set custom confidential attribute (mnspAdminNumber)
    - Rename AD object to reflect desired name
    - Repeat for as many users that are in the array
- Close all remote PS sessions
- Stop transcript logging
- Exit PS script
#>

##### ENHANCEMENT ##### general whatif's required to run in dry mode
##### ENHANCEMENT ##### general environment/vars area required to enable PS run outside of rundeck environment

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

##### ENHANCEMENT ##### toggle between local previously manually downloaded gsheet/csv, and geting gsheet each time
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
#ENHANCEMENT - if count 0 break out of script...


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

Write-Host "Updating users in destination..."
foreach ($user in $VerifiedUserData) {
    DashedLine01
    $LegacyUserMail = $user."Email Address (Main)" #current mail address
    $UPN = $user."UPN" # student MIS UPN)
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
        $UserToProcess = $(Get-ADUser -Filter "EmployeeNumber -like '$MISidComplete'" -Properties * | select-object $ADattribs)
        if ($UserToProcess.count -gt 1) {
            Write-Warning "Not an singular match..."
            $UserToProcess
            DashedLine02
                } else {
                    Write-Host "AD attributes found by searching for user with MIS ID: $MISidComplete"
                    $UserToProcess
                    $UsersFileServer = @()
                    $UsersFileServer = $UserToProcess.HomeDirectory.Split("\")[2]
                    $Legacyshare = $UserToProcess.HomeDirectory.Split("\")[3]
                    $ReplacementShare = $ReplacementUserMail.split("@")[0] + "$"
                    $ReplacementShareNoDollar = $ReplacementShare.split('$')[0]

                    Write-host "Legacy Share: $Legacyshare"
                    Write-host "Replacement Share: $ReplacementShare"
                    Write-Host "Replacement Share no Dollar: $ReplacementShareNoDollar"

                        DashedLine02
                        
                        #create remote PS session to users file share host...
                        Write-Host "Remote Session Info:"
                        Invoke-Command -computer  $UsersFileServer -ScriptBlock { #remote share rename scriptblock
                        $env:COMPUTERNAME
                        
                        #Get-ItemProperty $using:RegPath
                        Write-Host "Checking for share: $using:Legacyshare"
                        Get-smbshare -name $using:Legacyshare

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
                        Write-Host "updated multivalue registry key"
                        $test.$using:LegacyShare

                        ##### ENHANCEMENT ##### - appears to be a duplicate of lines 137 - 140
                        $PathToAlter = $test.$using:Legacyshare[3] #local path of share 
                        $PathToAlterVar1 = $PathToAlter.Substring(0, $PathToAlter.lastIndexOf('\')) #split using \ upto last delimeter
                        $PathToAlterVar2 = $PathToAlter.split("\")[-1] #split using \ return last element (username)
                        $PathToAlterOS = $PathToAlter.split("=")[-1] #remove $ from sharename

                        #build new path item
                        $PathToAlterRegItem = $PathToAlterVar1 + "\" + $using:ReplacementShareNoDollar

                        $test.$using:LegacyShare[6] = "ShareName=$using:ReplacementShare"
                        $test.$using:LegacyShare[3] = "$PathToAlterRegItem"

                        Write-host "`n---------`n"

                        Write-Host "updated multivalue registry key"
                        $test.$using:LegacyShare
                        Write-host "`n---------`n"

                        Rename-ItemProperty -Path $using:RegPath -Name $using:Legacyshare -NewName $using:ReplacementShare -verbose #rename registry key ##### ENHANCEMENT ##### whatif required
                        Set-ItemProperty -path $test.PSPath -name $using:ReplacementShare -Value $test.$using:LegacyShare -verbose # update reg key item multi values ##### ENHANCEMENT ##### whatif required

                        #rename existing user home drive folder:
                        Write-host "rename existing folder: $LegacyPathOS to $using:ReplacementShareNoDollar"
                        rename-item -path $LegacyPathOS -NewName $using:ReplacementShareNoDollar -verbose ##### ENHANCEMENT ##### whatif required

                        restart-service LanmanServer -verbose #restart service to reflect updated registry keys/values to present renamed share ##### ENHANCEMENT ##### highly inefficient consider restart of sevice once post mods per server

                        Write-Host "Replacement Share info:"
                        Get-smbshare -name $using:ReplacementShare
                        Write-host "`n---------`n"

                        }
                    Write-Host "updating AD user: "
                        $ReplacementUserPrincipalNameDomain = $UserToProcess.userPrincipalName.split("@")[1] #split using @ select 2nd element
                        $ReplacementUserPrincipalName = $ReplacementShareNoDollar + "@" + $ReplacementUserPrincipalNameDomain #rebuild replacement userPrincipalName
                        $ReplacementShareFull = "\\" + $UsersFileServer + "\" + $ReplacementShare 
                        set-aduser -Identity $UserToProcess.ObjectGUID -GivenName "$FirstName" -surname "$LastName" -email "$ReplacementUserMail" -SamAccountName "$ReplacementShareNoDollar" -DisplayName "$FirstName $LastName" -homeDirectory "$ReplacementShareFull" -userPrincipalName "$ReplacementUserPrincipalName" -verbose ##### ENHANCEMENT ##### whatif required
                        
                        # update mnspAdminNumber attribute
                        Set-ADUser -Identity $UserToProcess.ObjectGUID -Add @{mnspAdminNumber="$UPN"} ##### ENHANCEMENT ##### whatif required

                        $NewName = ($Yearprefix + $FirstName + "." + $LastName).ToLower()
                        get-aduser -Identity $UserToProcess.ObjectGUID | rename-ADobject -NewName "$NewName" -verbose ##### ENHANCEMENT ##### whatif required


                    DashedLine01
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

