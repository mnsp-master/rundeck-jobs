$mnspver = "0.0.7"

Write-Host $(Get-Date)
Write-Host "MNSP Version" $mnspver
#Start-Sleep 10
$ErrorActionPreference="Continue"
Set-Location $GamDir

function DashedLine {
Write-host "-----------------------------------------------------------`n"
}

#create api session to glpi instance...
$SessionToken = Invoke-RestMethod -Verbose "$AppURL/initSession" -Method Get -Headers @{"Content-Type" = "application/json";"Authorization" = "user_token $UserToken";"App-Token"=$AppToken}
#https://www.urldecoder.org/

$GLPIReplacementUserLogin = $GLPIReplacementUserName + "@" + $GoogleWorkspaceMailDomain # replacement loginname

$UserData = Invoke-RestMethod -Method GET -Uri "$AppURL/User/$GLPIuserID" -Headers @{"session-token"=$SessionToken.session_token; "App-Token" = "$AppToken"} -ContentType 'application/json'
Write-Host "Users current DN:" $UserData.user_dn
$updatedUID = "uid=" + $GLPIReplacementUserName + ","
$GLPIReplacementUID = $($UserData.user_dn) -replace '^(.+?),',$updatedUID #regex replace everything upto the first comma of DN (uid) with updated username

#create json content of values to update/set...
               $UpdateData = @{input=@{
               id=$GLPIuserID
               user_dn=$GLPIReplacementUID
               name=$GLPIReplacementUserLogin
               }
               }

           #update login (email attribute) and LDAP directory uid
           $jsonupdate = $updateData | ConvertTo-Json
           $UpdateResult = Invoke-RestMethod "$AppURL/User" -Method Put -Headers @{"session-token"=$SessionToken.session_token; "App-Token" = "$AppToken"} -Body $jsonUpdate -ContentType 'application/json'
           $UpdateResult

           #get users email id from emails table
           $UpdateResult1 = Invoke-RestMethod -Method GET -Uri "$AppURL/User/$GLPIuserID/UserEmail" -Headers @{"session-token"=$SessionToken.session_token; "App-Token" = "$AppToken"} -ContentType 'application/json'
           $updateResult1.id
           $id1 = $updateResult1.id #

           #create json content of values to update/set...
           $UpdateData2 = @{input=@{
           id=$id1
           email=$GLPIReplacementUserLogin
           }
           }
           
           #update users email address in email address table
           $jsonupdate2 = $updateData2 | ConvertTo-Json
           $UpdateResult2 = Invoke-RestMethod -Method PUT -Uri "$AppURL/UserEmail" -Headers @{"session-token"=$SessionToken.session_token; "App-Token" = "$AppToken"} -Body $jsonUpdate2 -ContentType 'application/json'
           $UpdateResult2
DashedLine

#close current api session...
Invoke-RestMethod "$AppURL/killSession" -Headers @{"session-token"=$SessionToken.session_token; "App-Token" = "$AppToken"}

<#

#Write-Host "Invoke-RestMethod "$AppURL/search/User?is_deleted=0&as_map=0&browse=0&criteria[0][link]=AND&criteria[0][field]=2&criteria[0][searchtype]=contains&criteria[0][value]=$GLPIuserID&itemtype=User&start=0" -Headers @{"session-token"=$SessionToken.session_token; "App-Token" = "$AppToken"}"
$UsersData = Invoke-RestMethod "$AppURL/search/User?is_deleted=0&as_map=0&browse=0&criteria[0][link]=AND&criteria[0][field]=2&criteria[0][searchtype]=contains&criteria[0][value]=$GLPIuserID&itemtype=User&start=0" -Headers @{"session-token"=$SessionToken.session_token; "App-Token" = "$AppToken"}

$UserData = @() #reset array
$UserData = Invoke-RestMethod -Method GET -Uri "$AppURL/User/$GLPIuserID" -Headers @{"session-token"=$SessionToken.session_token; "App-Token" = "$AppToken"} -ContentType 'application/json'
$UserData


$GoogleSourceSvcAccount = ("$GoogleServiceAccountPrefix" + "$GoogleWorkSpaceSource" + "@" + "$GGoogleWorkspaceSourceMailDomain")
Write-Host "Google Source Service Account: $GoogleSourceSvcAccount"

$GoogleWorkspaceSourceGroup = ("$GoogleWorkspaceSourceGroupPrefix" + "@" + "$GGoogleWorkspaceSourceMailDomain")
Write-Host "Getting members of users to process source group $GoogleWorkspaceSourceGroup"
Invoke-Expression "$GamDir\gam.exe print group-members group_ns $GoogleWorkspaceSourceGroup > $tempcsv"

#get verified user data
#if exist check & remove $tempcsv4
if (test-path $tempcsv4) { remove-item $tempcsv4 -force -verbose }

Start-sleep 2

Write-Host "downloading gsheet ID: $GoogleSheetID tab: $GoogleSheetTab01" #needs confirmation of approach
Invoke-Expression "$GamDir\gam.exe user $GoogleSourceSvcAccount get drivefile $GoogleSheetID format csv gsheet ""$GoogleSheetTab01"" targetfolder $DataDir targetname $tempcsv4"

$VerifiedUserData = Get-Content -path $tempcsv4 | select-object -skip 1 | convertFrom-csv | where { $_.$FieldMatch01 -like $FieldString } #import where field like $FieldMatch01, and skip 1st line
Write Host "Number of records matching selection criteria:" $VerifiedUserData.count

#if ($uuids.Contains($uuid)) { } # if var is in array


$UsersToProcess = @()
$UsersToProcess = Import-csv $tempcsv

foreach ($user in $UsersToProcess) {
    DashedLine
    $usermail = $User.email
    Write-Host "Processing $usermail"

    #set custom attribute
    
    DashedLine
}







#Start-sleep 5
#Write-host "--------------------------------------`n"

#Write-Host "Setting workspace destination: $GoogleWorkSpaceDestination"
#Invoke-Expression "$GamDir\gam.exe select $GoogleWorkSpaceDestination save" # swap/set google workspace
#Invoke-Expression "$GamDir\gam.exe"

#Write-host "--------------------------------------`n"

#Start-sleep 5
#Write-Host $(Get-Date)

#>