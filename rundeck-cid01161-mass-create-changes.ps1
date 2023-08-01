$mnspver = "0.0.0.0.0.7"
Write-Host $(Get-Date)
Write-Host "MNSP Version" $mnspver
Start-Sleep 10
$ErrorActionPreference="Continue"

#

$TargetEntityIDs = $($TargetEntityIDs.split(','))
foreach ($TargetEntityID in $TargetEntityIDs) {
        Write-Host "Target Entity ID:" $TargetEntityID
}

#<#
#create api session to glpi instance...
$SessionToken = Invoke-RestMethod -Verbose "$AppURL/initSession" -Method Get -Headers @{"Content-Type" = "application/json";"Authorization" = "user_token $UserToken";"App-Token"=$AppToken}
#https://www.urldecoder.org/

#get all entities from GLPI
$EntityResult = @() #empty array
$EntityResult = Invoke-RestMethod "$AppURL/getActiveEntities" -Headers @{"session-token"=$SessionToken.session_token; "App-Token" = "$AppToken"}

$EntityResult

#$EntityResult = Invoke-RestMethod "$AppURL/search/Entity?is_deleted=0&as_map=0&range=0-1000000&criteria[0][link]=AND&criteria[0][field]=1&criteria[0][searchtype]=notequals&criteria[0][value]=0&search=Search&itemtype=Entity&start=0" -Headers @{"session-token"=$SessionToken.session_token; "App-Token" = "$AppToken"}

#$entities = $EntityResult.data #convert api search into entities array
#$entities

#close current api session...
Invoke-RestMethod "$AppURL/killSession" -Headers @{"session-token"=$SessionToken.session_token; "App-Token" = "$AppToken"}
#>
