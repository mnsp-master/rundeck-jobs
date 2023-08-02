$mnspver = "0.0.0.0.0.1.5"
$TicketCreateUrl = "$AppURL/Ticket"
$ChangeCreateUrl = "$AppURL/Change"
$SetActiveEntity = "$AppURL/changeActiveEntities"
$EntityAttributesURL = "$AppURL/Entity"

Write-Host $(Get-Date)
Write-Host "MNSP Version" $mnspver
Start-Sleep 10
$ErrorActionPreference="Continue"

#

$TargetEntityIDs = $($TargetEntityIDs.split(','))


#create api session to glpi instance...
$SessionToken = Invoke-RestMethod -Verbose "$AppURL/initSession" -Method Get -Headers @{"Content-Type" = "application/json";"Authorization" = "user_token $UserToken";"App-Token"=$AppToken}
#https://www.urldecoder.org/


#$EntityResult = Invoke-RestMethod "$AppURL/search/Entity?is_deleted=0&as_map=0&range=0-1000000&criteria[0][link]=AND&criteria[0][field]=1&criteria[0][searchtype]=notequals&criteria[0][value]=0&search=Search&itemtype=Entity&start=0" -Headers @{"session-token"=$SessionToken.session_token; "App-Token" = "$AppToken"}

#$entities = $EntityResult.data #convert api search into entities array
#$entities

foreach ($TargetEntityID in $TargetEntityIDs) {
        Write-Host "Getting Entity Info..."
        $GetEntityAttributes = @()
        $GetEntityAttributes = Invoke-RestMethod -Method GET -Uri $EntityAttributesURL/$TargetEntityID -Headers @{"session-token"=$SessionToken.session_token; "App-Token" = "$AppToken" ; "ContentType" = "application/json"}
        $GetEntityAttributes

        Write-Host "Creating Change for entity ID:" $TargetEntityID
        Write-Host "Administrative Number: " $GetEntityAttributes.registration_number

        
        $data = @{
            "input" = @(
                @{
                    "content" = "$ItemDescription"
                    "name" = "$($GetEntityAttributes.registration_number) - $ItemTitle $(Get-Date)"
                    "_users_id_requester" = "47"
                    "entities_id" = "$TargetEntityID"
                    "priority" = "3"
                    "urgency" = "2"
                    "status" = "1"
                    "impact" = "3"
                }
            )
        }


    $json = $data | ConvertTo-Json
    $ApiAction = Invoke-RestMethod -Method POST -Uri $ChangeCreateUrl -Headers @{"session-token"=$SessionToken.session_token; "App-Token" = "$AppToken"} -Body $json -ContentType 'application/json'
    $ApiAction
    Write-Host "Created ID:" $ApiAction.id

}
#close current api session...
Invoke-RestMethod "$AppURL/killSession" -Headers @{"session-token"=$SessionToken.session_token; "App-Token" = "$AppToken"}

<#
Write-Host "get some info from GLPI..."
$EntityResult = @() #empty array
$EntityResult = Invoke-RestMethod "$AppURL/Entity/1" -Headers @{"session-token"=$SessionToken.session_token; "App-Token" = "$AppToken"}

$EntityResult


#>
