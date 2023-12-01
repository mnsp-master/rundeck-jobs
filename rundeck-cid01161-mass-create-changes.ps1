$mnspver = "0.0.0.0.0.3.2.5"
$TicketCreateUrl = "$AppURL/Ticket"
$ChangeCreateUrl = "$AppURL/Change"
$SetActiveEntity = "$AppURL/changeActiveEntities"
$EntityAttributesURL = "$AppURL/Entity"
$ProjectUpdateUrl = "$AppURL/Project"

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

################################ return GLPI plugin additional fields IDs #######################################
$ApiSearchResult = Invoke-RestMethod "$AppURL/listSearchOptions/Entity" -Headers @{"session-token"=$SessionToken.session_token; "App-Token" = "$AppToken"} # api serach query for glpi entities
$ApiSearchResult  | out-file -FilePath $temptxt # output api entity query to tmp txt doc
$ApiSearchResultSummary = Get-Content $temptxt | where-object {$_ -Like "*MNSP IT Adhoc*"} | Select-Object #filter to only include specific Plugin generated ID's

#$ApiSearchResultSummary
#get specific GLPI plugin additional fields object IDs...
$GLPIsearchStringSchoolType = "MNSP IT Adhoc - School Type"
$MNSPSchoolTypeID = $($ApiSearchResultSummary | where-Object {$_ -Like "*$GLPIsearchStringSchoolType*"}).split(":")[0].TrimEnd()
Write-Host "$GLPIsearchStringSchoolType ID: ---$MNSPSchoolTypeID---"

$GLPIsearchStringLevel3ITEngineer = "MNSP IT Adhoc - Level 3 IT Engineer"
$MNSPLevel3EngineerID = $($ApiSearchResultSummary | where-Object {$_ -Like "*$GLPIsearchStringLevel3ITEngineer*"}).split(":")[0].TrimEnd()
Write-Host "$GLPIsearchStringLevel3ITEngineer ID: ---$MNSPLevel3EngineerID---"

$GLPIsearchStringSchoolNameCode = "MNSP IT Adhoc - School Name Code"
$MNSPSchoolNameCodeID = $($ApiSearchResultSummary | where-Object {$_ -Like "*$GLPIsearchStringSchoolNameCode*"}).split(":")[0].TrimEnd()
Write-Host "$GLPIsearchStringSchoolNameCode ID: ---$MNSPSchoolNameCodeID---"

#close current api session...
#Invoke-RestMethod "$AppURL/killSession" -Headers @{"session-token"=$SessionToken.session_token; "App-Token" = "$AppToken"}
#exit


#$GLPIsearchStringGoogleWorkspaceChromebookBaseOU = "MNSP IT Adhoc - Google workspace chrome book base OU" #search string to return Plugin object ID
#$GoogleWorkspaceChromebookBaseOUID = $($ApiSearchResultSummary | Where-Object {$_ -Like "*$GLPIsearchStringGoogleWorkspaceChromebookBaseOU*"}).split(":")[0].TrimEnd() #get headteacher ID
#Write-host "$GLPIsearchStringGoogleWorkspaceChromebookBaseOU ID: ---$GoogleWorkspaceChromebookBaseOUID---"


#determine if input value is all primaries/secondaries/AP etc....
$EntitiesResult =@()
$apiQueryALL = @()
$SchoolType = @()

if ( $TargetEntityIDs -eq "1000" ) {
    Write-Host "Primaries ONLY...."
    $SchoolType = "1" # too hard coded - needs lookup mechanism
    $apiQueryALL = "?criteria[1][link]=AND&criteria[1][field]=$MNSPSchoolTypeID&criteria[1][searchtype]=equals&criteria[1][value]=$SchoolType&itemtype=Entity&start=0" #primaries
    $EntitiesResult = Invoke-RestMethod "$AppURL/search/Entity$apiQueryALL" -Headers @{"session-token"=$SessionToken.session_token; "App-Token" = "$AppToken"}
    $TargetEntityIDs = $EntitiesResult.data.2
    $TargetEntityIDs
    #close current api session...
    #Invoke-RestMethod "$AppURL/killSession" -Headers @{"session-token"=$SessionToken.session_token; "App-Token" = "$AppToken"}
    #exit
    }

elseif 
    ( $TargetEntityIDs -eq "1001" ) {
    Write-Host "Secondaries ONLY...."
    $SchoolType = "2" # too hard coded - needs lookup mechanism
    $apiQueryALL = "?criteria[1][link]=AND&criteria[1][field]=$MNSPSchoolTypeID&criteria[1][searchtype]=equals&criteria[1][value]=$SchoolType&itemtype=Entity&start=0" #primaries
    $EntitiesResult = Invoke-RestMethod "$AppURL/search/Entity$apiQueryALL" -Headers @{"session-token"=$SessionToken.session_token; "App-Token" = "$AppToken"}
    $TargetEntityIDs = $EntitiesResult.data.2
    $TargetEntityIDs
    #Invoke-RestMethod "$AppURL/killSession" -Headers @{"session-token"=$SessionToken.session_token; "App-Token" = "$AppToken"}
    #exit
    }


foreach ($TargetEntityID in $TargetEntityIDs) {
        Write-Host "Getting Entity Info..."
        $GetEntityAttributes = @()
        $GetEntityAttributes = Invoke-RestMethod -Method GET -Uri $EntityAttributesURL/$TargetEntityID -Headers @{"session-token"=$SessionToken.session_token; "App-Token" = "$AppToken" ; "ContentType" = "application/json"}
        $GetEntityAttributes

        #get additional fields plugin values
        $apiQuerySpecificID = "?as_map=0&browse=0&criteria[0][link]=AND&criteria[0][field]=76694&criteria[0][searchtype]=contains&criteria[0][value]=&criteria[1][link]=OR&criteria[1][field]=76692&criteria[1][searchtype]=contains&criteria[1][value]=&criteria[2][link]=OR&criteria[2][field]=76684&criteria[2][searchtype]=contains&criteria[2][value]=&criteria[3][link]=AND&criteria[3][field]=2&criteria[3][searchtype]=contains&criteria[3][value]=$TargetEntityID&itemtype=Entity&start=0"

        $EntityResult = Invoke-RestMethod "$AppURL/search/Entity$apiQuerySpecificID" -Headers @{"session-token"=$SessionToken.session_token; "App-Token" = "$AppToken"}
        Write-host "additional fields plugin values..."
        $EntityResult.data

        Write-Host "Creating Change for entity ID:" $TargetEntityID
        #Write-Host "Administrative Number: " $GetEntityAttributes.registration_number
        #Write-Host "SchoolNameCode: " $GetEntityAttributes.$GLPIsearchStringSchoolNameCodeID

        $dataName = @()
        #$dataName = "$($EntityResult.data.76694) - $ItemTitle" # $(Get-Date)
        #$dataName = "$($EntityResult.data.$MNSPSchoolNameCodeID) - $ItemTitle" # $(Get-Date)
        $dataName = "$($EntityResult.data.76684) - $ItemTitle" # $(Get-Date)
        #$($EntityResult.data.$MNSPSchoolNameCodeID)
        #$dataUsersIdAssign = $($EntityResult.data.76692)
        $dataUsersIdAssign = $($EntityResult.data.76683)
        #$dataUsersIdAssign = $($EntityResult.data.$MNSPLevel3EngineerID)
        #$($EntityResult.data.$MNSPLevel3EngineerID)

        #Write-Host "SNO 01..."
        #$dataName


        $data = @{
            "input" = @(
                @{
                    "content" = "$ItemDescription"
                    "name" = "$dataName"
                    "_users_id_requester" = "47"
                    "_users_id_assign" = "$dataUsersIdAssign"
                    "entities_id" = "$TargetEntityID"
                    "priority" = "3"
                    "urgency" = "2"
                    "status" = "1"
                    "impact" = "3"
                    "itilcategories_id" = "$GLPIITILCategoryID"
                }
            )
        }


    $json = $data | ConvertTo-Json
    $ApiAction = Invoke-RestMethod -Method POST -Uri $ChangeCreateUrl -Headers @{"session-token"=$SessionToken.session_token; "App-Token" = "$AppToken"} -Body $json -ContentType 'application/json'
    $ApiAction
    $CreatedChangeID = $($ApiAction.id)
    Write-Host "Created Change ID: " $CreatedChangeID

    #Link Change with Project ID:

        $ProjectData = @{
        "input" = @(
            @{
                "itemtype" = "Change"
                "items_id" = "$CreatedChangeID"
                "projects_id" = "$LinkedProjectID"
            }
        )
    }
    $ProjectDataJson = $ProjectData | ConvertTo-Json
    Invoke-RestMethod -Method POST -Uri $ProjectUpdateUrl/$LinkedProjectID/Itil_Project -Headers @{"session-token"=$SessionToken.session_token; "App-Token" = "$AppToken"} -Body $ProjectDataJson -ContentType 'application/json'

    ##create project task##
    $ProjectUpdateUrl = "$AppURL/Project"
        $ProjectData = @{
            "input" = @(
                @{
                    "itemtype" = "projecttask"
                    "projects_id" = "$LinkedProjectID"
                    "name" = "$dataName"
                    "content" = "$ItemDescription"
                    "projectstates_id" = "1"

                }
            )
        }
        $ProjectDataJson = $ProjectData | ConvertTo-Json
        Invoke-RestMethod -Method POST -Uri $ProjectUpdateUrl/$LinkedProjectID/ProjectTask -Headers @{"session-token"=$SessionToken.session_token; "App-Token" = "$AppToken"} -Body $ProjectDataJson -ContentType 'application/json' # 



}
#close current api session...
Invoke-RestMethod "$AppURL/killSession" -Headers @{"session-token"=$SessionToken.session_token; "App-Token" = "$AppToken"}

<#


$GLPIsearchStringSchoolNameCodeID = $($ApiSearchResultSummary | Where-Object {$_ -Like "*$GLPIsearchStringSchoolNameCode*"}).split(":")[0].TrimEnd() #get SchoolNameCode ID
Write-host "$GLPIsearchStringSchoolNameCode ID: ---$GLPIsearchStringSchoolNameCodeID---"


#$GLPIsearchStringHeadTeacher = "MNSP IT Adhoc - Head Teacher" #search string to return Plugin object ID
#$GLPIsearchStringSchoolNameCode = "MNSP IT Adhoc - SchoolNameCode" #search string to return Plugin object ID

#$GLPIsearchStringMNSPUpdateGoogleChromeDeviceUserAttribute = "MNSP IT Adhoc - Update Google Chrome Device user attribute" #search string to return Plugin object ID


#$GoogleWorkspaceChromebookBaseOUID = $($ApiSearchResultSummary | Where-Object {$_ -Like "*$GLPIsearchStringGoogleWorkspaceChromebookBaseOU*"}).split(":")[0].TrimEnd() #get headteacher ID
#Write-host "$GLPIsearchStringGoogleWorkspaceChromebookBaseOU ID: ---$GoogleWorkspaceChromebookBaseOUID---"

#$UpdateGoogleChromeDeviceUserAttributeID = $($ApiSearchResultSummary | Where-Object {$_ -Like "*$GLPIsearchStringMNSPUpdateGoogleChromeDeviceUserAttribute*"}).split(":")[0].TrimEnd() #get headteacher ID
#Write-host "$GLPIsearchStringMNSPUpdateGoogleChromeDeviceUserAttribute ID: ---$UpdateGoogleChromeDeviceUserAttributeID---"

Write-Host "get some info from GLPI..."
$EntityResult = @() #empty array
$EntityResult = Invoke-RestMethod "$AppURL/Entity/1" -Headers @{"session-token"=$SessionToken.session_token; "App-Token" = "$AppToken"}

$EntityResult


#>
