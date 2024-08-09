$mnspver = "0.0.7"

Write-Host $(Get-Date)
Write-Host "MNSP Version" $mnspver
#Start-Sleep 10
$ErrorActionPreference="Continue"
Set-Location $GamDir

#create api session to glpi instance...
$SessionToken = Invoke-RestMethod -Verbose "$AppURL/initSession" -Method Get -Headers @{"Content-Type" = "application/json";"Authorization" = "user_token $UserToken";"App-Token"=$AppToken}
#https://www.urldecoder.org/


#######
$previouslyProcessedbyID = @()
$previouslyProcessedbyID = import-csv $tempcsv1
#$previouslyProcessedbyID

# get all tickets that match search criteria...(value: 12)
$TicketResult = @()
$TicketResult = Invoke-RestMethod "$AppURL/search/Ticket?is_deleted=0&as_map=0&browse=0&criteria[0][link]=AND&criteria[0][field]=23&criteria[0][searchtype]=equals&criteria[0][value]=12&itemtype=Ticket&start=0" -Headers @{"session-token"=$SessionToken.session_token; "App-Token" = "$AppToken"}
$TicketData = $TicketResult.data

Write-Host "Ticket id's found that match search criteria:"
$TicketData.2 #ticket id's only
Write-Host "-----------------------------------------------`n"

#$TicketData = @()
foreach ($Ticket in $TicketData){

$TicketID = @()
$TicketID = $Ticket.2
Write-Host "Assessing Ticket ID:" $TicketID
#check if ticket has been previously processed...
Write-Host "ticket number check:" $previouslyProcessedbyID.GLPITicketID
if ($previouslyProcessedbyID -Match $TicketID) {
    Write-Host "ID: $($TicketID) is a Previously processed Ticket..."
    Write-Host "-----------------------------------------------`n"
    } else {
    Write-Host "ID: $($TicketID) will be processed..."
    Write-Host "GLPI User ID: $($Ticket.4)"
    #get user ID: Details
    $userDetail = @()
    $userDetail = Invoke-RestMethod "$AppURL/search/User?is_deleted=0&as_map=0&browse=0&criteria%5B0%5D%5Blink%5D=AND&criteria%5B0%5D%5Bfield%5D=2&criteria%5B0%5D%5Bsearchtype%5D=contains&criteria%5B0%5D%5Bvalue%5D=$($Ticket.4)&itemtype=User&start=0" -Headers @{"session-token"=$SessionToken.session_token; "App-Token" = "$AppToken"}
    $Mail = @()
    $Mail = $userDetail.data.5
    Write-Host "Users email address:" $Mail
    #process mail using gamxtd
    

    $TicketID | out-file -Append $tempcsv1 #csv must be encoded as UCS-2 LE BOM
    start-sleep 1
    Write-Host "-----------------------------------------------`n"
    
    }
}




Start-Sleep 10

#close current api session...
Invoke-RestMethod "$AppURL/killSession" -Headers @{"session-token"=$SessionToken.session_token; "App-Token" = "$AppToken"}

<#

# all tickets
$TicketResult = @()
$TicketResult = Invoke-RestMethod "$AppURL/search/Ticket?is_deleted=0&as_map=0&browse=0&criteria[0][link]=AND&criteria[0][field]=23&criteria[0][searchtype]=equals&criteria[0][value]=12&itemtype=Ticket&start=0" -Headers @{"session-token"=$SessionToken.session_token; "App-Token" = "$AppToken"}
$TicketResult.data

#All entities: - Production
$EntityResult = Invoke-RestMethod "$AppURL/search/Entity?is_deleted=0&as_map=0&range=0-1000000&criteria[0][link]=AND&criteria[0][field]=1&criteria[0][searchtype]=notequals&criteria[0][value]=0&search=Search&itemtype=Entity&start=0" -Headers @{"session-token"=$SessionToken.session_token; "App-Token" = "$AppToken"}

$entities = $EntityResult.data #convert api search into entities array
$SearchResult=@()

#>
