$mnspver = "0.0.6"

Write-Host $(Get-Date)
Write-Host "MNSP Version" $mnspver
#Start-Sleep 10
$ErrorActionPreference="Continue"
Set-Location $GamDir

#create api session to glpi instance...
$SessionToken = Invoke-RestMethod -Verbose "$AppURL/initSession" -Method Get -Headers @{"Content-Type" = "application/json";"Authorization" = "user_token $UserToken";"App-Token"=$AppToken}
#https://www.urldecoder.org/

# all tickets
$TicketResult = @()
$TicketResult = Invoke-RestMethod "$AppURL/search/Ticket?is_deleted=0&as_map=0&browse=0&criteria[0][link]=AND&criteria[0][field]=23&criteria[0][searchtype]=equals&criteria[0][value]=12&itemtype=Ticket&start=0" -Headers @{"session-token"=$SessionToken.session_token; "App-Token" = "$AppToken"}
$TicketResult.data


Start-Sleep 10

#close current api session...
Invoke-RestMethod "$AppURL/killSession" -Headers @{"session-token"=$SessionToken.session_token; "App-Token" = "$AppToken"}

<#
#All entities: - Production
$EntityResult = Invoke-RestMethod "$AppURL/search/Entity?is_deleted=0&as_map=0&range=0-1000000&criteria[0][link]=AND&criteria[0][field]=1&criteria[0][searchtype]=notequals&criteria[0][value]=0&search=Search&itemtype=Entity&start=0" -Headers @{"session-token"=$SessionToken.session_token; "App-Token" = "$AppToken"}

$entities = $EntityResult.data #convert api search into entities array
$SearchResult=@()

#>
