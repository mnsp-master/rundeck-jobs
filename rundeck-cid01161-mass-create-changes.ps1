$mnspver = "0.0.0.0.0.1"
Write-Host $(Get-Date)
Write-Host "MNSP Version" $mnspver
Start-Sleep 10
$ErrorActionPreference="Continue"



<#
#create api session to glpi instance...
$SessionToken = Invoke-RestMethod -Verbose "$AppURL/initSession" -Method Get -Headers @{"Content-Type" = "application/json";"Authorization" = "user_token $UserToken";"App-Token"=$AppToken}
#https://www.urldecoder.org/


#close current api session...
Invoke-RestMethod "$AppURL/killSession" -Headers @{"session-token"=$SessionToken.session_token; "App-Token" = "$AppToken"}
#>
