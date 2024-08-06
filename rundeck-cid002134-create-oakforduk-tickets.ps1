$mnspver = "0.0.7"
Write-Host $(Get-Date)
Write-Host "MNSP Version" $mnspver

$ErrorActionPreference="Continue"
Set-Location $GamDir

#Get/Confirm Google instance
Invoke-Expression "$GamDir\gam.exe info domain" 

$Query = "'subject:$GlpiTicketID $GLPITicketSubject'"

# get message id from search criteria ...
Invoke-Expression "$GamDir\gam.exe user $GLPIGmailAddress print messages query $Query" > $tempcsv

