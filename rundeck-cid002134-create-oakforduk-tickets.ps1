$mnspver = "0.0.16"
Write-Host $(Get-Date)
Write-Host "MNSP Version" $mnspver

$ErrorActionPreference="Continue"
Set-Location $GamDir

#Get/Confirm Google instance
#Invoke-Expression "$GamDir\gam.exe info domain" 

$Query = "'subject:$GlpiTicketID $GLPITicketSubject'"
#$Query = "'subject:$GlpiTicketID'"

# get message id from search criteria ...
#Write-Host "$GamDir\gam.exe user $GLPIGmailAddress print messages query $Query"
Invoke-Expression "$GamDir\gam.exe user $GLPIGmailAddress print messages query $Query"
Invoke-Expression "$GamDir\gam.exe user $GLPIGmailAddress print messages query $Query" | out-file $tempcsv

Write-host "importing csv data..."
$GmailMessage = import-csv $tempcsv

Write-Host "Getting mail domain from user..."
$SenderDomain = $GmailMessage.user.Split("@")[1]
$SenderDomain

Write-host "removing < and  > from message ID..."
$MessageID = $GmailMessage.'Message-ID'.TrimStart("<")
$MessageID = $MessageID.TrimEnd(">")
$MessageID

Write-Host "Setting message receiver..."
$MailReceiver = "$ReceiverPrefix@$senderDomain"
$MailReceiver

$RFCQuery = "'rfc822msgid:$MessageID'"
$Subject = "Alternate Subject"

Write-host "$GamDir\gam.exe user $GLPIGmailAddress forward threads to $MailReceiver query $RFCQuery subject $Subject"

