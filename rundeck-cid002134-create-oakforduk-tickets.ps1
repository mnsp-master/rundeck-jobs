$mnspver = "0.0.20"
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

#Getting mail domain from user...
$SenderDomain = $GmailMessage.user.Split("@")[1]
Write-host "sender domain:" $SenderDomain

#remove enclosing chevrons from message id
$MessageID = $GmailMessage.'Message-ID'.TrimStart("<")
$MessageID = $MessageID.TrimEnd(">")
Write-host "removing < and  > from message ID - trimmed ID:" $MessageID

#set message receiver by splitting requester and domain
$MailReceiver = "$ReceiverPrefix@$senderDomain"
Write-Host "Setting message receiver to:" $MailReceiver

$RFCQuery = "'rfc822msgid:$MessageID'" #concatenate message query

$Subject = "Alternate Subject"

#forward message using RFC message id to mail receiver...
Write-Host "$GamDir\gam.exe user $GLPIGmailAddress forward threads to $MailReceiver query $RFCQuery subject $Subject"
Invoke-Expression "$GamDir\gam.exe user $GLPIGmailAddress forward threads to $MailReceiver query $RFCQuery subject $Subject"

