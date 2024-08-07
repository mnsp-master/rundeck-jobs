$mnspver = "0.0.41"
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

Write-host "importing filtered csv data..."
$GmailMessage = import-csv $tempcsv | where { ( $_.'Message-ID' -like '*GLPI_*' -and $_.'Message-ID' -like '*solved*' -and $_.'Message-ID' -like '*wrisch-web05*' ) }

#Getting mail domain from user...
$SenderDomain = $GmailMessage.user.Split("@")[1]
Write-host "sender domain:" $SenderDomain

#remove enclosing chevrons from message id
$MessageID = $GmailMessage.'Message-ID'.TrimStart("<")
$MessageID = $MessageID.TrimEnd(">")
Write-host "removing < and  > from message ID - trimmed ID:" $MessageID

#set message receiver by splitting requester and domain
### DEV ### $MailReceiver = "$ReceiverPrefix@$senderDomain"
Write-Host "Setting message receiver to:" $MailReceiver

$RFCQuery = "'rfc822msgid:$MessageID'" #concatenate message query

$OriginalSubjectSRC = "$MessageID.Subject"

#split original ticket title elements...
$OriginalSubjectSRC = $GmailMessage.Subject
$substring01 = "Ticket Solved"
$substring02 = "MNSP ITSupport"

#ticket title...
$split01 = $OriginalSubjectSRC -Split $substring01
Write-Host "Ticket Title:"$split01[1] #after split
$subject01 = $split01[1]

#ticket number...
$split02 = $split01[0] -split $substring02 -replace("]","")
Write-Host "Ticket Number:"$split02[1]


$Subject = "'$subject01'"

#$subject = "'Modified subject'"

#forward message using RFC message id to mail receiver...
Write-Host "$GamDir\gam.exe user $GLPIGmailAddress forward threads to $MailReceiver query $RFCQuery subject $Subject"

Invoke-expression "$GamDir\gam.exe user $GLPIGmailAddress forward threads to $MailReceiver query $RFCQuery subject $subject doit"

#$exec = @'
#& $GamDir\gam.exe user $GLPIGmailAddress forward threads to $MailReceiver query $RFCQuery subject $Subject doit
#'@

#Invoke-Expression $exec

<#
$exec = @'
& "C:\Program Files\7-Zip\7z.exe" u -mx5 -tzip -r "$DestFileZip" "$DestFile"
'@

Invoke-Expression $exec
#>