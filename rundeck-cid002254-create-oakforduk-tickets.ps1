$mnspver = "0.0.19"

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

$GLPITicketID = @()
$GLPITicketID = $Ticket.2
Write-Host "Assessing Ticket ID:" $GLPITicketID
#check if ticket has been previously processed (imported csv data)...
Write-Host "ticket number check:" $previouslyProcessedbyID.TicketID
if ($previouslyProcessedbyID -Match $GLPITicketID) {
    Write-Host "ID: $($GLPITicketID) is a Previously processed Ticket..."
    Write-Host "-----------------------------------------------`n"
    } else {
    Write-Host "ID: $($GLPITicketID) will be processed..."
    Write-Host "GLPI User ID: $($Ticket.4)"
    #get user ID: Details
    $userDetail = @()
    $userDetail = Invoke-RestMethod "$AppURL/search/User?is_deleted=0&as_map=0&browse=0&criteria%5B0%5D%5Blink%5D=AND&criteria%5B0%5D%5Bfield%5D=2&criteria%5B0%5D%5Bsearchtype%5D=contains&criteria%5B0%5D%5Bvalue%5D=$($Ticket.4)&itemtype=User&start=0" -Headers @{"session-token"=$SessionToken.session_token; "App-Token" = "$AppToken"}
    $GLPIGmailAddress = @()
    $GLPIGmailAddress = $userDetail.data.5
    Write-Host "Users email address:" $GLPIGmailAddress

    #process mail using gamxtd....

            $Query = "'subject:#00$GlpiTicketID $GLPITicketSubject'"
            Write-host "Runing query:" $Query

            # get message id from search criteria ...
            Write-Host "$GamDir\gam.exe user $GLPIGmailAddress print messages query $Query"
            Invoke-Expression "$GamDir\gam.exe user $GLPIGmailAddress print messages query $Query"
            Invoke-Expression "$GamDir\gam.exe user $GLPIGmailAddress print messages query $Query" | out-file $tempcsv

            #TODO require null check.... 

            Write-host "importing filtered csv data..."
            $GmailMessage = import-csv $tempcsv | where { ( $_.'Message-ID' -like $MsgIDElement01 -and $_.'Message-ID' -like $MsgIDElement02 -and $_.'Message-ID' -like $MsgIDElement03 ) }

            Write-Host "Raw message:"
            $GmailMessage

            #Getting mail domain from user...
            $SenderDomain = $GmailMessage.user.Split("@")[1]
            Write-host "sender domain:" $SenderDomain

            #remove enclosing chevrons from message id
            $MessageID = $GmailMessage.'Message-ID'.TrimStart("<")
            $MessageID = $MessageID.TrimEnd(">")
            Write-host "removing < and > from message ID - trimmed ID:" $MessageID

            #set message receiver by splitting requester and domain
            $MailReceiver = "$ReceiverPrefix@$senderDomain"
            Write-Host "Setting message receiver to:" $MailReceiver

            $RFCQuery = "'rfc822msgid:$MessageID'" #concatenate message query

            $OriginalSubjectSRC = "$MessageID.Subject"

            #split original ticket title elements...
            $OriginalSubjectSRC = $GmailMessage.Subject

                #ticket title...
                $split01 = $OriginalSubjectSRC -Split $TicketTitleElement01
                Write-Host "Ticket Title:"$split01[1] #after split
                $subject01 = $split01[1]

                #ticket number...
                #$split02 = $split01[0] -split $TicketTitleElement02 -replace("]","") #also remove closing ]
                #Write-Host "Ticket Number:"$split02[1]
                Write-Host "Ticket Number: #00$GlpiTicketID"

                #
                $Subject = "'$subject01'"


            #forward message using RFC message id to new mail receiver...
            Write-Host "$GamDir\gam.exe user $GLPIGmailAddress forward threads to $MailReceiver query $RFCQuery subject $Subject" # logging
            Invoke-expression "$GamDir\gam.exe user $GLPIGmailAddress forward threads to $MailReceiver query $RFCQuery subject $subject"




    ### $GLPITicketID | out-file -Append $tempcsv1 #csv must be encoded as UCS-2 LE BOM
    start-sleep 1
    Write-Host "-----------------------------------------------`n"
    
    }
}

Start-Sleep 10

Write-Host "closing current api session..."
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
