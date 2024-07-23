$mnspver = "0.0.34"

Function GeneratePWD {
#Write-Host $i

# Generate a random code or password
$code = ""
$codeLength = 24
$allowedChars = "ABCDEFGHJKMNPQRSTWXYZabcdefghmnpqrstwxyz123456789"
$rng = new-object System.Security.Cryptography.RNGCryptoServiceProvider
$randomBytes = new-object "System.Byte[]" 1
# keep unbiased by making sure input range divides evenly by output range
$inputRange = $allowedChars.Length * [Math]::Floor(256 / $allowedChars.Length)
while($code.Length -lt $codeLength) {
    $rng.GetBytes($randomBytes)
    $byte = $randomBytes[0]
    if($byte -lt $inputRange) { # throw away out-of-range inputs
        $code += $allowedChars[$byte % $allowedChars.Length]
    }
}

$pwd = $code + "!"
$pwd

}

Write-Host $(Get-Date)
Write-Host "MNSP Version" $mnspver

$ErrorActionPreference="Continue"
Set-Location $GamDir

#Get/Confirm Google instance
Invoke-Expression "$GamDir\gam.exe info domain" 

Write-Host "get members of google group: $GoogleGroup of previously processed users..."
if (Test-Path -path $GoogleGroupUsersCSV ) {

        Write-Host "$GoogleGroupUsersCSV exists, deleting..."
        Remove-Item -Path $GoogleGroupUsersCSV -Force
    }
    Start-sleep 5
    Invoke-Expression "$GamDir\gam.exe print group-members group_ns $GoogleGroup > $GoogleGroupUsersCSV" -ErrorAction SilentlyContinue
$PreviouslyProcessedUsers = import-csv $GoogleGroupUsersCSV

#download gsheet
Write-Host "Downloading Googlesheet containing all required users..."

    if (Test-Path -path $SrcUserDataCSV ) {

        Write-Host "$SrcUserDataCSV exists, deleting..."
        Remove-Item -Path $SrcUserDataCSV -Force
    }

    Set-Location $GamDir

    Invoke-Expression "$GamDir\gam.exe user $GoogleGamMail get drivefile id $GsheetUserSourceID format csv targetfolder $datadir" -ErrorAction SilentlyContinue
Start-Sleep 2
$userSource = Import-Csv -Path $SrcUserDataCSV

#compare required accounts with previously processed accounts
foreach ($SrcUser in $userSource) {

    #additional condition check skip if supplied email value is empty
    if ($email -ne $NoMailCheck ) {

    $email = $($SrcUser.email)
    Write-Host "looking for email: $email"
    if ($PreviouslyProcessedUsers.email.Contains($email)) { 
    Write-Host "$email already processed skipping..." } else {

        #generate random password
        $password = $(GeneratePWD)

        #$email = $($SrcUser.email)
        $FirstName = $($SrcUser.FirstName)
        $LastName = $($SrcUser.LastName)

        $AppParams = -join ("&users[0][email]=",$email,"&users[0][firstname]=",$FirstName,"&users[0][lastname]=",$LastName,"&users[0][password]=",$Password)
        $AppFullURL = -join ($AppURL,$AppFunction,$APIToken,$AppParams) #create full rest api url

        Write-Host "Full App URL..."
        Write-host $AppFullURL

            if (!password) { #null password check
                Write-Warning "Password empty as of $(get-date)"
            } else {

                    #<#
                    #create user using api
                    $userResult = Invoke-RestMethod $AppFullURL #compose restapi
                    $userResult 

                    Write-Host "User creation response..."
                    $userResult.RESPONSE.MULTIPLE.SINGLE.KEY #create user response

                    Write-host "Adding user to Google Group: $GoogleGroup"
                    Invoke-Expression "$GamDir\gam.exe update group $GoogleGroup add member $email" -ErrorAction SilentlyContinue
                    #>

                    #send notification email to service Supplier - still needs work attaching csv.
                    #gam sendemail auser@domain from noreply@domain subject "test" message "test message"
            }

        }
    } else {
        Write-warning "No email attribute avaialble for $FirstName $LastName skipping..."
    }
}



#Write-Host "User creation response..."
#$userResult.RESPONSE.MULTIPLE.SINGLE.KEY #create user response

#}

<#

foreach ($SrcUser in $userSource) {

$email = $($SrcUser.email)
$FirstName = $($SrcUser.FirstName)
$LastName = $($SrcUser.LastName)

$AppParams = -join ("&users[0][email]=",$email,"&users[0][firstname]=",$FirstName,"&users[0][lastname]=",$LastName,"&users[0][password]=",$Password)
$AppFullURL = -join ($AppURL,$AppFunction,$APIToken,$AppParams) #create full rest api url

Write-Host "Full App URL..."
$AppFullURL


Function GeneratePwd {
# Generate random code or password
$code = ""
$codeLength = 24
$allowedChars = "ABCDEFGHJKMNPQRSTWXYZabcdefghmnpqrstwxyz123456789"
$rng = new-object System.Security.Cryptography.RNGCryptoServiceProvider
$randomBytes = new-object "System.Byte[]" 1
# keep unbiased by making sure input range divides evenly by output range
$inputRange = $allowedChars.Length * [Math]::Floor(256 / $allowedChars.Length)
while($code.Length -lt $codeLength) {
    $rng.GetBytes($randomBytes)
    $byte = $randomBytes[0]
    if($byte -lt $inputRange) { # throw away out-of-range inputs
        $code += $allowedChars[$byte % $allowedChars.Length]
    }
}

$pwd = $code + "!"

}


#>