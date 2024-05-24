$mnspver = "0.0.15"

Write-Host $(Get-Date)
Write-Host "MNSP Version" $mnspver

$ErrorActionPreference="Continue"
Set-Location $GamDir

#Get/Confirm Google instance
Invoke-Expression "$GamDir\gam.exe info domain" 

Write-Host "get members of google group of previously processed users..."
if (Test-Path -path $GoogleGroupUsersCSV ) {

        Write-Host "$GoogleGroupUsersCSV exists, deleting..."
        Remove-Item -Path $GoogleGroupUsersCSV -Force
    }
    Invoke-Expression "$GamDir\gam.exe gam print group-members group_ns $GoogleGroup > $GoogleGroupUsersCSV" -ErrorAction SilentlyContinue


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

foreach ($SrcUser in $userSource) {

$email = $($SrcUser.email)
$FirstName = $($SrcUser.FirstName)
$LastName = $($SrcUser.LastName)

$AppParams = -join ("&users[0][email]=",$email,"&users[0][firstname]=",$FirstName,"&users[0][lastname]=",$LastName,"&users[0][password]=",$Password)
$AppFullURL = -join ($AppURL,$AppFunction,$APIToken,$AppParams) #create full rest api url

Write-Host "Full App URL..."
$AppFullURL

#create user
#$userResult = Invoke-RestMethod $AppFullURL #compose restapi
#$userResult 

#Write-Host "User creation response..."
#$userResult.RESPONSE.MULTIPLE.SINGLE.KEY #create user response

}
