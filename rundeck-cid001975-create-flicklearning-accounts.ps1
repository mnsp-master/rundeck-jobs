$mnspver = "0.0.5"

Write-Host $(Get-Date)
Write-Host "MNSP Version" $mnspver

$ErrorActionPreference="Continue"
Set-Location $GamDir

#Get/Confirm Google instance
Invoke-Expression "$GamDir\gam.exe info domain" 


$AppParams = -join ("&users[0][email]=",$email,"&users[0][firstname]=",$FirstName,"&users[0][lastname]=",$LastName,"&users[0][password]=",$Password)
$AppFullURL = -join ($AppURL,$AppFunction,$APIToken,$AppParams) #create full rest api url

Write-Host "Full App URL..."
$AppFullURL

#create user
$userResult = Invoke-RestMethod $AppFullURL #compose restapi
$userResult 

Write-Host "User creation response..."
$userResult.RESPONSE.MULTIPLE.SINGLE.KEY #create user response

