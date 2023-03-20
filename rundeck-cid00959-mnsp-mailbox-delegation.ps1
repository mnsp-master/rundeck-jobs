cls

$GamDir = "D:\AppData\GAMXTD3\app"

#$WorkDir="D:\AppData\GAM"
$DataDir="D:\AppData\GAM\Data\C1417"
#$GamDir="$WorkDir"
$log = "$DataDir\transcript.log"
#$tempcsv="$DataDir\temp.csv"
#$tempcsv2="$DataDir\temp2.csv"
#$GoogleDocID="1nHwDxp9d72P-ZMbe1U8LGI2jp6admDcmZcH2a7j3uv4"

Start-Transcript -Path $log -Force -NoClobber -Append

Write-Host $(Get-Date)

$ErrorActionPreference="Stop"

Set-Location $GamDir

#Invoke-Expression "$GamDir\gam.exe info domain"

$action=@() #clear action variable
$trustee=@() #clear trustee variable
$receiver=@() #clear user receiving delegation right variable

#convert inputs into ps vars:
$action = "@option.GmailUserAction@"
$trustee = "@option.GmailToBeDelegated@"
$receiver = "@option.GmailAddress@"

Write-Host "Passed vars..."
$action
$trustee
$receiver

if ($action -eq "delegate") {
    #Write-Host "Granting @option.GmailAddress@ trustee asignment of @option.GmailToBeDelegated@"
    if ($receiver) {
    Invoke-Expression "$GamDir\gam.exe user $trustee @option.GmailUserAction@ to $receiver"
    #Write-Host "$GamDir\gam.exe user $trustee @option.GmailUserAction@ to $receiver"
    }
    else {
        Write-Host "OOPS - Please supply an email address as a delegate....`n"
        }
    }

elseif ($action -eq "delete") {
    #Write-Host "removing @option.GmailAddress@ as a trustee of @option.GmailToBeDelegated@"
    if ($receiver) {
    Invoke-Expression "$GamDir\gam.exe user $trustee @option.GmailUserAction@ delegate $receiver"
    #write-host "$GamDir\gam.exe user $trustee @option.GmailUserAction@ delegate $receiver"
        
    }
    else {
        Write-Host "OOPS - Please supply an email address from which to delete delegation....`n"
        }
    }
    
    Write-Host "----------------------------------------------------------"
    Write-Host "Delegates for email address: @option.GmailToBeDelegated@"
    #Invoke-Expression "$GamDir\gam.exe user @option.GmailToBeDelegated@ @option.GmailUserAction@ delegates"
    Invoke-Expression "$GamDir\gam.exe user $trustee print delegates"
    write-host "$GamDir\gam.exe user $trustee print delegates"

Stop-Transcript
