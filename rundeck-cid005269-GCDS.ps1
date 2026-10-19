$mnspver = "0.0.3"

Write-Host $(Get-Date)
Write-Host "MNSP Version" $mnspver

$ScriptName = Split-Path $PSCommandPath -Leaf
Write-Host "Executing Main Rundeck Job..."
Write-Host "MNSP script: $scriptName version: $mnspver"

#Start-Sleep 10
#$ErrorActionPreference="Continue"

$CustomerEmailAddress

Write-Host "Executing GCDS..."
Write-Host "$GCDSappPath $GCDSparams"

<#

$password = Get-NewPassword
#Write-Host "Confirm Password function: $password"


#Convert the plain text password into a SecureString
$SecurePassword = $password | ConvertTo-SecureString -AsPlainText -Force

#Find the user by their email address and reset the password
try {
    $ADuser = Get-ADUser -server $DC -Filter "mail -eq '$CustomerEmailaddress'"

    if ($null -eq $ADuser) {
        throw "No user found with email address: '$CustomerEmailaddress'"
    }

    $ADuser | Set-ADAccountPassword -server $DC -NewPassword $SecurePassword -Reset #-whatif
    $JobOutput01 = "Successfully reset password for: $CustomerEmailaddress to: $password"
    #generate GWS MFA backup codes
    $RawGWSbackupCodes = (& $GamDir\gam.exe user $CustomerEmailAddress update backupcodes) #| select-object -Skip 1
    $backupCodes = [regex]::Matches(($RawGWSbackupCodes -join ' '), '\b\d{8}\b').Value
    $CodesFormatted = $backupcodes -Join ' '

} catch {
    $JobOutput01 = "Failed to set password for user with email address: '$CustomerEmailaddress' Error: $($_.Exception.Message)"
}

$RundeckJobOutput = @"
$jobOutput01
 | Backup Codes:
$CodesFormatted
"@

DashedLine

#>