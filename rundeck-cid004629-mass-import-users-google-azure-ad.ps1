$mnspver = "0.0.22"

Function Get-NewPassword {
    $PwdUrl = $MNSPgetPasswordPRL

    #failsafe password
    $pwdFailsafe = $MNSPgetPasswordPRLfailsafe

    
    try {
        Write-Host "Attempting to retrieve password from web service..."
        $response = Invoke-WebRequest -Uri $PwdUrl -UseBasicParsing

        # Extract the content from the response object and store it in a variable.
        $password = $response.Content

        # Return the generated password.
        return $password
    }
    catch {
        # If the web request fails, a detailed error message is logged.
        Write-Error "Failed to retrieve password from $PwdUrl. Using failsafe password instead." -ErrorAction Stop

        # Return the failsafe password as the function's output.
        return $pwdFailsafe
    }
}



Write-Host $(Get-Date)
Write-Host "MNSP Version" $mnspver

$ScriptName = Split-Path $PSCommandPath -Leaf
Write-Host "Executing Main Rundeck Job..."
Write-Host "MNSP script: $scriptName version: $mnspver"

#Start-Sleep 10
$ErrorActionPreference="Continue"
Set-Location $GamDir

#Set Google Instance: Destination...
Write-Host "###### Set Google instance: Destination... ######"

Write-Host "Setting workspace Destination: $GoogleWorkSpaceDestination"
Invoke-Expression "$GamDir\gam.exe select $GoogleWorkSpaceDestination save" # swap/set google workspace
Invoke-Expression "$GamDir\gam.exe info domain"
start-sleep 3

Write-Host "School prefix: $SchoolCode"

Write-Host "Getting all OUs from supplied source: $GoogleSourceBaseOU"
$SourceGoogleOus = @()
Invoke-Expression "$GamDir\Gam.exe print orgs fromparent '$GoogleSourceBaseOU' > $tempcsv1"

$GoogleSourceOUs = import-csv -path $tempcsv1

<#
foreach ($SourceGoogleOU in $GoogleSourceOUs) {
    Write-Host "Processing: $($SourceGoogleOU)"
}
#>

Write-Host "Getting all users from supplied Source: $GoogleSourceBaseOU"
Invoke-Expression "$GamDir\Gam.exe ou_and_children '$GoogleSourceBaseOU' print fields primaryEmail,firstname,lastname,displayname,orgUnitPath, custom MNSP.adminNumber > $tempcsv2"

$UsersToProcess = import-csv -path $tempcsv2
Write-host "Number of users to process: " $UsersToProcess.count


$password = Get-NewPassword
Write-Host "Confirm Password function: $password"

DashedLine

