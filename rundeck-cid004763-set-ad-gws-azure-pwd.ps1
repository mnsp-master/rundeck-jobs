$mnspver = "0.0.4"

Function Get-NewPassword {
    $PwdUrl = $MNSPgetPasswordURL

    #failsafe password
    $pwdFailsafe = $MNSPgetPasswordURLfailsafe

    
    try {
        Write-Host "Attempting to retrieve password from web service..."
        $response = Invoke-WebRequest -Uri $PwdUrl -UseBasicParsing

        # Extract the content from the response object and store it in a variable.
        $newPassword = $response.Content

        # Return the generated password.
        return $newPassword
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
#$ErrorActionPreference="Continue"


$password = Get-NewPassword
Write-Host "Confirm Password function: $newPassword"

DashedLine

