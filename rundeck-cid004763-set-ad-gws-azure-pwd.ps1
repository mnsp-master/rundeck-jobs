$mnspver = "0.0.9"

Function Get-NewPassword {
    $PwdUrl = $MNSPgetPasswordURL

    #failsafe password
    $pwdFailsafe = $MNSPgetPasswordURLfailsafe

    
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
#$ErrorActionPreference="Continue"

$CustomerEmailAddress

$password = Get-NewPassword
#Write-Host "Confirm Password function: $password"
$RundeckJobOutput = "New User's email: $CustomerEmailaddress password: $password"

# 1. Convert the plain text password into a SecureString
$SecurePassword = $password | ConvertTo-SecureString -AsPlainText -Force

# 2. Find the user by their email address and reset the password
Get-ADUser -Filter "mail -eq '$CustomerEmailaddress'" | Set-ADAccountPassword -NewPassword $SecurePassword -Reset -whatif

DashedLine

