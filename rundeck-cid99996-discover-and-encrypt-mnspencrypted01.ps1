$mnspver = "3.2.7"


Write-Host $(Get-Date)
Write-Host "MNSP Version" $mnspver

$ScriptName = Split-Path $PSCommandPath -Leaf
Write-Host "Executing Main Rundeck Job..."
Write-Host "MNSP script: $scriptName version: $mnspver"

#Start-Sleep 10
#$ErrorActionPreference="Continue"

# Ensure working directories exist
@($LogDir, $DataDir) | ForEach-Object {
    if (-not (Test-Path -Path $_)) {
        New-Item -ItemType Directory -Path $_ -Force | Out-Null
    }
}

# --- Logging Setup ---
#$now           = Get-Date -Format "yyyyMMdd-HHmmss"  #ENHANCEMENT convert to Rundeck var
#$transcriptlog = Join-Path $LogDir "${now}_transcript.log"  #ENHANCEMENT convert to Rundeck var

#Start-Transcript -Path $transcriptlog -Force -NoClobber -Append
Clear-Host
Write-Host "MNSP Version: $mnspver" -ForegroundColor Yellow

# Convert the incoming string safely into a real boolean variable
#$mnspRun = [System.Convert]::ToBoolean($toggle)
#$mnspRun = $false # Uncomment to force dry run testing interactively

function GoogleSheetEncryptedPassword {
    param (
        [string]$Plaintext,
        [string]$Key
    )

    if ([string]::IsNullOrEmpty($Plaintext) -or [string]::IsNullOrEmpty($Key)) { return "" }

    $SecretSeedBytes = [System.Text.Encoding]::UTF8.GetBytes("secure_seed")
    $KeyBytes = [System.Text.Encoding]::UTF8.GetBytes($Key)
    $Hmac = New-Object System.Security.Cryptography.HMACSHA256
    $Hmac.Key = $SecretSeedBytes
    $SaltBytes = $Hmac.ComputeHash($KeyBytes)

    $ProcessedSalt = foreach ($b in $SaltBytes) {
        if ($b -gt 127) { [int]$b - 256 } else { [int]$b }
    }

    $OutputChars = [System.Collections.Generic.List[char]]::new()

    for ($i = 0; $i -lt $Plaintext.Length; $i++) {
        $c = [int]$Plaintext[$i]
        $SaltMod = [Math]::Abs($ProcessedSalt[$i % $ProcessedSalt.Count]) % 256
        $k = [int]$Key[$i % $Key.Length]
        $FinalCharValue = ($c -bxor $k -bxor $SaltMod) + 2000
        $OutputChars.Add([char]$FinalCharValue)
    }

    $JoinedString = -join $OutputChars
    $FinalBytes = [System.Text.Encoding]::UTF8.GetBytes($JoinedString)
    return [System.Convert]::ToBase64String($FinalBytes)
}

function Get-NewPassword {
    $PwdUrl = $MNSPgetPasswordURL
    $pwdFailsafe = $MNSPgetPasswordPRLfailsafe

    try {
        $response = Invoke-WebRequest -Uri $PwdUrl -UseBasicParsing -ErrorAction Stop
        return $response.Content.Trim()
    }
    catch {
        Write-Warning "Failed to retrieve password from $PwdUrl. Using failsafe password instead."
        return $pwdFailsafe
    }
}

# ==============================================================================
# PHASE 1: DOWNLOAD GOOGLE SHEET (SOURCE OF TRUTH)
# ==============================================================================
$FullPath      = Join-Path $DataDir $TempCSV

Write-Host "`n--- PHASE 1: DOWNLOADING GOOGLE SHEET ---" -ForegroundColor Cyan
if (Test-Path $FullPath) { Remove-Item $FullPath -Force -Verbose }

& "$GamDir\gam.exe" user $GamUser get drivefile $GoogleSheetID gsheet id:$TabID format csv targetfolder $DataDir targetname $TempCSV

if (-not (Test-Path $FullPath)) {
    Write-Error "Failed to download CSV from Google Drive." -ErrorAction Stop
}

$rawCsv = @(Import-Csv -Path $FullPath)

# Identify headers: Col 1 (Admin# / UPN) and Col 2 (EncryptedPassword / Col B)
$csvHeaders = if ($rawCsv.Count -gt 0) {
    $rawCsv[0].PSObject.Properties.Name
} else {
    ((Get-Content -Path $FullPath -TotalCount 1) -split ',') -replace '"',''
}

$Col1Name = if ($csvHeaders -contains "UPN") { "UPN" } else { $csvHeaders[0] }
$Col2Name = if ($csvHeaders -contains "EncryptedPassword") { "EncryptedPassword" } else { $csvHeaders[1] }

Write-Host "  -> Column 1 (Admin # / UPN): '$Col1Name'" -ForegroundColor DarkGray
Write-Host "  -> Column 2 (Trigger / Encrypted Password): '$Col2Name'" -ForegroundColor DarkGray

# Expandable list of rows
$DataList = [System.Collections.Generic.List[PSCustomObject]]::new()
foreach ($item in $rawCsv) {
    $DataList.Add($item)
}

# Index all admin numbers present in the Google Sheet
$SheetAdminNumbers = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
foreach ($r in $DataList) {
    $val = $r.$Col1Name
    if (-not [string]::IsNullOrWhiteSpace($val)) {
        [void]$SheetAdminNumbers.Add($val.Trim())
    }
}
Write-Host "Loaded $($DataList.Count) rows from Google Sheet ($($SheetAdminNumbers.Count) unique ID records)." -ForegroundColor Green

# ==============================================================================
# PHASE 2: AUDIT AD & RECONCILE MISSING USERS
# ==============================================================================

Write-Host "`n--- PHASE 2: AUDITING ACTIVE DIRECTORY & DETECTING MISSING USERS ---" -ForegroundColor Cyan

$ADprops = @(
    "Name", "displayName", "SamAccountName", "mail", "userPrincipalName", 
    "distinguishedName", "objectGUID", "mnspAdminNumber", "mnspEncrypted01", 
    "enabled", "company", "employeeID", "employeeType", "whenCreated", "whenChanged"
)

$NewlyAppendedCount = 0
$ouIndex = 0

foreach ($TargetOU in $TargetOUs) {
    $ouIndex++
    $ouName = ($TargetOU -split ',')[0] -replace '^OU=', ''

    Write-Progress -Activity "Auditing AD Users ($MISname)" `
                   -Status "Checking ($ouIndex/$($TargetOUs.Count)): $ouName" `
                   -PercentComplete (($ouIndex / $TargetOUs.Count) * 100)

    # Gather users whose mnspAdminNumber does not contain an underscore
    $Users = @(Get-ADUser -Filter 'mnspAdminNumber -notlike "*_*"' -SearchBase $TargetOU -SearchScope Subtree -Properties $ADprops)

    foreach ($adUser in $Users) {
        if ([string]::IsNullOrWhiteSpace($adUser.mnspAdminNumber)) { continue }
        $adminNum = $adUser.mnspAdminNumber.Trim()

        # If user exists in AD but is NOT in the Google Sheet at all:
        if (-not $SheetAdminNumbers.Contains($adminNum)) {
            Write-Host "  [+] AD user missing from Google Sheet: $($adUser.Name) ($adminNum). Appending to sheet..." -ForegroundColor Yellow
            
            # Leave Column 2 blank so Phase 3 generates their password and updates both AD and Google Sheets
            $newRow = [ordered]@{}
            foreach ($col in $csvHeaders) {
                switch ($col.Trim()) {
                    $Col1Name           { $newRow[$col] = $adminNum; break }
                    $Col2Name           { $newRow[$col] = ""; break }
                    "Name"              { $newRow[$col] = $adUser.Name; break }
                    "displayName"       { $newRow[$col] = $adUser.displayName; break }
                    "userPrincipalName" { $newRow[$col] = $adUser.userPrincipalName; break }
                    "mail"              { $newRow[$col] = $adUser.mail; break }
                    "SamAccountName"    { $newRow[$col] = $adUser.SamAccountName; break }
                    default             { $newRow[$col] = "" }
                }
            }

            $DataList.Add([PSCustomObject]$newRow)
            [void]$SheetAdminNumbers.Add($adminNum)
            $NewlyAppendedCount++
        }
    }
}

Write-Progress -Activity "Auditing AD Users ($MISname)" -Completed

if ($NewlyAppendedCount -gt 0) {
    Write-Host "Appended $NewlyAppendedCount new user(s) discovered in AD into the dataset." -ForegroundColor Green
} else {
    Write-Host "All evaluated AD users already exist in the Google Sheet." -ForegroundColor Green
}

# ==============================================================================
# PHASE 3: PROCESS ROWS WHERE COL 1 EXISTS AND COL 2 IS BLANK
# ==============================================================================

Write-Host "`n--- PHASE 3: PROCESSING MANAGED PASSWORD RESETS ---" -ForegroundColor Cyan
$UpdatedCount = 0

foreach ($Row in $DataList) {
    $UpnVal           = $Row.$Col1Name
    $ExistingPassword = $Row.$Col2Name

    # MASTER TRIGGER: If Col 1 has a value, but Col 2 is empty/blank/cleared
    # This fires regardless of whether AD's mnspEncrypted01 is currently populated
    if ([string]::IsNullOrWhiteSpace($UpnVal) -or -not [string]::IsNullOrWhiteSpace($ExistingPassword)) {
        continue
    }

    $UpnVal = $UpnVal.Trim()
    Write-Host "`n[!] Triggering password update for: $UpnVal (Column 2 is blank)" -ForegroundColor Yellow

    # Match AD User
    $userToUpdate = Get-ADUser -Filter "mnspAdminNumber -eq '$UpnVal'" -Properties $ADprops

    if (-not $userToUpdate) {
        Write-Warning "  -> AD User with mnspAdminNumber '$UpnVal' not found in Active Directory. Skipping."
        continue
    }

    # Verify OU membership rule
    $isBeaconUser = [bool]($BeaconOUs | Where-Object { $userToUpdate.DistinguishedName -like "*,$_" })

    if ($MISname -eq "Arbor (MNSP)" -and $isBeaconUser) {
        Write-Host "  -> Skipping Beacon OU user '$($userToUpdate.SamAccountName)' (Run mode: $MISname)." -ForegroundColor DarkGray
        continue
    }

    if ($MISname -eq "Bromcom (Beacon)" -and -not $isBeaconUser) {
        Write-Host "  -> Skipping non-Beacon OU user '$($userToUpdate.SamAccountName)' (Run mode: $MISname)." -ForegroundColor DarkGray
        continue
    }

    # Informative log if AD already had an encrypted password being overwritten
    if (-not [string]::IsNullOrWhiteSpace($userToUpdate.mnspEncrypted01)) {
        Write-Host "  -> Overriding existing AD mnspEncrypted01 attribute as requested by empty Column 2." -ForegroundColor Cyan
    }

    Write-Host "  -> Generating new password for: $($userToUpdate.Name) ($($userToUpdate.SamAccountName))" -ForegroundColor Cyan

    # Retry loop with AD password policy validation
    $pwdIsValid   = $false
    $maxAttempts  = 5
    $attemptCount = 0

    while (-not $pwdIsValid -and $attemptCount -lt $maxAttempts) { 
        $attemptCount++
        $PlainPassword = Get-NewPassword 
        
        # Split password safely
        $PasswordParts = $PlainPassword.Split('&')
        $PwdCheck01    = $PasswordParts[0]
        $PwdCheck02    = if ($PasswordParts.Count -gt 1) { $PasswordParts[1] -replace '\d+','' } else { "" }

        # Name similarity pre-check
        if ($userToUpdate.Name -like "*$PwdCheck01*" -or ($PwdCheck02 -and $userToUpdate.Name -like "*$PwdCheck02*")) { 
            Write-Warning "  -> '$($userToUpdate.Name)' and '$PlainPassword' are too similar. Requesting another..."
            Start-Sleep -Seconds 2
            continue
        }

        # Attempt setting the password in AD
        $SecurePassword = ConvertTo-SecureString -AsPlainText $PlainPassword -Force
        try {
            Write-Host "  -> Setting AD password (Attempt $attemptCount of $maxAttempts)..."
            # Set-ADAccountPassword -Identity $userToUpdate.objectGuid -NewPassword $SecurePassword -Reset -ErrorAction Stop
            
            Write-Host "  -> AD password updated successfully!" -ForegroundColor Green
            $pwdIsValid = $true
        }
        catch {
            Write-Warning "  -> AD rejected password for '$($userToUpdate.Name)': $($_.Exception.Message)"
            Write-Warning "  -> Retrying with a new password..."
            Start-Sleep -Seconds 2
        }
    }

    if (-not $pwdIsValid) {
        Write-Error "  -> Exceeded maximum password attempts ($maxAttempts) for '$($userToUpdate.Name)'. Skipping."
        continue
    }

    # Encrypt password for Google Sheet
    $EncryptedOutput = GoogleSheetEncryptedPassword -Plaintext $PlainPassword -Key $Key

    # Update mnspEncrypted01 attribute in Active Directory (overwrites previous value)
    Write-Host "  -> Overwriting mnspEncrypted01 attribute in AD"
    # Set-ADUser -Identity $userToUpdate.objectGuid -Replace @{ mnspEncrypted01 = $EncryptedOutput }

    Start-Sleep -Seconds 1 # TID60279 throttle
    
    # Update Column 2 in the local dataset
    $Row.$Col2Name = $EncryptedOutput
    $UpdatedCount++
}

# ==============================================================================
# PHASE 4: UPLOAD UPDATED DATASET BACK TO GOOGLE SHEETS
# ==============================================================================

Write-Host "`n--- PHASE 4: UPLOADING TO GOOGLE SHEETS ---" -ForegroundColor Cyan

if ($UpdatedCount -gt 0 -or $NewlyAppendedCount -gt 0) {
    # Export full dataset back to CSV
    $DataList | Export-Csv -Path $FullPath -NoTypeInformation -Force

    Write-Host "Uploading back to Google Sheets ($UpdatedCount password(s) generated/updated, $NewlyAppendedCount new user(s) added)..." -ForegroundColor Cyan
    & "$GamDir\gam.exe" user $GamUser update drivefile $GoogleSheetID localfile $FullPath retainname gsheet id:$TabID
    
    $RundeckJobOutput = "Success! Google Sheet ($MISname) updated successfully." 
    Write-Host "Success! Google Sheet ($MISname) updated successfully." -ForegroundColor Green
} else {
    Write-Host "Google Sheet ($MISname) is already up to date. No upload required." -ForegroundColor Yellow
    $RundeckJobOutput = "Google Sheet ($MISname) is already up to date. No upload required."
}

# Cleanup
if (Test-Path $FullPath) { Remove-Item $FullPath -Force }

Write-Host "`nSync run completed." -ForegroundColor Cyan
Stop-Transcript

DashedLine

<#
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
#>
