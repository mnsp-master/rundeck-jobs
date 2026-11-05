$mnspver = "0.0.8"


Write-Host $(Get-Date)
Write-Host "MNSP Version" $mnspver

$ScriptName = Split-Path $PSCommandPath -Leaf
Write-Host "Executing Main Rundeck Job..."
Write-Host "MNSP script: $scriptName version: $mnspver"

#Start-Sleep 10
#$ErrorActionPreference="Continue"

# ==============================================================================
# Script: Download, Tag, and Combine Bromcom MIS Reports via GAM
# ==============================================================================


# Ensure working directories exist
@($LogDir, $DataDir) | ForEach-Object {
    if (-not (Test-Path -Path $_)) {
        New-Item -ItemType Directory -Path $_ -Force | Out-Null
    }
}

# --- Logging Setup ---
$now           = Get-Date -Format "yyyyMMdd-HHmmss"  #ENHANCEMENT convert to Rundeck var
$transcriptlog = Join-Path $LogDir "${now}_transcript.log"  #ENHANCEMENT convert to Rundeck var

Start-Transcript -Path $transcriptlog -Force -NoClobber -Append
Clear-Host
Write-Host "MNSP Version: $mnspver" -ForegroundColor Yellow

# --- Clean Working Directory ---
Write-Host "Clearing existing CSV files from: $DataDir" -ForegroundColor Cyan
Remove-Item -Path "$DataDir\*.csv" -Force -Verbose -ErrorAction SilentlyContinue

# List to accumulate rows across all schools in memory
$allRecords    = [System.Collections.Generic.List[object]]::new()
$scriptSuccess = $true

# --- Main Processing Loop ---
try {
    foreach ($school in $schools) {
        $prefix = $school.Prefix
        Write-Host "`n========================================================" -ForegroundColor Cyan
        Write-Host "Processing: $($school.SchoolName) [$prefix]" -ForegroundColor Cyan
        Write-Host "========================================================" -ForegroundColor Cyan

        # Snapshot existing files so we can pinpoint new downloads
        $filesBefore = Get-ChildItem -Path $DataDir -File | Select-Object -ExpandProperty FullName

        # Query targeting only active Inbox emails from the last 24 hours (excluding Trash)
        $query = "label:inbox -in:trash newer_than:1d subject:$prefix*"

        Write-Host "Searching for matching emails..."
        & "$GamDir\gam.exe" user $GWSmail print messages query $query
        if ($LASTEXITCODE -ne 0) {
            throw "GAM failed to query mailbox for prefix '$prefix' (Exit code: $LASTEXITCODE)."
        }

        Write-Host "Downloading attachments..."
        & "$GamDir\gam.exe" user $GWSmail show messages query $query saveattachments targetfolder $DataDir
        if ($LASTEXITCODE -ne 0) {
            throw "GAM failed to download attachments for prefix '$prefix' (Exit code: $LASTEXITCODE)."
        }

        # Identify newly downloaded files
        $filesAfter = Get-ChildItem -Path $DataDir -File | Select-Object -ExpandProperty FullName
        $newFiles   = $filesAfter | Where-Object { $_ -notin $filesBefore }

        # Filter strictly for CSV files
        $newCsvFiles = $newFiles | Where-Object { $_ -like "*.csv" }

        # FATAL CHECK: Did GAM download an attachment for this school?
        if (-not $newCsvFiles) {
            throw "CRITICAL: No CSV attachment found/downloaded for school '$($school.SchoolName)' (Prefix: $prefix). Aborting run to protect Google Sheets data."
        }

        # Guard against duplicate emails: Select only the newest file if multiple were downloaded
        $fileToProcess = $newCsvFiles | Get-Item | Sort-Object LastWriteTime -Descending | Select-Object -First 1
        $file = $fileToProcess.FullName

        # Clean up any older/duplicate CSV attachments downloaded in this iteration
        $newCsvFiles | Where-Object { $_ -ne $file } | ForEach-Object {
            Remove-Item -Path $_ -Force -ErrorAction SilentlyContinue
        }

        Write-Host "Importing file: $(Split-Path $file -Leaf)"
        
        # Import and discard trailing blank lines/summary rows
        $csvData = Import-Csv -Path $file | Where-Object { 
            -not [string]::IsNullOrWhiteSpace($_.'Admission Number') 
        }

        # FATAL CHECK: Does the CSV contain valid data?
        if (-not $csvData -or $csvData.Count -eq 0) {
            throw "CRITICAL: Downloaded CSV for '$prefix' contains 0 valid student records. Aborting run."
        }

        # 1. Clean up Enrolment values using the lookup table
        foreach ($row in $csvData) {
            if ($row.PSObject.Properties[$enrolmentCol]) {
                $rawVal = $row.$enrolmentCol
                if ($enrolmentMap.ContainsKey($rawVal)) {
                    $row.$enrolmentCol = $enrolmentMap[$rawVal]
                }
            }
        }
        
        # 2. Inject School metadata (Institution first, School Name second, followed by original fields)
        $updatedData = $csvData | Select-Object `
            @{Name = 'Institution'; Expression = { $school.Institution }},
            @{Name = 'School Name'; Expression = { $school.SchoolName }},
            *

        # Generate school-prefixed filename
        $cleanLeafName = Split-Path $file -Leaf
        $prefixedName  = if ($cleanLeafName.StartsWith("${prefix}_")) { $cleanLeafName } else { "${prefix}_$cleanLeafName" }
        $targetPath    = Join-Path $DataDir $prefixedName

        # Save the individual school file
        $updatedData | Export-Csv -Path $targetPath -NoTypeInformation -Encoding UTF8
        
        if ($file -ne $targetPath -and (Test-Path $file)) {
            Remove-Item -Path $file -Force
        }

        # Append rows to our central collection
        foreach ($row in $updatedData) {
            $allRecords.Add($row)
        }

        Write-Host " -> Saved to: $prefixedName" -ForegroundColor Green
        Write-Host " -> Added $($updatedData.Count) valid student rows to master collection." -ForegroundColor Green

        # Trash processed emails to keep inbox clean and avoid re-downloading tomorrow
        Write-Host "Trashing processed mailbox message(s)..."
        & "$GamDir\gam.exe" user $GWSmail trash messages query $query 
        #doit temp disable
        if ($LASTEXITCODE -ne 0) {
            Write-Warning "Failed to trash processed email for '$prefix' (Exit code: $LASTEXITCODE). Check mailbox permissions."
        }
    }

    # --- Pre-Upload Integrity Validation ---
    Write-Host "`n========================================================" -ForegroundColor Magenta
    Write-Host "Validating Complete Dataset" -ForegroundColor Magenta
    Write-Host "========================================================" -ForegroundColor Magenta

    $collectedInstitutions = $allRecords.'Institution' | Select-Object -Unique
    if ($collectedInstitutions.Count -ne $schools.Count) {
        throw "CRITICAL: Dataset incomplete! Expected $($schools.Count) schools, but only found $($collectedInstitutions.Count) in memory. Google Sheets upload cancelled."
    }

    # --- Combine into Master File ---
    # 1. Export the master CSV
    $allRecords | Export-Csv -Path $MasterCsvPath -NoTypeInformation -Encoding UTF8

    # 2. Rename specific headers in line 1 using dictionary lookup
    $csvContent = Get-Content -Path $MasterCsvPath -Encoding UTF8
    foreach ($oldHeader in $headerMap.Keys) {
        $csvContent[0] = $csvContent[0] -replace [regex]::Escape($oldHeader), $headerMap[$oldHeader]
    }

    # 3. Write back to file
    $csvContent | Set-Content -Path $MasterCsvPath -Encoding UTF8

    Write-Host "Successfully generated master combined CSV:" -ForegroundColor Green
    Write-Host "Path:        $MasterCsvPath" -ForegroundColor Green
    Write-Host "Total Rows:  $($allRecords.Count)" -ForegroundColor Green

    # 4. Upload to Google Sheets
    Write-Host "`nUploading $($allRecords.Count) records to Google Sheet tab ($TabID)..." -ForegroundColor Cyan
    & "$GamDir\gam.exe" user $GamUser update drivefile $GoogleSheetID csvsheet id:$TabID localfile $MasterCsvPath retainname

    if ($LASTEXITCODE -ne 0) {
        throw "GAM failed to upload CSV to Google Sheets (Exit code: $LASTEXITCODE)."
    }

    Write-Host "Google Sheet tab updated successfully." -ForegroundColor Green
}
catch {
    $scriptSuccess = $false
    Write-Host "`n[FATAL ERROR] Pipeline aborted: $_" -ForegroundColor Red
}
finally {
    Write-Host "`nScript run finished. Stopping transcript." -ForegroundColor Cyan
    Stop-Transcript
}

# Return non-zero exit code to Windows Task Scheduler if an error occurred
if (-not $scriptSuccess) {
    exit 1
}

$RundeckJobOutput = "Testing..."

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
