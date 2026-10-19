$mnspver = "0.0.19"

Write-Host $(Get-Date)
Write-Host "MNSP Version" $mnspver

$ScriptName = Split-Path $PSCommandPath -Leaf
Write-Host "Executing Main Rundeck Job..."
Write-Host "MNSP script: $scriptName version: $mnspver"

#Start-Sleep 10
#$ErrorActionPreference="Continue"

#$CustomerEmailAddress

$RundeckJobOutput = "Processed CLI: $GCDSappPath $GCDSparams"
Write-Output "Windows actual execution user: $(([Security.Principal.WindowsIdentity]::GetCurrent()).Name)"

# mode set from rundeck job option
if ($mode -eq "PRODUCTION") {
    Write-Output "WARNING: Launching GCDS in PRODUCTION mode."
    $GCDSargs = @("-a", "-f", "-o", "-c", $GCDSconfigXML)
} else {
    Write-Output "INFO: Launching GCDS in DRY RUN mode."
    $GCDSargs = @("-o", "-c", $GCDSconfigXML)
}

# 4. Execute the command natively passing the argument array
Write-Output "Executing GCDS command with arguments: $GCDSargs"
& $GCDSappPath $GCDSargs

start-sleep 2

# get latest GCDS log file
$latestLog = Get-ChildItem -Path $GCDSlogFolder -File | 
             Where-Object { $_.LastWriteTime -gt (Get-Date).AddMinutes(-5) } | 
             Sort-Object LastWriteTime -Descending | 
             Select-Object -First 1

$logLines = Get-Content -Path $latestLog.FullName
    # Match EITHER "Proposed Change Report" OR "Change Status Report"
    $matchedLine = $logLines | Select-String -Pattern "(Proposed Change Report|Change Status Report)" | Select-Object -First 1


if ($matchedLine) {
        # Extract the line number where the summary starts (0-indexed adjustment)
        $startIndex = $matchedLine.LineNumber - 1
        
        # Output everything from that line onwards
        $logLines[$startIndex..($logLines.Count - 1)] | Out-String | Write-Output
    } else {
        Write-Warning "Could not find 'Proposed Change Report' marker. Dumping last 50 lines instead:"
        $logLines[-50..-1] | Out-String | Write-Output
    }
    
    Write-Output "======================================================================"
    Write-Output "                END OF INJECTED GCDS SUMMARY REPORT                   "
    Write-Output "======================================================================"
    Write-Output "`n"
} else {
    Write-Warning "Automation Warning: Could not locate a fresh GCDS log file in \$logFolder to inject into the transcript."
}
