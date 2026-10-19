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


