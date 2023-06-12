Clear-host

$mnspVer = "1.0.0.0.7"
$root = "D:" # drive letter root to host scripts/logs etc
$MSSQLroot = "D:" # drive letter root of MSSQL data directory

$CID = "C01090"
$DataDir = "$root\AppData\Rundeck\$CID\Data"
$LogDir = "$root\AppData\Rundeck\$CID\Logs"
$transcriptlog = "$LogDir\$(Get-date -Format yyyyMMdd-HHmmss)_transcript.log"

Start-Transcript -Path $transcriptlog -Force -NoClobber -Append

Write-Host $(Get-Date)

$ErrorActionPreference="Stop"

#Parameters
$Path = "$MSSQLroot\Microsoft SQL Server\MSSQL12.SDSSIMS\MSSQL\Backup" # Path where file(s) are located 
$Days = "14" # Number of days before current date
$FileNamePatternMatch = "PARS_*.BAK" # filename pattern to match for deletion
 
#remove all Files created older than the last n days
Get-ChildItem -path "$Path\$FileNamePatternMatch" -Recurse -File | Where CreationTime -lt  (Get-Date).AddDays(-$Days) | Remove-Item -verbose

#remove all log files older than 30 days
Get-ChildItem "$logdir\*_transcript.log" -Recurse -File | Where CreationTime -lt  (Get-Date).AddDays(-30) | Remove-Item -verbose

Stop-Transcript
