Clear-host

$CID = "C01090"
$DataDir = "D:\AppData\Rundeck\$CID\Data"
$LogDir = "D:\AppData\Rundeck\$CID\Logs"
$transcriptlog = "$LogDir\$(Get-date -Format yyyyMMdd-HHmmss)_transcript.log"

Start-Transcript -Path $transcriptlog -Force -NoClobber -Append

Write-Host $(Get-Date)

$ErrorActionPreference="Stop"

#Set-Location $GamDir

#Parameters
$Path = "D:\Microsoft SQL Server\MSSQL12.SDSSIMS\MSSQL\Backup" # Path where the file is located 
$Days = "15" # Number of days before current date
$FileNamePatternMatch = "PARS_*.BAK"
 
#Get All Files modified more than the last n days
Get-ChildItem -path "$Path\PARS_*.BAK" -Recurse -File | Where CreationTime -lt  (Get-Date).AddDays(-$Days) | Remove-Item -verbose

#remove all log files older than 30 days
Get-ChildItem "$logdir\*_transcript.log" -Recurse -File | Where CreationTime -lt  (Get-Date).AddDays(-30) | Remove-Item -verbose

Stop-Transcript
