cls

$CID="C01090"
$DataDir="D:\AppData\Rundeck\$CID\Data"
$LogDir="D:\AppData\Rundeck\$CID\Logs"
$transcriptlog = "$LogDir\$(Get-date -Format yyyyMMdd-HHmmss)_transcript.log"

Start-Transcript -Path $transcriptlog -Force -NoClobber -Append

Write-Host $(Get-Date)

$ErrorActionPreference="Stop"

#Set-Location $GamDir

#Parameters
$Path = "D:\Microsoft SQL Server\MSSQL12.SDSSIMS\MSSQL\Backup" # Path where the file is located 
$Days = "30" # Number of days before current date
$FileNamePatternMatch = "PARS_*.BAK"
 
#Calculate Cutoff date
#$CutoffDate = (Get-Date).AddDays(-$Days)
 
#Get All Files modified more than the last 30 days
#Get-ChildItem -Path $Path -Recurse -File | Where-Object { $_.LastWriteTime -lt $CutoffDate } #| Remove-Item –Force -Verbose
Get-ChildItem "$Path\$FileNamePatternMatch" -Recurse -File | Where CreationTime -lt  (Get-Date).AddDays(-30)

Stop-Transcript

<#
$WorkDir="D:\AppData\Rundeck"
$CID="DEV9999"
$GamDir="D:\AppData\GAMXTD3\app"
$DataDir="D:\AppData\Rundeck\$CID\Data"
$LogDir="D:\AppData\Rundeck\$CID\Logs"
$transcriptlog = "$LogDir\$(Get-date -Format yyyyMMdd-HHmmss)_transcript.log"

Get-ChildItem "$Root\$RundeckDir\*_transcript.log" -Recurse -File | Where CreationTime -lt  (Get-Date).AddDays(-30) | Remove-Item -verbose

#>