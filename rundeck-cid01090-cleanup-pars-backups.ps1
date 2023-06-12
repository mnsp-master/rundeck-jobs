cls

$WorkDir="D:\AppData\Rundeck"
$DataDir="D:\AppData\Rundeck\C01090"
$log = "$DataDir\Logs\transcript.log"

Start-Transcript -Path $log -Force -NoClobber -Append

Write-Host $(Get-Date)

$ErrorActionPreference="Stop"

#Set-Location $GamDir

#Parameters
$Path = "D:\Microsoft SQL Server\MSSQL12.SDSSIMS\MSSQL\Backup" # Path where the file is located 
$Days = "30" # Number of days before current date
$FileNameMatch = "PARS_*.BAK"
 
#Calculate Cutoff date
$CutoffDate = (Get-Date).AddDays(-$Days)
 
#Get All Files modified more than the last 30 days
Get-ChildItem -Path $Path -Recurse -File | Where-Object { $_.LastWriteTime -lt $CutoffDate } #| Remove-Item –Force -Verbose

Stop-Transcript
