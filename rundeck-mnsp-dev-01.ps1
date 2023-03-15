Clear-Host

$CID="@option.ChangeID@"
$GamDir="D:\AppData\GAMXTD3\app"
$DataDir="D:\AppData\Rundeck\$CID\Data"
$LogDir="D:\AppData\Rundeck\$CID\Logs"
$transcriptlog = "$LogDir\$(Get-date -Format yyyyMMdd-HHmmss)_transcript.log"
$tempcsv="$DataDir\temp.csv"
$tempcsv2="$DataDir\temp2.csv"
$now = $(Get-Date -Format "dd MMMM yyyy HHHH:mm:s")
$SimsPWD = "@option.SIMSuserPASS@" #rundeck key vault

$SimsReport = "$DataDir\IT_DSXAttendanceYTD_dev.RptDef" #convert to rundeck defined option

$GoogleDocIDsimsInstances = "1R6l9sZKnY2Ii6qgcyYNoRccIZg0tZL6jhwCgzzPOvs4" #convert to rundeck defined option
$GoogleDocIDsimsReportDefinitions = "1SYowGUfzpnVBx3MlVz0aVlVMF2NVrpzgBHWXRKK2dHk" #convert to rundeck defined option

$SimsInstancesCSV = "$datadir\DSX-sims-instances-CID00120.csv" #convert to rundeck defined option
$SimsReportDefsCSV = "$dataDir\DEV-sims-ReportDefChecksums.csv" #convert to rundeck defined option
$GoogleGamMail = "@option.GoogleGamMail@" #rundeck key vault

Start-Transcript -Path $transcriptlog -Force -NoClobber -Append

Clear-Host
Write-Host "Downloading Googlesheet containing all sims instances, instance name etc..."

Stop-Transcript
