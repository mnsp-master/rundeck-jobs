Clear-Host

$CID = "@option.ChangeID@"
$GamDir = "D:\AppData\GAMXTD3\app"
$DataDir =" D:\AppData\Rundeck\$CID\Data"
$LogDir = "D:\AppData\Rundeck\$CID\Logs"
$transcriptlog = "$LogDir\$(Get-date -Format yyyyMMdd-HHmmss)_transcript.log"
$tempcsv = "$DataDir\temp.csv"
$tempcsv2 = "$DataDir\temp2.csv"
$now = $(Get-Date -Format "dd MMMM yyyy HHHH:mm:s")
$SimsPWD = "@option.SIMSuserPASS@" #rundeck key vault

$SimsReport = "@option.SimsReport@" # sims report to download push
$SimsReportUser = "@option.SimsReportUser@" # sims user to get/push report def
$GoogleDocIDsimsInstances = "@option.GoogleDocIDsimsInstances@" #ghseet containing all instances to be processed
$GoogleDocReportSourceID = "@option.GoogleDocReportSourceID@" #exported/uploaded sims report definition

$SimsInstancesCSV = "$datadir\@option.SimsInstancesCSV@.csv"
$GoogleGamMail = "@option.GoogleGamMail@"

Clear-Host
Get-Variable | format-table -Wrap -Autosize

Write-Host "Downloading Googlesheet containing all sims instances, instance name etc..."
Write-Host "Passed vars..."
$SimsInstancesCSV

Stop-Transcript
