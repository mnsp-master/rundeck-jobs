#common template that will download raw defined ps script and execute it, substituting localised vars as defined in the rundeck job

Clear-Host
$mnspver = "0.0.0.0.0.3"
$Root = ":@option.Root@"
$CID = "C@option.ChangeID@"
$GamDir = "$Root\AppData\GAMXTD3\app"
$DataDir = "$Root\AppData\Rundeck\$CID\Data"
$LogDir = "$Root\AppData\Rundeck\$CID\Logs"
$transcriptlog = "$LogDir\$(Get-date -Format yyyyMMdd-HHmmss)_transcript.log"
#$tempcsv = "$DataDir\temp.csv"
#$temptxt = "$DataDir\temp.txt"
$GitHubPS = "$DataDir\JobToRun.ps1"
$GitHubUri = "@option.RawGitHubRepoSource@"

#set/update/add variables (using rundeck job option(s) values) below as required by the individual local rundeck job:
#e.g: $MyVariable1 = "@option.rundeck-option@"
$GLPIapiAppToken = "@option.GLPI-API-App-Token-01@"
$GLPIuserApiToken = "@option.GLPI-API-User-Token-01@"
$AppURL = "@option.GLPI-api-AppURL@"

#start transaction log
Start-Transcript -Path $transcriptlog -Force -NoClobber -Append
Write-Host $(Get-Date)

#download raw powershell script from github
Invoke-WebRequest -Uri $GitHubUri -OutFile $GitHubPS

$ErrorActionPreference="Continue"

#execute downloaded raw github PS script (will inherrit variables defined above)...
. $GitHubPS

Stop-Transcript
