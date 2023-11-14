#common template that will download raw defined ps script and execute it, substituting localised vars as defined in rundeck job

Clear-Host
$mnspver = "0.0.0.0.0.1"
$Root = ":@option.Root@"
$CID="C@option.ChangeID@"
$GamDir="$Root\AppData\GAMXTD3\app"
$DataDir="$Root\AppData\Rundeck\$CID\Data"
$LogDir="$Root\AppData\Rundeck\$CID\Logs"
$transcriptlog = "$LogDir\$(Get-date -Format yyyyMMdd-HHmmss)_transcript.log"
#$tempcsv = "$DataDir\temp.csv"
#$temptxt = "$DataDir\temp.txt"
$GitHubPS = "$DataDir\JobToRun.ps1"
$GitHubUri = "@option.RawGitHubRepoSource@"

#set/update/add vars below as required by the individual local rundeck job:
$GLPIapiAppToken = "@option.GLPI-API-App-Token-01@"
$GLPIuserApiToken = "@option.GLPI-API-User-Token-01@"
$AppURL = "@option.GLPI-api-AppURL@"
$UserToken = "@option.GLPI-API-User-Token-01@"
$AppToken = "@option.GLPI-API-App-Token-01@"

Start-Transcript -Path $transcriptlog -Force -NoClobber -Append
Write-Host $(Get-Date)

Invoke-WebRequest -Uri $GitHubUri -OutFile $GitHubPS

$ErrorActionPreference="Continue"

#execute downloaded raw github PS script...
. $GitHubPS

Stop-Transcript
