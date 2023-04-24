#upload sims report definitions to numerous sims instances
Clear-Host

$CID="DEV9999"
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

Clear-Host
Write-Host "Downloading Googlesheet containing all sims instances, instance name etc..."
if (Test-Path -path $SimsInstancesCSV ) {

	Write-Host "$SimsInstancesCSV exists, deleting..."
	Remove-Item -Path $SimsInstancesCSV -Force
}

Set-Location $GamDir
Invoke-Expression "$GamDir\gam.exe user $GoogleGamMail get drivefile id $GoogleDocIDsimsInstances format csv targetfolder $datadir" -ErrorAction SilentlyContinue
Start-Sleep 2
$simsinstances = Import-Csv -Path $SimsInstancesCSV


Write-Host "Downloading Googlesheet containing all Sims Report Definitions to import, name, google id, checksum etc..."
if (Test-Path -path $SimsReportDefsCSV ) {

	Write-Host "$SimsReportDefsCSV exists, deleting..."
	Remove-Item -Path $SimsReportDefsCSV -Force
}

Set-Location $GamDir
Invoke-Expression "$GamDir\gam.exe user $GoogleGamMail get drivefile id $GoogleDocIDsimsInstances format csv targetfolder $datadir" -ErrorAction SilentlyContinue
Start-Sleep 2
$simsReportDefinitions = Import-Csv -Path $SimsReportDefsCSV



#loop through each sims instance
foreach ($sims in $simsinstances) {
	write-host "IP           :" $sims.ip
    write-host "MSSQL        :" $sims.SQLinstance
    Write-Host "Sims DB      :" $sims.dbname
    write-host "GdocID       :" $sims.GoogleDocID
    Write-Host "School       :" $sims.School
    Write-Host "DfE num      :" $sims.DFEnumber
    Write-Host "Rundeck User :" $sims.SimsUser

	$now = $(Get-Date -Format "dd MMMM yyyy HHHH:mm:s")
	$simsServerName = "$($sims.ip)\$($sims.SQLInstance)"
	$SimsDatabaseName = "$($sims.dbname)"
	$simsSchool = "$($sims.school)"
	#$GoogleDocTitle = "DSX Attendance - $simsSchool : StartDate:$XMLdateStart EndDate:$XMLdateEnd ReportRuntime: $now"
	#$GoogleDocID = "$($sims.GoogleDocID)"
	$simsDFE = "$($sims.DFEnumber)"	
    $SimsUser = "$($sims.SimsUser)"


		Write-Host "Importing sims report $SimsReport using command line report Importer..."
		clear-content $tempcsv -Force
		#create sims commandlinereported command line
		#$simsApp = "C:\PROGRA~2\SIMS\SIMS~1.net\CommandReporter.exe /SERVERNAME:$simsServerName /DATABASENAME:$SimsDatabaseName /USER:$SimsUser /PASSWORD:$SimsPWD /REPORT:'$SimsReport' /OUTPUT:$tempcsv"
		$simsApp = "C:\PROGRA~2\SIMS\SIMS~1.net\CommandReportImporter.exe /SERVERNAME:$simsServerName /DATABASENAME:$SimsDatabaseName /USER:$SimsUser /PASSWORD:$SimsPWD /REPORT:'$SimsReport'"

		try {
		Invoke-Expression "& $simsApp"
		
		    
		} catch {

			Write-warning "Issue here: $error[0] $message"
		}

		#$ErrorActionPreference="Stop"

		#Write-Host "replacing content of existing google sheet with upto date data..."
		#Write-Host "Title: $GoogleDocTitle"
		#Write-Host "DocID: $GoogleDocID"


		#Set-Location $GamDir
		#Write-Host "replacing content of existing google sheet with upto date data..."
		#$gamApp = "$GamDir\gam.exe user $GoogleGamMail update drivefile id $GoogleDocID newfilename '$GoogleDocTitle' localfile $tempcsv"
		#$gamApp
		#Invoke-expression "& $gamApp"
		###Invoke-Expression "$GamDir\gam.exe user $GoogleGamMail update drivefile id $GoogleDocID newfilename '$GoogleDocTitle' localfile $tempcsv" -ErrorAction SilentlyContinue
		write-host "------------------------------------`n"	
}

Stop-Transcript

<#
online file checksum tool
https://emn178.github.io/online-tools/sha256_checksum.html

$XMLdateStart = "2021-09-02"
$XMLdateEnd = "$((Get-Date).ToString('yyyy-MM-dd'))"

$SimsParamXML = "$DataDir\dsx-attendance.xml"
#$SimsParamTXT = "$DataDir\dsx-attendance.txt"
$SimsStartDate = "DATE_START"
$SimsEndDate = "DATE_END"
$simsinstances = @()

#end date of report - now
$SimsReportEndDate = $((Get-date).ToString('yyyy-MM-ddTHH:mm:ss'))
$SimsReportEndDate
Write-Host "XML Start Date :" $XMLdateStart "(UK date format)"
Write-Host "XML End Date   :" $XMLdateEnd  "(UK date format)"
Write-Host "`n"

#generate xml from sims report
$simsAppXML = "C:\PROGRA~2\SIMS\SIMS~1.net\CommandReporter.exe /SERVERNAME:$simsServerName /DATABASENAME:$SimsDatabaseName /USER:$SimsUser /PASSWORD:$SimsPWD /REPORT:'$SimsReport' /PARAMDEF /OUTPUT:$SimsParamXML"

try {
		Invoke-Expression "& $simsAppXML"
		#$simsApp 
		    
		} catch {

			Write-warning "Issue here: $error[0] $message"
		}



#replace all end dates in exported sims xml
$xml = @()
$xml =[xml](Get-content $SimsParamXML)

$nodes = $xml.SelectNodes("//End")
$nodes

foreach ($node in $nodes) {

$node.InnerText
$node.InnerText = "$SimsReportEndDate"

}

$xml.Save($SimsParamXML)



#>