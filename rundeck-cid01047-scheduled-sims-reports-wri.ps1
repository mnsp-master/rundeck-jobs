Clear-Host
$mnspVer = "0.0.0.0.3"
#Get-Variable | format-table -Wrap -Autosize
Write-Host "MNSP Version: $mnspVer"

Write-Host "Downloading Googlesheet containing all sims report(s) info, instance name,report def, target gsheet etc..."
    if (Test-Path -path $SimsReportsSourceCSV ) {

        Write-Host "$SimsReportsSourceCSV exists, deleting..."
        Remove-Item -Path $SimsReportsSourceCSV -Force
    }

    Set-Location $GamDir

    Invoke-Expression "$GamDir\gam.exe user $GoogleGamMail get drivefile id $SimsReportsSourceCSVGsheetID format csv targetfolder $datadir" -ErrorAction SilentlyContinue
Start-Sleep 2
$SimsReportDefs = Import-Csv -Path $SimsReportsSourceCSV

Write-Host "loop through each Sims ReportDef..."

foreach ($SimsReportDef in $SimsReportDefs) {
    write-host "------------------------------------------------------------------"
	write-host "SCOMIS Host :" $simsreportdef.host
    write-host "MSSQL       :" $simsreportdef.SQLinstance
    Write-Host "Sims DB     :" $simsreportdef.dbname
    write-host "GsheetID    :" $simsreportdef.GoogleGsheetTargetID
    Write-Host "School      :" $simsreportdef.School
    Write-Host "DfE num     :" $simsreportdef.DFEnumber
    Write-Host "Report Name :" $simsreportdef.GoogleGsheetTitle

	$now = $(Get-Date -Format "dd MMMM yyyy HHHH:mm:s")
	$simsServerName = "$($simsreportdef.host)\$($simsreportdef.SQLInstance)"
	$SimsDatabaseName = "$($simsreportdef.dbname)"
	$simsSchool = "$($simsreportdef.school)"
	#$GoogleDocTitle = "DSX Attendance - $simsSchool : StartDate:$XMLdateStart EndDate:$XMLdateEnd ReportRuntime: $now"
	$GoogleDocID = "$($simsreportdef.GoogleGsheetTargetID)"
	$simsDFE = "$($simsreportdef.DFEnumber)"
    #$simsReporterImporter = "C:\PROGRA~2\SIMS\SIMS~1.net\CommandReporterImporter.exe /SERVERNAME:$simsServerName /DATABASENAME:$SimsDatabaseName /USER:$SimsReportUser /PASSWORD:$SimsPWD /REPORT:'$SimsReport'"
    #$simsReporterImporter

}


<#
Write-Host "Passed vars..."
$SimsReportsSourceCSV
$SimsReport
$GoogleDocIDsimsInstances
$GoogleGamMail
$SimsReportUser

$AutomaticVariables = Get-Variable
function cmpv {
    Compare-Object (Get-Variable) $AutomaticVariables -Property Name -PassThru | Where -Property Name -ne "AutomaticVariables"
}

cmpv

Write-Host "Downloading Sims Report Definition to distribute to all sims instances..."
    if (Test-Path -path $SimsReport ) {

        Write-Host "$SimsReport exists, deleting..."
        Remove-Item -Path $SimsReport -Force
    }

    Set-Location $GamDir
    Invoke-Expression "$GamDir\gam.exe user $GoogleGamMail get drivefile id $SimsReportGoogleDocID targetname $SimsReport" -ErrorAction SilentlyContinue
#>