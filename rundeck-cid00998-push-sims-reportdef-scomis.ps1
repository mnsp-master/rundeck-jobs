Clear-Host
$mnspVer = "0.0.0.0.9"
#Get-Variable | format-table -Wrap -Autosize
Write-Host "MNSP Version: $mnspVer"



Write-Host "Downloading Googlesheet containing all sims instances, instance name etc..."
    if (Test-Path -path $SimsInstancesCSV ) {

        Write-Host "$SimsInstancesCSV exists, deleting..."
        Remove-Item -Path $SimsInstancesCSV -Force
    }

    Set-Location $GamDir

    Invoke-Expression "$GamDir\gam.exe user $GoogleGamMail get drivefile id $GoogleDocIDsimsInstances format csv targetfolder $datadir" -ErrorAction SilentlyContinue
Start-Sleep 2
$simsinstances = Import-Csv -Path $SimsInstancesCSV

Write-Host "Downloading Sims Report Definition to distribute to all sims instances..."
    if (Test-Path -path $SimsReport ) {

        Write-Host "$SimsReport exists, deleting..."
        Remove-Item -Path $SimsReport -Force
    }

    Set-Location $GamDir
    Invoke-Expression "$GamDir\gam.exe user $GoogleGamMail get drivefile id $SimsReportGoogleDocID targetname '$SimsReport'" -ErrorAction SilentlyContinue
    Invoke-Expression "$GamDir\gam.exe user $GoogleGamMail print filelist select $SimsReportDefsGoogleFolderID fields id,name > $tempcsv" -ErrorAction SilentlyContinue
    #gam user snoble@writhlington.org.uk print filelist select 1Mw7sWJIdXAGJLBbbOV0CXlX-NibOX4o5 fields id,name > f:\tmp\test.csv

Write-Host "loop through each sims instance..."

foreach ($sims in $simsinstances) {
    write-host "------------------------------------------------------------------"
	write-host "IP      :" $sims.ip
    write-host "MSSQL   :" $sims.SQLinstance
    Write-Host "Sims DB :" $sims.dbname
    write-host "GdocID  :" $sims.GoogleDocID
    Write-Host "School  :" $sims.School
    Write-Host "DfE num :" $sims.DFEnumber

	$now = $(Get-Date -Format "dd MMMM yyyy HHHH:mm:s")
	$simsServerName = "$($sims.ip)\$($sims.SQLInstance)"
	$SimsDatabaseName = "$($sims.dbname)"
	$simsSchool = "$($sims.school)"
	#$GoogleDocTitle = "DSX Attendance - $simsSchool : StartDate:$XMLdateStart EndDate:$XMLdateEnd ReportRuntime: $now"
	$GoogleDocID = "$($sims.GoogleDocID)"
	$simsDFE = "$($sims.DFEnumber)"
    $simsReporterImporter = "C:\PROGRA~2\SIMS\SIMS~1.net\CommandReporterImporter.exe /SERVERNAME:$simsServerName /DATABASENAME:$SimsDatabaseName /USER:$SimsReportUser /PASSWORD:$SimsPWD /REPORT:'$SimsReport'"
    $simsReporterImporter

}


<#
Write-Host "Passed vars..."
$SimsInstancesCSV
$SimsReport
$GoogleDocIDsimsInstances
$GoogleGamMail
$SimsReportUser

$AutomaticVariables = Get-Variable
function cmpv {
    Compare-Object (Get-Variable) $AutomaticVariables -Property Name -PassThru | Where -Property Name -ne "AutomaticVariables"
}

cmpv
#>
