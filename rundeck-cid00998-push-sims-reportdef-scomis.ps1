Clear-Host
#Get-Variable | format-table -Wrap -Autosize

Write-Host "Passed vars..."
$SimsInstancesCSV
$SimsReport
$GoogleDocIDsimsInstances
$GoogleGamMail
$SimsReportUser

Write-Host "Downloading Googlesheet containing all sims instances, instance name etc..."

if (Test-Path -path $SimsInstancesCSV ) {

	Write-Host "$SimsInstancesCSV exists, deleting..."
	Remove-Item -Path $SimsInstancesCSV -Force
}

Set-Location $GamDir

Invoke-Expression "$GamDir\gam.exe user $GoogleGamMail get drivefile id $GoogleDocIDsimsInstances format csv targetfolder $datadir" -ErrorAction SilentlyContinue
Start-Sleep 2
$simsinstances = Import-Csv -Path $SimsInstancesCSV.csv

#loop through each sims instance
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

}


<#
$AutomaticVariables = Get-Variable
function cmpv {
    Compare-Object (Get-Variable) $AutomaticVariables -Property Name -PassThru | Where -Property Name -ne "AutomaticVariables"
}

cmpv
#>