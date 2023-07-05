Clear-Host
$mnspVer = "0.0.0.0.0.1"
#Get-Variable | format-table -Wrap -Autosize
Write-Host "MNSP Script Version: $mnspVer"


Write-Host "Downloading Googlesheet containing all sims report(s) info, instance name,report def, target gsheet etc..."
    if (Test-Path -path $datadir\$SimsReportsSourceCSV ) {

        Write-Host "$datadir\$SimsReportsSourceCSV exists, deleting..."
        Remove-Item -Path $datadir\$SimsReportsSourceCSV -Force
    }

    Set-Location $GamDir

    Invoke-Expression "$GamDir\gam.exe user $GoogleGamMail get drivefile id $SimsReportsSourceCSVGsheetID format csv targetfolder $datadir" -ErrorAction SilentlyContinue
Start-Sleep 2

#force delete any previous data...
Remove-item -Path $DataDir\tmp_*.csv -Force 
Start-sleep 2

$SimsReportDefs = Import-Csv -Path $datadir\$SimsReportsSourceCSV
#Get-Variable | Format-Table -Wrap -AutoSize

Write-Host "loop through each Sims ReportDef..."

foreach ($SimsReportDef in $SimsReportDefs) {

	$now = $(Get-Date -Format "dd MMMM yyyy HHHH:mm:ss")
	$simsServerName = "$($simsreportdef.host)\$($simsreportdef.SQLInstance)"
	$SimsDatabaseName = "$($simsreportdef.dbname)"
	$simsSchool = "$($simsreportdef.school)"
	$GoogleSheetID = "$($simsreportdef.DevGoogleGsheetTargetID)"
	$simsDFE = "$($simsreportdef.DFEnumber)"
    $simsSchoolShortName = "$($simsreportdef.SchoolShortName)"
    $simsReportName = "$($simsreportdef.SimsReportDefName)"
    $GoogleSheetTitle = "$($simsreportdef.GoogleGsheetTitle) - $simsSchoolShortName : ReportRuntime: $now"
    $tempcsv = "$dataDir\tmp_$GoogleSheetID.csv"
    
    write-host "------------------------------------------------------------------"
	write-host "SimsServerName     :" $simsServerName
    Write-Host "Sims DB            :" $SimsDatabaseName
    write-host "GsheetID output    :" $GoogleSheetID
    Write-Host "School             :" $simsSchoolShortName
    Write-Host "DfE num            :" $simsDFE
    Write-Host "Report Name        :" $simsReportName
    Write-Host "Google Sheet Title :" $GoogleSheetTitle

    #create and execute sims commandlinereported command line
    $simsReporterApp = "C:\PROGRA~2\SIMS\SIMS~1.net\CommandReporter.exe /SERVERNAME:$simsServerName /DATABASENAME:$SimsDatabaseName /USER:$SimsReportUser /PASSWORD:$SimsPWD /REPORT:'$simsReportName' /OUTPUT:$tempcsv"
    #Invoke-expression "$simsReporterApp" -ErrorAction SilentlyContinue
    Invoke-Expression "& $simsReporterApp " | Tee-object -variable 'result'
    #$result #uncomment to assist in error checking...
    if ($result -like "*error*" ) {Write-warning "Issue here... $result"}

    #create and execute gamxtd3 command line
    Write-Host "replacing content of existing google sheet with upto date data..."
    $GamApp = "$GamDir\gam.exe user $GoogleGamMail update drivefile id $GoogleSheetID newfilename '$GoogleSheetTitle' localfile $tempcsv"
    #$GamApp
    Invoke-Expression "& $GamApp " | Tee-object -variable 'result2'
    $result2 #uncomment to assist in error checking...
    #if ($result2 -notlike "*Updated with content from*" ) {Write-warning "Issue here... $result2"} #not working - false positive...

}


<#
    #Invoke-Expression "$GamDir\gam.exe user user@domain.com update drivefile id $GoogleDocID newfilename '$GoogleDocTitle' localfile $tempcsv" -ErrorAction SilentlyContinue
    #Invoke-Expression "$GamDir\gam.exe user ######## update drivefile id $GoogleDocID newfilename 'Parent Primary emails - all years: $(Get-Date -Format "dd MMMM yyyy HHHH:mm:s")' localfile $tempcsv convert" -ErrorAction SilentlyContinue
# = "C:\PROGRA~2\SIMS\SIMS~1.net\CommandReporter.exe /SERVERNAME:$simsServerName /DATABASENAME:$SimsDatabaseName /USER:$SimsReportUser /PASSWORD:$SimsPWD /REPORT:'$simsReportName'"
    

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