Clear-Host
$mnspVer = "0.0.0.0.0.1.9"
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
Remove-item -Path $DataDir\tmp_*.* -Force -verbose
Start-sleep 2

$SimsReportDefs = Import-Csv -Path $datadir\$SimsReportsSourceCSV
#Get-Variable | Format-Table -Wrap -AutoSize

Write-Host "loop through each Sims Instance..."

foreach ($SimsReportDef in $SimsReportDefs) {

	$now = $(Get-Date -Format "dd MMMM yyyy HHHH:mm:ss")
	$simsServerName = "$($simsreportdef.host)\$($simsreportdef.SQLInstance)"
	$SimsDatabaseName = "$($simsreportdef.dbname)"
	$simsSchool = "$($simsreportdef.school)"
	$GoogleSheetID = "$($simsreportdef.GoogleDocID)"
	$simsDFE = "$($simsreportdef.DFEnumber)"
    $simsSchoolShortName = "$($simsreportdef.SchoolShortName)"
    $simsSchoolFullName = "$($simsreportdef.SchoolFullName)"
    #$simsReportName = "$($simsreportdef.SimsReportDefName)"
    #$GoogleSheetTitle = "$($simsreportdef.GoogleGsheetTitle) - $simsSchoolShortName : ReportRuntime: $now"
    $GoogleSheetTitle = "Sims Report: $simsReportName : ReportRuntime: $now"
    $tempcsv = "$dataDir\tmp_$simsSchoolShortName.csv"
    $tempcsvutf8 = "$dataDir\tmp_$($simsSchoolShortName)_utf8.csv"
    $SimsParamXML = "$DataDir\tmp_$($simsSchoolShortName).xml"
    $SimsReportEndDate = $((Get-date).ToString('yyyy-MM-ddTHH:mm:ss'))
    
    write-host "------------------------------------------------------------------"
	write-host "SimsServerName      :" $simsServerName
    Write-Host "Sims DB             :" $SimsDatabaseName
    write-host "GsheetID output     :" $GoogleSheetID
    Write-Host "School Short Name   :" $simsSchoolShortName
    Write-Host "School Long Name    :" $simsSchoolFullName
    Write-Host "DfE num             :" $simsDFE
    Write-Host "Report Name         :" $simsReportName
    Write-Host "Google Sheet Title  :" $GoogleSheetTitle
    Write-Host "Sims Report End date:" $SimsReportEndDate
    write-host "------------------------------------------------------------------"

    Write-Host "Generate xml from sims report cli..."
    
    $simsAppXML = "C:\PROGRA~2\SIMS\SIMS~1.net\CommandReporter.exe /SERVERNAME:$simsServerName /DATABASENAME:$SimsDatabaseName /USER:$SimsReportUser /PASSWORD:$SimsPWD /REPORT:'$simsReportName' /PARAMDEF /OUTPUT:$SimsParamXML"

    try {
            # Generating XML from sims report...
            Invoke-Expression "& $simsAppXML"
            $simsApp #enable to output full cli to transaction log
                
            } catch {

                Write-warning "Issue here: $error[0] $message"
            }

                Write-Host "replacing all end dates in exported sims xml..."
                $xml = @()
                $xml =[xml](Get-content $SimsParamXML)

                $nodes = $xml.SelectNodes("//End")
                #$nodes print nodes to transcript log

                foreach ($node in $nodes) {

                #$node.InnerText #print all existing node dates to transcript log...
                $node.InnerText = "$SimsReportEndDate"

        }
        $xml.Save($SimsParamXML) #replace extracted XML content with updated/date replaced nodes...

    #create and execute sims commandlinereported command line
    #$simsReporterApp = "C:\PROGRA~2\SIMS\SIMS~1.net\CommandReporter.exe /SERVERNAME:$simsServerName /DATABASENAME:$SimsDatabaseName /USER:$SimsReportUser /PASSWORD:$SimsPWD /REPORT:'$simsReportName' /OUTPUT:$tempcsv"
    
    #create and execute sims commandlinereported command line - including xml params
    $simsReporterApp = "C:\PROGRA~2\SIMS\SIMS~1.net\CommandReporter.exe /SERVERNAME:$simsServerName /DATABASENAME:$SimsDatabaseName /USER:$SimsReportUser /PARAMFILE:$SimsParamXML /PASSWORD:$SimsPWD /REPORT:'$simsReportName' /OUTPUT:$tempcsv"

    #$simsReporterApp #enable to output full cli to transaction log
    Invoke-Expression "& $simsReporterApp " | Tee-object -variable 'result'
    #$result #uncomment to assist in error checking...
    if ($result -like "*error*" ) {Write-warning "Issue here... $result"}

    #fix for issue: 'utf-8' codec can't decode byte 0xe9 in position 2: invalid continuation byte
    Write-Host "convert sims csv output to utf8..."
    start-sleep 3
    get-content $tempcsv | set-content -encoding utf8 $tempcsvutf8

    #create and execute gamxtd3 command line
    Write-Host "replacing content of existing google sheet tab with upto date data..."
    $GamApp = "$GamDir\gam.exe user $GoogleGamMail update drivefile id $GoogleSheetID newfilename '$GoogleSheetTitle' csvsheet $simsSchoolShortName localfile $tempcsvutf8"
    #$GamApp #enable to output full cli to transaction log
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