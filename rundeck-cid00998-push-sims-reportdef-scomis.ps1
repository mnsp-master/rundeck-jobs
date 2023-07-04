Clear-Host
$mnspVer = "0.0.0.0.9.8"
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

Write-Host "Downloading Sims Report Definitions to distribute to all sims instances..."
    if (Test-Path -path $SimsReport ) {

        Write-Host "$SimsReport exists, deleting..."
        Remove-Item -Path $SimsReport -Force
    }

    Set-Location $GamDir
    #Invoke-Expression "$GamDir\gam.exe user $GoogleGamMail get drivefile id $SimsReportGoogleDocID targetname '$SimsReport'" -ErrorAction SilentlyContinue
    $GamRepDefsGet = "$GamDir\gam.exe user $GoogleGamMail print filelist select $SimsReportDefsGoogleFolderID fields id,name > $tempcsv" #-ErrorAction SilentlyContinue
    Invoke-expression "& $GamRepDefsGet"
    
    $SimsReportDefsArray = Import-Csv -Path $tempcsv
    #$SimsReportDefsArray 
    #$SimsReportDefsArray.name

foreach ($SimsReportsDef in $SimsReportDefsArray) { 
        write-host "------------------------------------------------------------------"
        Write-Host "Pushing Report definition..."
        Write-Host "ReportDefName: " $SimsReportsDef.name
        Write-Host "ID           : " $SimsReportsDef.id
        $ReportDefName = $($SimsReportsDef.name)

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
                #$simsReporterImporter = "C:\PROGRA~2\SIMS\SIMS~1.net\CommandReporterImporter.exe /SERVERNAME:$simsServerName /DATABASENAME:$SimsDatabaseName /USER:$SimsReportUser /PASSWORD:$SimsPWD /REPORT:'$SimsReport'"
                $simsReporterImporter = "C:\PROGRA~2\SIMS\SIMS~1.net\CommandReporterImporter.exe /SERVERNAME:$simsServerName /DATABASENAME:$SimsDatabaseName /USER:$SimsReportUser /PASSWORD:$SimsPWD /REPORT:'$ReportDefName'"
                
                $simsReporterImporter

                }
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
