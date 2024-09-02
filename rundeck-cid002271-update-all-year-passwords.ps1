$mnspver = "0.0.2"
Write-Host $(Get-Date)
Write-Host "MNSP Version" $mnspver

$ErrorActionPreference="Continue"
Set-Location $GamDir

Write-Host "emptying temp csv of all data..."
Clear-Content $tempcsv
sleep 2
write-host $csvheader
$csvheader | out-file -filepath $tempcsv -Append #create blank csv with simple header