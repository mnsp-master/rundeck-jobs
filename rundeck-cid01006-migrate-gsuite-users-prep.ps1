Clear-Host
$mnspVer = "0.0.0.1.0"

Write-Host "MNSP Version: $mnspVer"
#Get-Variable | format-table -Wrap -Autosize

#$csv = "C:\temp\users.csv"
#$migratedDomain = "dev.mnsp.org.uk"
$SourceGsuiteOrgUnits = @("$SourceGsuiteOrgUnit1","$SourceGsuiteOrgUnit2")
set-location $GamDir


invoke-expression ".\gam.exe ou_and_children 'staff' print allfields  | out-file $tempcsv" #produce csv header

$users = @()
$users = import-csv -Path $tempcsv | Where-Object { $_.suspended -notlike "True" } AND { $_.primaryEmail -like "mendipstudioschool.org.uk"} #exclude any suspended accounts

Write-host "Number of source users to process..." $users.count

foreach ($user in $users) {

$username = @($user.primaryEmail.Split("@")[0])
$migratedUser =  "$username@$migratedDomain"
$migratedUser

} 

