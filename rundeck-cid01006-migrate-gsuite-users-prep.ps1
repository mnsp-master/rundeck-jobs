Clear-Host
$mnspVer = "0.0.0.1.6"

Write-Host "MNSP Version: $mnspVer"
#Get-Variable | format-table -Wrap -Autosize

#$csv = "C:\temp\users.csv"
#$migratedDomain = "dev.mnsp.org.uk"
$SourceGsuiteOrgUnits = @("$SourceGsuiteOrgUnit1","$SourceGsuiteOrgUnit2")
set-location $GamDir


invoke-expression ".\gam.exe ou_and_children 'staff' print allfields  | out-file $tempcsv" #produce csv header

$users = @()
$users = import-csv -Path $tempcsv | Where-Object { ($_.suspended -notlike "True") -and ($_.primaryEmail -like "*mendipstudioschool.org.uk")} #exclude any suspended accounts, include only @mendipstudioschool.org.uk 

Write-host "Number of source users to process..." $users.count

foreach ($user in $users) {
Write-Host "--------------------------------------------------------"
$username = @($user.primaryEmail.Split("@")[0])
$migratedUser =  "$username@$migratedDomain"
$surName = $user.'name.familyName'
$firstName = $user.'$user.givenName'
#$user
$user.id
$user.primaryEmail
$surName
$firstName
$user.fullName
$user.orgUnitPath
$migratedUser

} 

