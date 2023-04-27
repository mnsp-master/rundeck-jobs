Clear-Host
$mnspVer = "0.0.0.0.2"

Write-Host "MNSP Version: $mnspVer"
Get-Variable | format-table -Wrap -Autosize

#$csv = "C:\temp\users.csv"
#$migratedDomain = "dev.mnsp.org.uk"

set-location $GamDir

invoke-expression ".\gam.exe ou_and_children 'staff/Non-Teaching Staff/men' print allfields  | out-file $tempcsv"

$users =@()
$users = import-csv -Path $tempcsv #| Where-Object { $_.suspended -notlike "True" } #exclude any suspended accounts
$users

Write-host "Number of source users to process..." $users.count
#$users.suspended

foreach ($user in $users) {

$username = @($user.primaryEmail.Split("@")[0])
$migratedUser =  "$username@$migratedDomain"
$migratedUser

} 
