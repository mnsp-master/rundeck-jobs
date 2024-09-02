$mnspver = "0.0.15"
Write-Host $(Get-Date)
Write-Host "MNSP Version" $mnspver

$ErrorActionPreference="Continue"
Set-Location $GamDir

Write-Host "emptying temp csv of all data..."
Clear-Content $tempcsv
sleep 2
write-host $csvheader
$csvheader | out-file -filepath $tempcsv -Append #create blank csv with simple header

$PlainPassword = $ADDomainPass
$SecurePassword = $PlainPassword | ConvertTo-SecureString -AsPlainText -Force
$UserName = $ADDomainUser
$Credentials = New-Object System.Management.Automation.PSCredential -ArgumentList $UserName, $SecurePassword

$users = Get-ADUser -Credential $Credentials -filter * -SearchBase $ADSearchBase -properties sAMAccountName,mail,displayName | select sAMAccountName,mail,DisplayName
#$users

foreach ($user in $users) {

$pwd = $(Invoke-WebRequest -Uri $PwdGenURL -UseBasicParsing)
#    $pwd.Content
    #$pwd.StatusCode
        if ($pwd.StatusCode -eq 200) {
        Write-Host "proceed with pwd reservation"
        $password = $($pwd.Content)
        #Write-Host "Password: " $password
        } else {
        Write-Error "No Webserver, or pwd received"
        $password = "Js653151MH$"
        }

$Fullname = $($user.displayName)
$email = $($user.mail)

Write-Host "Processing User: $($user.displayName) $($user.mail) $($user.sAMAccountName) new password as of $(Get-Date): $password"

#Fullname,username,email,password

#Set-ADAccountPassword -Credential $Credentials -Identity $user -Reset -NewPassword (ConvertTo-SecureString -AsPlainText $password -Force) 
 
sleep 1
 
"$($user.displayName),$($user.sAMAccountName),$($user.mail),$password" | out-file -filepath $tempcsv -Append
 
Write-Host "`n"
}

