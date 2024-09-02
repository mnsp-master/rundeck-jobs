$mnspver = "0.0.10"
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

$users = Get-ADUser -Credential $Credentials -filter * -SearchBase $ADSearchBase -properties sAMAccountName,mail,GivenName,Surname | select sAMAccountName,mail,GivenName,Surname

foreach ($user in $users.sAMAccountName) {

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

$Fullname = "$($user.GivenName) $($User.Surname)"
$email = "$($user.mail)"

Write-Host "Processing User: $Fullname $user $email new password as of $(Get-Date): $password"

#Fullname,username,email,password

#Set-ADAccountPassword -Credential $Credentials -Identity $user -Reset -NewPassword (ConvertTo-SecureString -AsPlainText $password -Force) 
 
sleep 1
 
#"$($user),$password" | out-file -filepath $tempcsv -Append
 
Write-Host "`n"
}

