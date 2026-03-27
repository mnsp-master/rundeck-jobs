Clear-Host
$mnspver = "0.0.21"

Function Get-NewPassword {
    $PwdUrl = $MNSPgetPasswordPRL

    #failsafe password
    $pwdFailsafe = $MNSPgetPasswordPRLfailsafe

    
    try {
        Write-Host "Attempting to retrieve password from web service..."
        $response = Invoke-WebRequest -Uri $PwdUrl -UseBasicParsing

        # Extract the content from the response object and store it in a variable.
        $password = $response.Content

        # Return the generated password.
        return $password
    }
    catch {
        # If the web request fails, a detailed error message is logged.
        Write-Error "Failed to retrieve password from $PwdUrl. Using failsafe password instead." -ErrorAction Stop

        # Return the failsafe password as the function's output.
        return $pwdFailsafe
    }
}


##############################################
## PRE Main section - prpepare environment ###
##############################################

$ScriptName = Split-Path $PSCommandPath -Leaf
Write-Host "Checking/preparing expected rundeck job environment..."
Write-Host "MNSP script: $scriptName version: $mnspver"

$tempcsv1 = "$DataDir\temp1.csv"
$tempcsv2 = "$DataDir\temp2.csv"
$tempcsv3 = "$DataDir\temp3.csv"
$tempcsv4 = "$DataDir\temp4.csv"
$temptxt1 = "$DataDir\temp1.txt"
$temptxt2 = "$DataDir\temp2.txt"

Write-host "Rundeck job details:"
Write-Host "Rundeck user mail: $mailRecepient"
Write-Host "Rundeck jobname: $jobName"
BlankLIne

Write-Host "Checking expected folders exist..."
# Create Folders
$structure = @($DataDir, $LogDir)
foreach ($dir in $structure) {
    if (!(Test-Path $dir)) {
        Write-Host "Creating directory: $dir"
        New-Item -Path $dir -ItemType Directory -Verbose
    } else {
        Write-Host "Folder already exists: $dir"
    }
}
Write-Host "Checking expected files exist..."
# Create Files
$dataFiles = @("temp1.csv", "temp2.csv", "temp3.csv", "temp4.csv", "temp1.txt","temp2.txt")
foreach ($fileName in $dataFiles) {
    $filePath = Join-Path $DataDir $fileName
    if (!(Test-Path $filePath)) {
        Write-Host "Creating file: $filePath"
        New-Item -Path $filePath -ItemType File -Verbose
    } else {
        Write-Host "File already exists: $filePath"
    }
}
BlankLine

if (Test-Path $GitHubPS01) {
    Write-Host "Path $GitHubPS01 exists. Deleting..."
    # -Recurse handles folders with content; -Force handles read-only files
    Remove-Item -Path $GitHubPS01 -Force -Recurse -Verbose 
} else {
    Write-Host "Path $GitHubPS01 not found, skipping delete."
}
DashedLine

start-sleep 1

##############################################
########## Main section to execute ###########
##############################################

DashedLine
Write-Host "Downloading Main PS script: $GitHubUri01 from github..."

# Ensure TLS 1.2 is enabled (GitHub requires this, older PS versions default to 1.1)
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

try {
    # Download the script
    Invoke-WebRequest -Uri $GitHubUri01 -OutFile $GitHubPS01 -UseBasicParsing
    
    if (Test-Path $GitHubPS01) {
        Write-Host "Download successful. $GitHubPS01..."
        
        # Dot-sourcing the script to run it in the current scope
        . $GitHubPS01
    } else {
        throw "File $GitHubPS01 was not created after download."
    }
}
catch {
    Write-Error "Failed to download or execute Main script: $_"
    exit 1 # Ensures Rundeck sees the failure
}
DashedLine

##############################################
### POST Main section mail transaction log ###
##############################################

$emailBody = @"
<html>
<head>
<style>
  table { border-collapse: collapse; width: 100%; font-family: sans-serif; }
  th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
  th { background-color: #f2f2f2; width: 30%; }
</style>
</head>
<body>
  <h3>Rundeck Job Execution Report</h3>
  <p><strong>NOTE:</strong> Full transatcion rundeck job log attached below</p>
  <table>
    <tr><th>Job Name</th><td>$RundeckJobName</td></tr>
    <tr><th>Job ID</th><td>$JobID</td></tr>
    <tr><th>Project</th><td>$Project</td></tr>
    <tr><th>Execution ID</th><td>$ExecID</td></tr>
    <tr><th>Executed By</th><td>$ExecutingUser</td></tr>
    <tr><th>Timestamp</th><td>$now</td></tr>
    <tr><th>Full Execution (Debug)</th><td><a href='$ExecutionURL'>Click here to view complete debug log (only accessible within MNSP networks)</a></td></tr>
  </table>
  <p>Please find the attached transcript log for full details.</p>
</body>
</html>
"@


#create credential object to authenticate to smtp 
[SecureString]$securepassword = $password | ConvertTo-SecureString -AsPlainText -Force
$credential = New-Object System.Management.Automation.PSCredential -ArgumentList $username, $securepassword

#copy contents of transcript log to temp txt file to use as attachment, as log will be in use, and cannot be directly attached...
$transcriptlogCopy = $(Get-Content $transcriptlog | set-content $transcriptlogTemp)
$attachment = $transcriptlogTemp

$mailParams = @{
    SmtpServer  = "smtp.gmail.com"
    Port        = 587
    UseSsl      = $true
    From        = $from
    To          = $mailRecepient
    Subject     = $subject
    Body        = $emailBody
    BodyAsHtml  = $true
    Attachments = $attachment
    Credential  = $credential
    Verbose     = $true
}

Send-MailMessage @mailParams

start-sleep 2

#clean up tmp transcript
Remove-item $attachment -Force