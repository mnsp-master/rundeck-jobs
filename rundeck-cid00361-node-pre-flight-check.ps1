#create common file/folder structure that may/will be required for all jobs in rundeck library

Clear-Host
$mnspver = "1.0.1"

#$Root = "D:"
$Root = "@option.root@"
$CID="C@option.ChangeID@"
$GamDir="$Root\AppData\GAMXTD3\app"
$DataDir="$Root\AppData\Rundeck\$CID\Data"
$LogDir="$Root\AppData\Rundeck\$CID\Logs"
$transcriptlog = "$LogDir\$(Get-date -Format yyyyMMdd-HHmmss)_transcript.log"
$tempcsv="$DataDir\temp.csv"
$now = $(Get-Date -Format "dd MMMM yyyy HHHH:mm:s")

#Start-Transcript -Path $transcriptlog -Force -NoClobber -Append
#Write-Host $(Get-Date)

$structure = @("$DataDir","$LogDir") #base directories to include
$folder = @() #empty array

Foreach ($folder in $structure) {

Write-Host "Looking for folder: " $folder

    if (Test-Path $folder) {
   
        Write-Host "Folder Exists"
        }
        else
            {
  
            #Create directory if not exists
            Write-Host "creating directory structure" $folder
            New-Item $folder -ItemType Directory -Verbose

            }

}


$dataFiles = @("data.csv","temp.csv","temp1.csv","temp2.csv","temp3.csv")
$file = @()

Foreach ($file in $dataFiles) {

Write-Host "Looking for file: " $DataDir\$file

    if (Test-Path $DataDir\$file) {
   
        Write-Host "File Exists nothing to do..."
        }
        else
            {
  
            #Create file(s) if not exists
            Write-Host "creating file:" $DataDir\$file
            New-Item $DataDir\$file -ItemType file -Verbose

            }

}
