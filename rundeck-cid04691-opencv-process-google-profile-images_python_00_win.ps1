$mnspver = "1.0.2" #use python for all image coordinates
Clear-Host

function DashedLine {
Write-host "-----------------------------------------------------------`n"
}

$workdir = "$HOME\opencv-dev" #windows dev environment
$LogDir = "$workdir\Logs"
$now = $(Get-date -Format yyyyMMdd-HHmmss)
$transcriptlog = "$LogDir\$(Get-date -Format yyyyMMdd-HHmmss)_transcript.log"
$datadir = "$workdir\Data"

$datasrc = "$dataDir\source\"
$dataout = "$datadir\output\$now"
$passports = "$dataout\250x250"
$vignettes = "$dataout\vignettes"
$temp = "$dataout\temp"
$exiftoolAppVersion = "exiftool-13.55_64" # https://exiftool.org/

$photosSrc = $(Get-ChildItem -Path $datasrc )
$ImgDimensions = @()
$ImgDimensionX = @()
$ImgDimensionY = @()
$convertX = "x"

Start-Transcript -Path $transcriptlog -Force -NoClobber -Append
Write-Host "MNSP version:" $mnspver

#get python script from github
    $gitHubPythonSrcURI = "https://raw.githubusercontent.com/mnsp-master/rundeck-jobs/refs/heads/main/rundeck-cid04691-opencv-process-images_python_dev-02-win.py"
    $pythonScriptName = ($gitHubPythonSrcURI -split '/')[-1]

    if (Test-Path "$workdir\$pythonScriptName") {
        Remove-item "$workdir\$pythonScriptName" -Force -verbose
    }

    try {
    Write-host "downloading python script: $pythonScriptName from github..."
    invoke-webrequest -Uri $gitHubPythonSrcURI -OutFile "$workdir\$pythonScriptName"
    } catch {
        Write-Error "Failed to download $pythonScriptName Error: $($_.Exception.Message)"
        exit 1
    }

#create directory structure
New-item -Path $dataout -ItemType Directory
New-item -Path $passports -ItemType Directory
New-item -Path $vignettes -ItemType Directory
DashedLine

foreach ($photo in $photosSrc) {
    Write-Host "Processing Image Details:"
    $filePath = $photo.FullName
    $fileName = $photo.name
    $fileBaseName = $photo.BaseName
    
    $metaData = & "$workdir\$exiftoolAppVersion\exiftool.exe" -json $filePath | convertFrom-Json
    $ImgEXIFDateTimeOriginal = $metaData.DateTimeOriginal
    $ImgDimensionX = $metaData.ImageWidth
    $ImgDimensionY = $metaData.Imageheight

    Write-Host "FilePath: $filePath"
    Write-Host "FileName: $fileName"
    Write-Host "BaseName: $fileBaseName `n"

    #resize source image if original is too large - opencv can return unpredictable results if the source image is too large
    if ( $ImgDimensionX -gt 1030 ) {
        Write-Warning "Source Image: $filename is too large for open-cv processing, scaling down..."
        & $WorkDir\ImageMagick\magick.exe $filePath -resize "1024X>" $filePath
        start-sleep 1
        $metaData = & "$workdir\$exiftoolAppVersion\exiftool.exe" -json $filePath | convertFrom-Json
        $ImgDimensionX = $metaData.ImageWidth
        $ImgDimensionY = $metaData.Imageheight
        Write-Host "Updated image dimensions:" $ImgDimensionX $ImgDimensionY
    }

                
        #generate image with bounding box,center and csv data of metadata for image...
        & python "$workDir\$pythonScriptName" $filePath $dataOut

            #set coordinates from python processing...
            if (Test-path -path $dataout\$FileBaseName.csv) {
            
            #adress multiple faces detected in csv, select highest confidence value row...
            $pythonCoords = Import-csv -path $dataOut\$FileBaseName.csv |
                Sort-object { [double]$_.confidence } -descending |
                Select-object -first 1

            Write-Host "Python Library coordinates..."
            $pythonCoords
            $unit = [Math]::Round([double]$pythonCoords.CenterX - [double]$pythonCoords.StartX)
                } else {
                Write-Host "No Face csv detected, moving to next image..."
                DashedLine
                continue
            }

        #only proceed if facial detection confidence is above 0.# percent...
        $faceDetectionScore = [double]$PythonCoords[0].confidence #[double] deals with decimal values
        if ($faceDetectionScore -ge 0.8) {
            Write-Host "Face detected with confidence:" $faceDetectionScore
            Write-Host "Image Dimension X     :" $ImgDimensionX
            Write-Host "Image Dimension Y     :" $ImgDimensionY
            Write-Host "EXIF DateTimeOriginal :" $ImgEXIFDateTimeOriginal
            Write-Host "Square Unit           :" $unit
                
            $OriginTop = $($pythonCoords.CenterY - ($unit * 2)) #new Horizontal Origin Value
            Write-Host "New Horizontal Origin :" $OriginTop
                
            $OriginLeft = $($pythonCoords.CenterX - ($unit *2))
            Write-Host "New Vertical Origin   :" $OriginLeft
                
            $Coords = $($unit * 2)    
            $CoordXY = ($unit * 4)
                
            Write-Host "New XY Coordinate     :" $CoordXY
            Write-Host "Output Image          : " $dataout\$fileName

            if ($OriginTop -lt 0 -or $OriginLeft -lt 0) {
                Write-Warning "One or more Origin values: $OriginTop $OriginLeft are negative, consider providing a better image..."
            }
            
            #remove background from source image first...
            $TMPIMG1 = "${fileBaseName}_$(Get-Date -Format HHmmss)"  #temporary unique filename
            
            # use pyton library rembg to replace background with solid white...
            & rembg i -m u2net -bgc 255 255 255 255 -a -ae 5 $FilePath $dataOut/$TMPIMG1.png

            #replaces transparent bg with solid white...
            & $WorkDir\ImageMagick\magick.exe $dataout\$TMPIMG1.png -background white -alpha remove -alpha off $dataout/$fileName 
            
            #center produced image...
            & $WorkDir\ImageMagick\magick.exe $dataout\$TMPIMG1.png -crop "${CoordXY}x${CoordXY}+$OriginLeft+$OriginTop" +repage -gravity center -extent "${CoordXY}x${CoordXY}" "$dataout/$fileName"
            
            #produce 250x250 pixel image in $passports directory...
            & $WorkDir\ImageMagick\magick.exe $dataout\$fileName -resize 250x250 $passports/$fileName
            
            #produce example of eventual profile image...
            & $WorkDir\ImageMagick\magick.exe $passports\$fileName `
            "(" +clone -threshold 101% -fill white -draw "circle 125,125 125,0" ")" `
            -alpha off -compose copy_opacity -composite -background "#F1F3F4" -alpha remove -alpha off `
            $vignettes\${FileBaseName}_vignette.png

            remove-item $dataout\$TMPIMG1.png -force -verbose # delete temp file
                
            Start-Sleep 1
            DashedLine
        } else {
            Write-Warning "No Face detected for file: $filePath"
    }
}

<#
Write-Host "Cleaning up temporary files..."
remove-item $dataout\*.csv -force -verbose
remove-item $dataout\detected*.* -force -verbose
#>

Stop-Transcript
