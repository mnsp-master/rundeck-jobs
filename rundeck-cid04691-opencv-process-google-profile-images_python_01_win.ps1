$mnspver = "0.0.26_19_16_13" #use python for all image coordinates
Clear-Host

function DashedLine {
Write-host "-----------------------------------------------------------`n"
}

$workdir = "$HOME\opencv-dev" #windows dev environment
$LogDir = "$workdir\Logs"
$now = $(Get-date -Format yyyyMMdd-HHmmss)
$transcriptlog = "$LogDir\$(Get-date -Format yyyyMMdd-HHmmss)_transcript.log"
$datadir = "$workdir\Data"

$datasrc = "$dataDir\source4\"
$dataout = "$datadir\output\$now"
$passports = "$dataout\250x250"
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

New-item -Path $dataout -ItemType Directory
New-item -Path $passports -ItemType Directory
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
            #check for csv (none produced if no faces detected)
            if (Test-path -path $dataout\$FileBaseName.csv) {
            $pythonCoords = import-csv -path $dataout\$FileBaseName.csv #update to use Variable(s) 
            Write-Host "Python Library coordinates..."
            $pythonCoords
            $unit = [Math]::Round($pythonCoords.CenterX - $pythonCoords.StartX)
                } else {
                Write-Host "No Face csv detected, moving to next image..."
                DashedLine
                continue
            }

        #only proceed if facial detection confidence is above 0.7 %
        $faceDetectionScore = $PythonCoords.confidence
        if ($faceDetectionScore -ge 0.7) {
            Write-Host "Face detected with confidence:" $faceDetectionScore
            Write-Host "Image Dimension X     :" $ImgDimensionX
            Write-Host "Image Dimension Y     :" $ImgDimensionY
            Write-Host "EXIF DateTimeOriginal :" $ImgEXIFDateTimeOriginal
            Write-Host "Square Unit           :" $unit
                
            #$OriginTop = $($imgCoordinates.Split(" ")[1] -$unit) #new Horizontal Origin Value
            $OriginTop = $($pythonCoords.CenterY - ($unit * 2)) #new Horizontal Origin Value
            Write-Host "New Horizontal Origin :" $OriginTop
                
            #$OriginLeft = $($imgCoordinates.Split(" ")[0] -$unit) #minus full unit
            $OriginLeft = $($pythonCoords.CenterX - ($unit *2))
            Write-Host "New Vertical Origin   :" $OriginLeft
                
            #$Coords = $($imgCoordinates.Split(" ")[2] )
            $Coords = $($unit * 2)
                
            #$CoordXY = ([int]$Coords * 2)
            $CoordXY = ($unit * 4)
                
            Write-Host "New XY Coordinate     :" $CoordXY
            Write-Host "Output Image          : " $dataout\$fileName

            if ($OriginTop -lt 0 -or $OriginLeft -lt 0) {
                Write-Warning "One or more Origin values: $OriginTop $OriginLeft are negative, consider providing a better image..."
            }
            
            #remove background from source image first...
            $TMPIMG1 = "${fileBaseName}_$(Get-Date -Format HHmmss)"  #temporary unique filename
            
            # use pyton library rembg to replace background with solid white
            & rembg i -m u2net -bgc 255 255 255 255 -a -ae 5 $FilePath $dataOut/$TMPIMG1.png

            #replaces transparent bg with solid white
            & $WorkDir\ImageMagick\magick.exe $dataout\$TMPIMG1.png -background white -alpha remove -alpha off $dataout/$fileName 
            
            #center produced image
            & $WorkDir\ImageMagick\magick.exe $dataout\$TMPIMG1.png -crop "${CoordXY}x${CoordXY}+$OriginLeft+$OriginTop" +repage -gravity center -extent "${CoordXY}x${CoordXY}" "$dataout/$fileName"
            
            #produce 250x250 pixel image in $passports directory
            & $WorkDir\ImageMagick\magick.exe $dataout\$fileName -resize 250x250 $passports/$fileName
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

<#
#& convert $dataout/$TMPIMG1.png -crop "${CoordXY}x${CoordXY}+$OriginLeft+$OriginTop" +repage -gravity center -background white -extent "${CoordXY}x${CoordXY}" "$dataout/$fileName"
#$ImgDimensions = $((Invoke-Expression "identify $filePath").Split(" ")[2])
    #$ImgDimensionX = $($ImgDimensions.Split("x")[0])
    #$ImgDimensionY = $($ImgDimensions.Split("x")[1])

    #$imgCoordinates = $(Invoke-Expression "facedetect $filePath --best")
    #$imgCoordinatesCentre = $(Invoke-Expression "facedetect $filePath --best -c")

    #$OriginLeft = $($imgCoordinates.Split(" ")[0] -$unit/2)#minus half unit

    #$CoordXY
        #$CoordX = $($OriginLeft + ($unit * 4))
        #$CoordY = $($OriginTop + ($unit * 4))
        #Write-Host "New Y Coordinate     :" $CoordY
    #Invoke-expression "convert $filePath -crop $CoordXY$convertX$CoordXY+$OriginLeft+$OriginTop $dataout/$fileName"
    #Invoke-Expression "convert $dataout/$fileName -resize 250x250 $passports/$fileName"

        #NOTE: can fail to give 1:1 ratio image under some source image secnarios...
                #& convert $filePath -crop $CoordXY$convertX$CoordXY+$OriginLeft+$OriginTop $dataout/$fileName

                <# original process
                #resolves issue if detrmined co-ordinates are out of range of source image - not 1:1 ratio:
                

                $TMPIMG1 = "${fileBaseName}_$(Get-Date -Format HHmmss)"  #temporary unique filename
                & rembg i $dataout/$fileName $dataout/$TMPIMG1.png # use pyton library rembg to remove background 

                & convert $dataout/$TMPIMG1.png -background white -alpha remove -alpha off $dataout/$fileName #replaces transparent bg with solid white

                & convert $dataout/$fileName -resize 250x250 $passports/$fileName #produce 250x250 pixel image in $passports directory
                
                remove-item $dataout/$TMPIMG1.png -force -verbose # delete temp file
                #>

#>