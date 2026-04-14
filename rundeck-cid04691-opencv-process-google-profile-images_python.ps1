$mnspver = "0.0.26_6"
Clear-Host

function DashedLine {
Write-host "-----------------------------------------------------------`n"
}

$workdir = "$HOME/Documents/opencv-dev" #linux dev environment
$LogDir = "$workdir/Logs"
$now = $(Get-date -Format yyyyMMdd-HHmmss)
$transcriptlog = "$LogDir\$(Get-date -Format yyyyMMdd-HHmmss)_transcript.log"
$datadir = "$workdir/Data"

$datasrc = "$dataDir/source4/"
$dataout = "$datadir/output/$now"
$passports = "$dataout/250x250"
$photosSrc = $(Get-ChildItem -Path $datasrc ) # | Select-Object -ExpandProperty FullName)
$ImgDimensions = @()
$ImgDimensionX = @()
$ImgDimensionY = @()
$convertX = "x"
Start-Transcript -Path $transcriptlog -Force -NoClobber -Append
Write-Host "MNSP version:" $mnspver

New-item -Path $dataout -ItemType Directory
New-item -Path $passports -ItemType Directory
DashedLine

foreach ($photo in $photosSrc) {
    Write-Host "Processing Image Details:"
    $filePath = $photo.FullName
    $fileName = $photo.name
    $fileBaseName = $photo.BaseName
    
    Write-Host "FilePath: $filePath"
    Write-Host "FileName: $fileName"
    Write-Host "BaseName: $fileBaseName `n"
    
    $ImgDimensionX, $ImgDimensionY = ( & identify -format "%w,%h" $filePath).Split(',') #get image dimensions
    $ImgEXIFDateTimeOriginal = (& identify -format "%[EXIF:DateTimeOriginal]" $filePath) #get image creation date using EXIF

    # attempt detection using facedetect
    $imgCoordinates = ( & facedetect $filePath --best) # get bounding box of face

    # if facedetect fails, trigger python fallback
        if ([string]::IsNullOrWhiteSpace($imgCoordinates)) {
                Write-Warning "No coordinates found for image: $filePath" #no face detected in source image
                Write-Host "try alternative python method..."
                & python3 $workDir/cid04691_01.py $filePath $dataOut #update to use Variable(s) [TODO]
                #set coordinates from python processing... [TODO]
                $pythonCoords = import-csv -path $dataout/face_metadata.csv #update to use Variable(s) [TODO]
                Write-Host "Python Library coordinates..."
                $pythonCoords

                DashedLine
                continue 
        }

    Write-Host "Face detected by facedetect Processing image..."
    $imgCoordinatesCentre = ( & facedetect $filePath --best -c) # get center of image co-ordinates
    
    #if ($null -ne $imgCoordinates) { #check for detected face in source image
        
        Write-Host "Image Dimension X     :" $ImgDimensionX
        Write-Host "Image Dimension Y     :" $ImgDimensionY
        Write-Host "EXIF DateTimeOriginal :" $ImgEXIFDateTimeOriginal
        Write-Host "Derived Face coords   :" $imgCoordinates
        Write-Host "Derived face Center   :" $imgCoordinatesCentre
        
        #unit - detected bounding box divided in half
        $unit = $([Math]::Round($imgCoordinates.Split(" ")[3]/2)) # divide one of square values in half, rounded up to next whole number
        Write-Host "Square Unit           :" $unit
        
        $OriginTop = $($imgCoordinates.Split(" ")[1] -$unit) #new Horizontal Origin Value
        Write-Host "New Horizontal Origin :" $OriginTop
        
        $OriginLeft = $($imgCoordinates.Split(" ")[0] -$unit) #minus full unit
        Write-Host "New Vertical Origin  :" $OriginLeft
        
        $Coords = $($imgCoordinates.Split(" ")[2] )
        
        $CoordXY = ([int]$Coords * 2)
        
        Write-Host "New XY Coordinate     :" $CoordXY
        Write-Host "Output Image: " $dataout/$fileName

        #NOTE: can fail to give 1:1 ratio image under some source image secnarios...
        #& convert $filePath -crop $CoordXY$convertX$CoordXY+$OriginLeft+$OriginTop $dataout/$fileName

        #resolves issue if detrmined co-ordinates are out of range of source image - not 1:1 ratio:
        & convert $filePath -crop "${CoordXY}x${CoordXY}+$OriginLeft+$OriginTop" +repage -gravity center -background white -extent "${CoordXY}x${CoordXY}" "$dataout/$fileName"

        $TMPIMG1 = "${fileBaseName}_$(Get-Date -Format HHmmss)"  #temporary unique filename
        & rembg i $dataout/$fileName $dataout/$TMPIMG1.png # use pyton library rembg to remove background 

        & convert $dataout/$TMPIMG1.png -background white -alpha remove -alpha off $dataout/$fileName #replaces transparent bg with solid white

        & convert $dataout/$fileName -resize 250x250 $passports/$fileName #produce 250x250 pixel image in $passports directory
        
        remove-item $dataout/$TMPIMG1.png -force -verbose # delete temp file
        
        Start-Sleep 1

    #    }
    #    else {

    DashedLine
}

Stop-Transcript

<#
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
#>