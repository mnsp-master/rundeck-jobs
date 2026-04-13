Clear-Host

function DashedLine {
Write-host "-----------------------------------------------------------`n"
}

$mnspver = "0.0.4"
$workdir = "$HOME/Documents/opencv-dev" #linux dev environment
$LogDir = "$workdir/Logs"
$now = $(Get-date -Format yyyyMMdd-HHmmss)
$transcriptlog = "$LogDir\$(Get-date -Format yyyyMMdd-HHmmss)_transcript.log"
$datadir = "$workdir/Data"
#$datasrc = "$workdir/Data/source/dev"
$datasrc = "$dataDir/source3/"
$dataout = "$datadir/output/$now"
$passports = "$dataout/250x250"
$photosSrc = $(Get-ChildItem -Path $datasrc ) # | Select-Object -ExpandProperty FullName)
$ImgDimensions = @()
$ImgDimensionX = @()
$ImgDimensionY = @()
$convertX = "x"
Start-Transcript -Path $transcriptlog -Force -NoClobber -Append



New-item -Path $dataout -ItemType Directory
New-item -Path $passports -ItemType Directory
DashedLine


#$photosSrc = Import-Csv -Path $datadir/data.csv
Clear-Host

foreach ($photo in $photosSrc) {
    Write-Host "Processing Image: " $photo
    $filePath = $photo.FullName
    $fileName = $photo.name
    
    #$ImgDimensions = $((Invoke-Expression "identify $filePath").Split(" ")[2])
    #$ImgDimensionX = $($ImgDimensions.Split("x")[0])
    #$ImgDimensionY = $($ImgDimensions.Split("x")[1])

    $ImgDimensionX, $ImgDimensionY = ( & identify -format "%w,%h" $filePath).Split(',') #get image dimensions
    $ImgEXIFDateTimeOriginal = (& identify -format "%[EXIF:DateTimeOriginal]" $filePath) #get image creation date

    #$imgCoordinates = $(Invoke-Expression "facedetect $filePath --best")
    $imgCoordinates = ( & facedetect $filePath --best)

    #$imgCoordinatesCentre = $(Invoke-Expression "facedetect $filePath --best -c")
    $imgCoordinatesCentre = ( & facedetect $filePath --best -c)
    
    if ($null -ne $imgCoordinates) { #check for detected face in source image
        
        #Write-Host "Image Dimensions     :" $ImgDimensions
        Write-Host "Image Dimension X     :" $ImgDimensionX
        Write-Host "Image Dimension Y     :" $ImgDimensionY
        Write-Host "EXIF DateTimeOriginal :" $ImgEXIFDateTimeOriginal
        Write-Host "Derived Face coords   :" $imgCoordinates
        Write-Host "Derived face Center   :" $imgCoordinatesCentre
        #$unit = $($imgCoordinates.Split(" ")[3]/2) # divide one of square values in half
        $unit = $([Math]::Round($imgCoordinates.Split(" ")[3]/2)) # divide one of square values in half, rounded up to next whole number
        Write-Host "Square Unit           :" $unit
        $OriginTop = $($imgCoordinates.Split(" ")[1] -$unit) #new Horizontal Origin Value
        Write-Host "New Horizontal Origin :" $OriginTop
        #$OriginLeft = $($imgCoordinates.Split(" ")[0] -$unit/2)#minus half unit
        $OriginLeft = $($imgCoordinates.Split(" ")[0] -$unit)#minus full unit
        Write-Host "New Vertical Origin  :" $OriginLeft
        $Coords = $($imgCoordinates.Split(" ")[2] )
        #$Coords
        $CoordXY = ([int]$Coords * 2) 
        #$CoordXY
        #$CoordX = $($OriginLeft + ($unit * 4))
        Write-Host "New XY Coordinate     :" $CoordXY
        #$CoordY = $($OriginTop + ($unit * 4))
        #Write-Host "New Y Coordinate     :" $CoordY
        Write-Host "Output Image: " $dataout/$fileName

        #Invoke-expression "convert $filePath -crop $CoordXY$convertX$CoordXY+$OriginLeft+$OriginTop $dataout/$fileName"
        #& convert $filePath -crop $CoordXY$convertX$CoordXY+$OriginLeft+$OriginTop $dataout/$fileName
        
        #force - not ideal
        & convert $filePath -crop "${CoordXY}x${CoordXY}+$OriginLeft+$OriginTop" +repage -gravity center -background white -extent "${CoordXY}x${CoordXY}" "$dataout/$fileName"
        $IMG = $(Get-Date -Format HHmmss)
        & rembg i $dataout/$fileName $dataout/$IMG.png # use pyton library to remove background 
        
        & convert $dataout/$fileName -resize 250x250 $passports/$fileName
        
        #Invoke-Expression "convert $dataout/$fileName -resize 250x250 $passports/$fileName"
        Start-Sleep 1

        }
        else {Write-Warning "No coordinates found for image: $filePath"} #no face detected in source image
    DashedLine
}



Stop-Transcript

