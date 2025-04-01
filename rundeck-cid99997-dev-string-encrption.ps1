$mnspver = "0.0.1"

Write-Host $(Get-Date)
Write-Host "MNSP Version" $mnspver
#Start-Sleep 10
$ErrorActionPreference="Continue"
Set-Location $GamDir

function DashedLine {
Write-host "-----------------------------------------------------------`n"
}


function Set-Key {
param([string]$string)
$length = $string.length
$pad = 32-$length
if (($length -lt 16) -or ($length -gt 32)) {Throw "String must be between 16 and 32 characters"}
#a key of 128 bits can be specified as a byte array of 16 decimal numerals. Similarly, 192-bit and 256-bit keys correspond to byte arrays of 24 and 32 decimal numerals, respectively.
$encoding = New-Object System.Text.ASCIIEncoding
$bytes = $encoding.GetBytes($string + "0" * $pad)
return $bytes
}

function Set-EncryptedData {
param($key,[string]$plainText)
$securestring = new-object System.Security.SecureString
$chars = $plainText.toCharArray()
foreach ($char in $chars) {$secureString.AppendChar($char)}
$encryptedData = ConvertFrom-SecureString -SecureString $secureString -Key $key
return $encryptedData
}

function Get-EncryptedData {
param($key,$data)
$data | ConvertTo-SecureString -key $key |
ForEach-Object {[Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($_))}
}

$keyinput = Read-Host -Prompt 'Key...' #prompt user for encrpytion key
$key = Set-Key $keyinput
Write-Host "Thanks..."

    $pwd = $(Invoke-WebRequest -Uri "$webPWDSourceURL")
#    $pwd.Content
    #$pwd.StatusCode
        if ($pwd.StatusCode -eq 200) {
        #proceed with pwd reservation

        $toMatch = $($pwd.Content)
        $plainText = $toMatch

                Write-Host "Plain text pwd : " $plainText
                $encryptedText = Set-EncryptedData -key $key -plainText $plaintext
                Write-Host "encrypted      : " $encryptedText

                $DecryptedText = Get-EncryptedData -data $encryptedText -key $key
                Write-Host "decrypted      : " $DecryptedText

                if ($plainText -eq $DecryptedText) { #confirm encrypt/decrypted strings match
                Write-Host "Match" 

                filepath $sheetcsv -Append


                }
        }



=VLOOKUP(C3,'7. MIS Import'!A1:C24,3)

=IF(AND(C3='7. MIS Import'!A2,D3='7. MIS Import'!B2,"Match"))
=IF(AND(C3='7. MIS Import'!A2,D3='7. MIS Import'!B2),"Match","no match")

=IF(AND(K3=I3,J3=L3),"Match","no match")

=QUERY({IMPORTRANGE("https://docs.google.com/spreadsheets/d/1pzJclT8TWNvek-SFQzlb8MdQIXN7BmIMOaW7y9hCVCU","'7. MIS Import'!B1:B") ; IMPORTRANGE("https://docs.google.com/spreadsheets/d/1pzJclT8TWNvek-SFQzlb8MdQIXN7BmIMOaW7y9hCVCU","'7. MIS Import'!C2:C")}, "select Col1 where not(Col1=")")"

={FILTER (IMPORTRANGE("https://docs.google.com/spreadsheets/d/1pzJclT8TWNvek-SFQzlb8MdQIXN7BmIMOaW7y9hCVCU","'7. MIS Import'!A2:A"),IMPORTRANGE("https://docs.google.com/spreadsheets/d/1pzJclT8TWNvek-SFQzlb8MdQIXN7BmIMOaW7y9hCVCU","'7. MIS Import'!A2:A") <> "") ; IMPORTRANGE("https://docs.google.com/spreadsheets/d/1pzJclT8TWNvek-SFQzlb8MdQIXN7BmIMOaW7y9hCVCU","'7. MIS Import'!B2:B"),IMPORTRANGE("https://docs.google.com/spreadsheets/d/1pzJclT8TWNvek-SFQzlb8MdQIXN7BmIMOaW7y9hCVCU","'7. MIS Import'!B2:B") <> ""}

=CONCATENATE(ARRAYFORMULA('7. MIS Import'!A2:A)," ",ARRAYFORMULA ('7. MIS Import'!B2:B))