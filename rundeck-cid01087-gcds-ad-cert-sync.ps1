<#
  .SYNOPSIS
  Exports the bound AD LDAPS encryption certificate and imports it into
  Google Cloud Directory Sync's Java keystore, so that GCDS syncs will succeed.

  .DESCRIPTION
  Google Cloud Directory Sync runs on Java. Java maintains its own trusted keystore,
  separate from the host operating system. Often, this keystore grows stale when updates
  are neglected. Further, the keystore would never contain certificate information for
  self-signed or internally-distributed certificates.

  In order to make GCDS work with TLS using secure LDAPS binding, it is necessary to
  export your trusted certificate from the machine's certificate store and import it into
  the GCDS-bundled Java Runtime Environment's certificate store.

  Given a ComputerName and Port, this script will connect to the named DC and determine the
  thumbprint of the certificate bound to the DC on the specific port.

  Using this thumbprint, the script then exports the certificate from the Local Computer's MY (Personal)
  certificate store. This does NOT include the private key, and therefore it's safe to do this.

  Next, the script deletes and re-imports the certificate into the JRE certificate store.

  .PARAMETER ComputerName
  Use the fully-qualified network name of the machine. We're assuming this is the same network name
  that will be used in GCDS to bind against the DC, and is also the CommonName represented in the certificate.

  .PARAMETER Port
  Usually this will be 636, but could be custom depending on your environment.

  .OUTPUTS
  Will list the thumbprint of the cert found and will show stderr and stdout of the keytool commands.
  Error handling could definitely be beefed up here.

  .EXAMPLE
  C:\PS> .\Update-JavaDomainControllerCertificate.ps1 -ComputerName my.domain.com -Port 636

#>

[CmdletBinding()]
param (
    [Parameter(Mandatory=$true)]
    [string]
    $ComputerName,

    [int]
    $Port = 636
)
$FilePath = "$($Env:TEMP)\adcert.crt"
$Certificate = $null
$TcpClient = New-Object -TypeName System.Net.Sockets.TcpClient
try {

    $TcpClient.Connect($ComputerName, $Port)
    $TcpStream = $TcpClient.GetStream()

    $Callback = { param($sender, $cert, $chain, $errors) return $true }

    $SslStream = New-Object -TypeName System.Net.Security.SslStream -ArgumentList @($TcpStream, $true, $Callback)
    try {

        $SslStream.AuthenticateAsClient('')
        $Certificate = $SslStream.RemoteCertificate

    } finally {
        $SslStream.Dispose()
    }

} finally {
    $TcpClient.Dispose()
}

if ($Certificate) {
    if ($Certificate -isnot [System.Security.Cryptography.X509Certificates.X509Certificate2]) {
        $Certificate = New-Object -TypeName System.Security.Cryptography.X509Certificates.X509Certificate2 -ArgumentList $Certificate
    }
    Write-Output "Found Certificate:"
    Write-Output $Certificate
}

Export-Certificate -Cert $Certificate -Force -FilePath $FilePath | Out-Null

Set-Location -Path "C:\Program Files\Google Cloud Directory Sync\jre"

# Delete existing entry
& .\bin\keytool -keystore lib\security\cacerts -storepass changeit -delete -noprompt -alias $ComputerName 2>&1 | %{ "$_" }

# Add entry
& .\bin\keytool -keystore lib\security\cacerts -storepass changeit -importcert -noprompt -file $FilePath -alias $ComputerName 2>&1 | %{ "$_" }

Remove-Item -Path $FilePath -Force