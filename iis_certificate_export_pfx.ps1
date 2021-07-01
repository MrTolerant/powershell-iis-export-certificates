$CERT_EXPORT_PATH = 'C:\iis_exported_certificates\'
$SITE_NAME = 'Default Web Site'

if (!(Test-Path $CERT_EXPORT_PATH )) { mkdir $CERT_EXPORT_PATH }

$certs = @()
$date = Get-Date

Get-ChildItem -Path IIS:\SslBindings | Where-Object { $_.Sites.Value.Contains($SITE_NAME) } | ForEach-Object -Process `
{
    $Password = ("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz".tochararray() | sort { Get-Random })[0..14] -join ''
    $secret = ConvertTo-SecureString -String $Password -Force -AsPlainText

    $certificate = Get-ChildItem -Recurse -Path cert:\localmachine\ | Where-Object -Property Thumbprint -EQ -Value $_.Thumbprint

    if ($certificate -and $_.host -and $certificate.NotAfter -gt $date) {
        $certName = $certificate.Subject.Split(', ')[0] -replace 'CN=', '' -replace '\*.', ''
        $hostname = $_.host
        $filePath = "$CERT_EXPORT_PATH$certName.pfx"
        $certificate | Export-PfxCertificate -FilePath $filePath -Password $secret | Out-Null
  
        $tmp = "$certName.pfx,$Password,$hostname"
        $certs += """$tmp"","
    }
}
$certs
