<#
.SYNOPSIS
Ensures a self-signed development code signing certificate exists and prints its thumbprint.

.DESCRIPTION
Called by the EnsureDevSigningCertificate MSBuild target when Navferty.pfx is absent.
The certificate is kept in CurrentUser\My and exported to the given pfx path (needed by the VSTO signing targets).
The only output is the thumbprint. The certificate is for local builds and CI only and is not trusted by end users.
#>
param(
    [Parameter(Mandatory)][string]$PfxPath
)

$ErrorActionPreference = 'Stop'
$subject = 'CN=Navferty Dev Certificate'

$cert = Get-ChildItem Cert:\CurrentUser\My |
    Where-Object { $_.Subject -eq $subject -and $_.HasPrivateKey -and $_.NotAfter -gt (Get-Date).AddDays(7) } |
    Sort-Object NotAfter -Descending |
    Select-Object -First 1

$export = -not (Test-Path $PfxPath)

if (-not $cert) {
    $cert = New-SelfSignedCertificate -Type CodeSigningCert -Subject $subject `
        -CertStoreLocation Cert:\CurrentUser\My -NotAfter (Get-Date).AddYears(2)
    $export = $true
}

if ($export) {
    New-Item -ItemType Directory -Force -Path (Split-Path $PfxPath) | Out-Null
    $password = ConvertTo-SecureString ([guid]::NewGuid().ToString()) -AsPlainText -Force
    Export-PfxCertificate -Cert $cert -FilePath $PfxPath -Password $password | Out-Null
}

Write-Output $cert.Thumbprint
