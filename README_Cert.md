# Code signing certificates

## Local development

No setup is needed. VSTO requires signed manifests, so when `NavfertyExcelAddIn/Navferty.pfx` is absent the build
runs `scripts/Ensure-DevCertificate.ps1` (via the `EnsureDevSigningCertificate` target in the csproj). It creates a
self-signed `CN=Navferty Dev Certificate` in your `CurrentUser\My` store and signs with it. CI uses the same path.

The development certificate is not trusted by end users. Real releases are signed in the `production` GitHub
environment from the `CERT_PFX_BASE64` and `CERT_PASSWORD` secrets. `Navferty.pfx` is git-ignored and must not be committed.

## Release certificate

### Generate self-signed code signing certificate

Use certreq so that the certificate has no EKU (Enhanced Key Usage) restrictions, matching the Visual Studio generated certificate.

Fill the password in the script before running.

Do not forget to update the .csproj file with the new thumbprint after generating the certificate.

```powershell
# Clean up and start fresh
Remove-Item navferty.* -Force -ErrorAction SilentlyContinue

# Create certificate request INF file
$infContent = @"
[Version]
Signature = "`$Windows NT$"

[NewRequest]
Subject = "CN=Navferty"
KeySpec = 1
KeyLength = 2048
Exportable = TRUE
MachineKeySet = FALSE
ProviderName = "Microsoft Enhanced RSA and AES Cryptographic Provider"
ProviderType = 24
RequestType = Cert
HashAlgorithm = SHA256
ValidityPeriod = Years
ValidityPeriodUnits = 5

[Extensions]
; Intentionally empty - no EKU
"@

# Save INF file
$infPath = "C:\Temp\cert-test\navferty.inf"
$cerPath = "C:\Temp\cert-test\navferty.cer"
Set-Content -Path $infPath -Value $infContent

Write-Host "Creating certificate with certreq..."
certreq -new -f $infPath $cerPath

if (Test-Path $cerPath) {
    Write-Host "✓ Certificate created: $cerPath"
    
    # Import to certificate store
    $cert = Import-Certificate -FilePath $cerPath -CertStoreLocation Cert:\CurrentUser\My
    
    Write-Host "✓ Certificate imported to store"
    Write-Host "  Thumbprint: $($cert.Thumbprint)"
    Write-Host "  Subject: $($cert.Subject)"
    
    # Check EKU
    Write-Host "`nEnhanced Key Usage:"
    if ($cert.EnhancedKeyUsageList.Count -eq 0) {
        Write-Host "  ✓ NONE (unrestricted - matches VS certificate)"
        $ekuStatus = "SUCCESS"
    } else {
        Write-Host "  ✗ HAS EKU:"
        $cert.EnhancedKeyUsageList | Format-Table FriendlyName, ObjectId
        $ekuStatus = "FAILED"
    }
    
    # Export as PFX
    $password = ConvertTo-SecureString -String "password-here" -Force -AsPlainText
    $pfxPath = "C:\repos\NavfertyExcelAddIn\NavfertyExcelAddIn\Navferty.pfx"
    
    Export-PfxCertificate -Cert $cert -FilePath $pfxPath -Password $password
    
    Write-Host "`n=== RESULT ==="
    Write-Host "Status: $ekuStatus"
    Write-Host "PFX Location: $pfxPath"
    Write-Host "Thumbprint: $($cert.Thumbprint)"
    Write-Host "`nNext steps:"
    Write-Host "1. Update .csproj: <ManifestCertificateThumbprint>$($cert.Thumbprint)</ManifestCertificateThumbprint>"
    Write-Host "2. Update the CERT_PASSWORD and CERT_PFX_BASE64 GitHub secrets"
    Write-Host "3. Do NOT commit Navferty.pfx (it is git-ignored)"
    
    # Cleanup temp files
    Remove-Item $infPath, $cerPath -Force
} else {
    Write-Host "✗ Certificate creation failed"
    Write-Host "certreq output should appear above"
}
```
