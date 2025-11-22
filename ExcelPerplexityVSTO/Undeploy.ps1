# Excel Perplexity VSTO Add-in Undeployment Script
# This script unregisters the add-in for the current user

$ErrorActionPreference = "Stop"

Write-Host "Excel Perplexity VSTO Add-in Undeployment" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host ""

# Remove registry key
$registryPath = "HKCU:\Software\Microsoft\Office\Excel\Addins\ExcelPerplexityVSTO"

if (Test-Path $registryPath) {
    Write-Host "Removing registry entries..." -ForegroundColor Yellow
    Remove-Item $registryPath -Force -Recurse
    Write-Host "Registry entries removed successfully!" -ForegroundColor Green
} else {
    Write-Host "Add-in is not registered." -ForegroundColor Yellow
}

# Remove certificate from trusted publishers
Write-Host ""
Write-Host "Removing certificate from Trusted Publishers..." -ForegroundColor Yellow
try {
    $store = New-Object System.Security.Cryptography.X509Certificates.X509Store([System.Security.Cryptography.X509Certificates.StoreName]::TrustedPublisher, [System.Security.Cryptography.X509Certificates.StoreLocation]::CurrentUser)
    $store.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite)

    $certs = $store.Certificates | Where-Object { $_.Subject -eq "CN=ExcelPerplexityVSTO" }
    foreach ($cert in $certs) {
        $store.Remove($cert)
        Write-Host "Removed certificate: $($cert.Thumbprint)" -ForegroundColor Green
    }

    $store.Close()
} catch {
    Write-Host "WARNING: Could not remove certificate automatically." -ForegroundColor Yellow
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ""
Write-Host "=========================================" -ForegroundColor Cyan
Write-Host "Undeployment completed!" -ForegroundColor Green
Write-Host ""
Write-Host "Please restart Excel for changes to take effect." -ForegroundColor Yellow
