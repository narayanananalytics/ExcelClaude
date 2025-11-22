# Excel Perplexity VSTO Add-in Deployment Script
# This script registers the add-in for the current user

$ErrorActionPreference = "Stop"

Write-Host "Excel Perplexity VSTO Add-in Deployment" -ForegroundColor Cyan
Write-Host "=========================================" -ForegroundColor Cyan
Write-Host ""

# Get the manifest path
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$manifestPath = Join-Path $scriptPath "ExcelPerplexityVSTO\bin\Release\ExcelPerplexityVSTO.vsto"

# Check if manifest exists
if (-not (Test-Path $manifestPath)) {
    Write-Host "ERROR: Manifest file not found at: $manifestPath" -ForegroundColor Red
    Write-Host "Please build the project first using: msbuild ExcelPerplexityVSTO.sln /p:Configuration=Release" -ForegroundColor Yellow
    exit 1
}

Write-Host "Found manifest at: $manifestPath" -ForegroundColor Green

# Check for VSTO Runtime
Write-Host "Checking for Visual Studio Tools for Office Runtime..." -ForegroundColor Yellow
$vstoRuntime = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\VSTO Runtime Setup\v4R" -ErrorAction SilentlyContinue
if ($null -eq $vstoRuntime) {
    Write-Host "WARNING: VSTO Runtime may not be installed!" -ForegroundColor Yellow
    Write-Host "Download from: https://www.microsoft.com/en-us/download/details.aspx?id=105522" -ForegroundColor Yellow
} else {
    Write-Host "VSTO Runtime found: Version $($vstoRuntime.Version)" -ForegroundColor Green
}

# Create registry key
$registryPath = "HKCU:\Software\Microsoft\Office\Excel\Addins\ExcelPerplexityVSTO"
Write-Host ""
Write-Host "Creating registry entries..." -ForegroundColor Yellow

if (Test-Path $registryPath) {
    Write-Host "Removing existing registration..." -ForegroundColor Yellow
    Remove-Item $registryPath -Force
}

New-Item -Path $registryPath -Force | Out-Null
Set-ItemProperty -Path $registryPath -Name "Description" -Value "Perplexity AI Assistant for Excel with OHLC charting capabilities"
Set-ItemProperty -Path $registryPath -Name "FriendlyName" -Value "Excel Perplexity VSTO Add-in"
Set-ItemProperty -Path $registryPath -Name "LoadBehavior" -Value 3 -Type DWord
Set-ItemProperty -Path $registryPath -Name "Manifest" -Value "$manifestPath|vstolocal"

Write-Host "Registry entries created successfully!" -ForegroundColor Green

# Trust the certificate
Write-Host ""
Write-Host "Trusting the self-signed certificate..." -ForegroundColor Yellow
$certPath = Join-Path $scriptPath "ExcelPerplexityVSTO\ExcelPerplexityVSTO_TemporaryKey.pfx"
if (Test-Path $certPath) {
    try {
        $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
        $cert.Import($certPath, "password", [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::UserKeySet)

        $store = New-Object System.Security.Cryptography.X509Certificates.X509Store([System.Security.Cryptography.X509Certificates.StoreName]::TrustedPublisher, [System.Security.Cryptography.X509Certificates.StoreLocation]::CurrentUser)
        $store.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite)
        $store.Add($cert)
        $store.Close()

        Write-Host "Certificate trusted successfully!" -ForegroundColor Green
    } catch {
        Write-Host "WARNING: Could not trust certificate automatically. You may need to trust it manually." -ForegroundColor Yellow
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    }
} else {
    Write-Host "WARNING: Certificate file not found at: $certPath" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "=========================================" -ForegroundColor Cyan
Write-Host "Deployment completed successfully!" -ForegroundColor Green
Write-Host ""
Write-Host "Next steps:" -ForegroundColor Yellow
Write-Host "1. Close all Excel instances" -ForegroundColor White
Write-Host "2. Open Excel" -ForegroundColor White
Write-Host "3. Go to File > Options > Add-ins" -ForegroundColor White
Write-Host "4. Check that 'Excel Perplexity VSTO Add-in' is listed and enabled" -ForegroundColor White
Write-Host "5. Look for the custom ribbon tab in Excel" -ForegroundColor White
Write-Host ""
Write-Host "To uninstall, run: .\Undeploy.ps1" -ForegroundColor Cyan
