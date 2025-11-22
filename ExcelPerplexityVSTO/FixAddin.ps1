# Fix Excel Perplexity VSTO Add-in
# This script resets LoadBehavior and ensures proper trust

$ErrorActionPreference = "Stop"

Write-Host "Excel Perplexity VSTO Add-in Fix" -ForegroundColor Cyan
Write-Host "=================================" -ForegroundColor Cyan
Write-Host ""

# Check if add-in is registered
$registryPath = "HKCU:\Software\Microsoft\Office\Excel\Addins\ExcelPerplexityVSTO"

if (-not (Test-Path $registryPath)) {
    Write-Host "ERROR: Add-in is not registered!" -ForegroundColor Red
    Write-Host "Please run Deploy.ps1 first." -ForegroundColor Yellow
    exit 1
}

# Get current LoadBehavior
$loadBehavior = (Get-ItemProperty $registryPath).LoadBehavior

Write-Host "Current LoadBehavior: $loadBehavior" -ForegroundColor Yellow

if ($loadBehavior -eq 2) {
    Write-Host "LoadBehavior is 2 - Excel disabled the add-in due to an error." -ForegroundColor Red
    Write-Host ""
    Write-Host "Common causes:" -ForegroundColor Yellow
    Write-Host "  1. Certificate not trusted" -ForegroundColor White
    Write-Host "  2. VSTO Runtime not installed" -ForegroundColor White
    Write-Host "  3. Missing dependencies" -ForegroundColor White
    Write-Host "  4. Code error in ThisAddIn_Startup" -ForegroundColor White
    Write-Host ""
}

# Reset LoadBehavior to 3
Write-Host "Resetting LoadBehavior to 3 (Load on startup)..." -ForegroundColor Yellow
Set-ItemProperty -Path $registryPath -Name "LoadBehavior" -Value 3 -Type DWord
Write-Host "LoadBehavior reset to 3" -ForegroundColor Green

# Trust the certificate
Write-Host ""
Write-Host "Ensuring certificate is trusted..." -ForegroundColor Yellow
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$certPath = Join-Path $scriptPath "ExcelPerplexityVSTO\ExcelPerplexityVSTO_TemporaryKey.pfx"

if (Test-Path $certPath) {
    try {
        # Import certificate to Trusted Publishers
        $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
        $cert.Import($certPath, "password", [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::UserKeySet)

        $store = New-Object System.Security.Cryptography.X509Certificates.X509Store([System.Security.Cryptography.X509Certificates.StoreName]::TrustedPublisher, [System.Security.Cryptography.X509Certificates.StoreLocation]::CurrentUser)
        $store.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite)

        # Check if already trusted
        $existing = $store.Certificates | Where-Object { $_.Thumbprint -eq $cert.Thumbprint }

        if ($existing) {
            Write-Host "Certificate already trusted (Thumbprint: $($cert.Thumbprint))" -ForegroundColor Green
        } else {
            $store.Add($cert)
            Write-Host "Certificate added to Trusted Publishers (Thumbprint: $($cert.Thumbprint))" -ForegroundColor Green
        }

        $store.Close()
    } catch {
        Write-Host "WARNING: Could not trust certificate automatically" -ForegroundColor Yellow
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host ""
        Write-Host "Please trust manually:" -ForegroundColor Yellow
        Write-Host "  1. Open certmgr.msc" -ForegroundColor White
        Write-Host "  2. Import $certPath to Trusted Publishers" -ForegroundColor White
        Write-Host "  3. Password: password" -ForegroundColor White
    }
} else {
    Write-Host "WARNING: Certificate file not found at: $certPath" -ForegroundColor Yellow
}

# Check VSTO Runtime
Write-Host ""
Write-Host "Checking VSTO Runtime..." -ForegroundColor Yellow
$vstoRuntime = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\VSTO Runtime Setup\v4R" -ErrorAction SilentlyContinue
if ($null -eq $vstoRuntime) {
    Write-Host "WARNING: VSTO Runtime may not be installed!" -ForegroundColor Red
    Write-Host "Download from: https://www.microsoft.com/en-us/download/details.aspx?id=105522" -ForegroundColor Yellow
    Write-Host "Latest version: Visual Studio 2010 Tools for Office Runtime (v10.0.60917)" -ForegroundColor White
} else {
    Write-Host "VSTO Runtime found: Version $($vstoRuntime.Version)" -ForegroundColor Green
}

# Check dependencies
Write-Host ""
Write-Host "Checking dependencies..." -ForegroundColor Yellow
$manifestPath = (Get-ItemProperty $registryPath).Manifest
$manifestPath = $manifestPath -replace '\|vstolocal$', ''
$binPath = Split-Path -Parent $manifestPath

$requiredFiles = @(
    "ExcelPerplexityVSTO.dll",
    "ExcelPerplexityVSTO.dll.manifest",
    "ExcelPerplexityVSTO.vsto",
    "Newtonsoft.Json.dll"
)

$allFilesExist = $true
foreach ($file in $requiredFiles) {
    $filePath = Join-Path $binPath $file
    if (Test-Path $filePath) {
        Write-Host "  [OK] $file" -ForegroundColor Green
    } else {
        Write-Host "  [MISSING] $file" -ForegroundColor Red
        $allFilesExist = $false
    }
}

if (-not $allFilesExist) {
    Write-Host ""
    Write-Host "ERROR: Missing dependencies! Please rebuild the project:" -ForegroundColor Red
    Write-Host "  msbuild ExcelPerplexityVSTO.sln /p:Configuration=Release" -ForegroundColor Yellow
}

# Enable VSTO logging for diagnostics
Write-Host ""
Write-Host "Enabling VSTO diagnostic logging..." -ForegroundColor Yellow
$vstoSecurityPath = "HKCU:\Software\Microsoft\VSTO\Security"
if (-not (Test-Path $vstoSecurityPath)) {
    New-Item -Path $vstoSecurityPath -Force | Out-Null
}
Set-ItemProperty -Path $vstoSecurityPath -Name "LoadBehavior" -Value 2 -Type DWord
Write-Host "VSTO logging enabled. Logs will be in: $env:TEMP\VSTO\" -ForegroundColor Green

Write-Host ""
Write-Host "=================================" -ForegroundColor Cyan
Write-Host "Fix completed!" -ForegroundColor Green
Write-Host ""
Write-Host "Next steps:" -ForegroundColor Yellow
Write-Host "1. Close ALL Excel instances (check Task Manager)" -ForegroundColor White
Write-Host "2. Open Excel" -ForegroundColor White
Write-Host "3. Check File > Options > Add-ins > COM Add-ins" -ForegroundColor White
Write-Host "4. If LoadBehavior changes back to 2, check:" -ForegroundColor White
Write-Host "   - Event Viewer: Windows Logs > Application" -ForegroundColor White
Write-Host "   - VSTO Logs: $env:TEMP\VSTO\" -ForegroundColor White
Write-Host ""
Write-Host "If the problem persists, there may be a code error." -ForegroundColor Yellow
Write-Host "Try running from Visual Studio (F5) for better error messages." -ForegroundColor Yellow
