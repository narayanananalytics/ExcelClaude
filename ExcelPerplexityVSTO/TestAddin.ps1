# Test Excel Perplexity VSTO Add-in
# This script resets LoadBehavior, closes Excel, and opens it for testing

$ErrorActionPreference = "Stop"

Write-Host "Excel Perplexity VSTO Add-in Test" -ForegroundColor Cyan
Write-Host "===================================" -ForegroundColor Cyan
Write-Host ""

# Reset LoadBehavior to 3
$registryPath = "HKCU:\Software\Microsoft\Office\Excel\Addins\ExcelPerplexityVSTO"
Write-Host "Resetting LoadBehavior to 3..." -ForegroundColor Yellow
Set-ItemProperty -Path $registryPath -Name "LoadBehavior" -Value 3 -Type DWord
Write-Host "LoadBehavior set to 3" -ForegroundColor Green

# Close all Excel instances
Write-Host ""
Write-Host "Closing all Excel instances..." -ForegroundColor Yellow
$excelProcesses = Get-Process excel -ErrorAction SilentlyContinue
if ($excelProcesses) {
    $excelProcesses | Stop-Process -Force
    Start-Sleep -Seconds 2
    Write-Host "Excel closed" -ForegroundColor Green
} else {
    Write-Host "No Excel instances running" -ForegroundColor Green
}

# Wait a moment
Write-Host ""
Write-Host "Waiting 3 seconds..." -ForegroundColor Yellow
Start-Sleep -Seconds 3

# Open Excel
Write-Host ""
Write-Host "Opening Excel..." -ForegroundColor Yellow
Start-Process "excel.exe"

Write-Host ""
Write-Host "===================================" -ForegroundColor Cyan
Write-Host "Excel is starting..." -ForegroundColor Green
Write-Host ""
Write-Host "Check the following:" -ForegroundColor Yellow
Write-Host "1. Does Excel open without errors?" -ForegroundColor White
Write-Host "2. In Excel, go to: File > Options > Add-ins" -ForegroundColor White
Write-Host "3. Manage: COM Add-ins > Go..." -ForegroundColor White
Write-Host "4. Is 'Excel Perplexity VSTO Add-in' listed and CHECKED?" -ForegroundColor White
Write-Host ""
Write-Host "If LoadBehavior changes back to 2:" -ForegroundColor Yellow
Write-Host "  - There's still a code error" -ForegroundColor White
Write-Host "  - Use Visual Studio F5 to see the exact error" -ForegroundColor White
Write-Host ""
Write-Host "To check LoadBehavior:" -ForegroundColor Cyan
Write-Host '  Get-ItemProperty "HKCU:\Software\Microsoft\Office\Excel\Addins\ExcelPerplexityVSTO"' -ForegroundColor Gray
