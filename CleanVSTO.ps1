# VSTO Add-in Cleanup Script (PowerShell)
# Run this as Administrator to completely remove VSTO add-in installations

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "VSTO Add-in Cleanup Script" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "This script will:" -ForegroundColor Yellow
Write-Host "1. Close Excel" -ForegroundColor Yellow
Write-Host "2. Clear ClickOnce cache" -ForegroundColor Yellow
Write-Host "3. Remove registry entries" -ForegroundColor Yellow
Write-Host "4. Clean deployment manifests" -ForegroundColor Yellow
Write-Host ""
Write-Host "Press Ctrl+C to cancel, or press any key to continue..." -ForegroundColor Yellow
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

# Step 1: Close Excel
Write-Host "`nStep 1: Closing Excel..." -ForegroundColor Green
try {
    $excelProcesses = Get-Process -Name EXCEL -ErrorAction SilentlyContinue
    if ($excelProcesses) {
        $excelProcesses | Stop-Process -Force
        Write-Host "✓ Excel closed successfully" -ForegroundColor Green
    } else {
        Write-Host "✓ Excel was not running" -ForegroundColor Green
    }
} catch {
    Write-Host "⚠ Could not close Excel: $_" -ForegroundColor Yellow
}

# Step 2: Clear ClickOnce cache
Write-Host "`nStep 2: Clearing ClickOnce cache..." -ForegroundColor Green
Write-Host "This may take a minute..." -ForegroundColor Gray

try {
    # Method 1: Use rundll32
    Start-Process "rundll32" -ArgumentList "dfshim CleanOnlineAppCache" -Wait -NoNewWindow
    Write-Host "✓ ClickOnce cache cleared (rundll32)" -ForegroundColor Green
} catch {
    Write-Host "⚠ Rundll32 method failed: $_" -ForegroundColor Yellow
}

# Method 2: Delete cache folders directly
$clickOncePath = "$env:LocalAppData\Apps\2.0"
if (Test-Path $clickOncePath) {
    try {
        Remove-Item -Path $clickOncePath -Recurse -Force -ErrorAction Stop
        Write-Host "✓ ClickOnce cache folders removed" -ForegroundColor Green
    } catch {
        Write-Host "⚠ Could not remove cache folders: $_" -ForegroundColor Yellow
    }
} else {
    Write-Host "✓ No ClickOnce cache folders found" -ForegroundColor Green
}

# Step 3: Remove registry entries
Write-Host "`nStep 3: Removing registry entries..." -ForegroundColor Green

$registryPaths = @(
    "HKCU:\Software\Microsoft\Office\Excel\Addins\ExcelPerplexityVSTO",
    "HKCU:\Software\Microsoft\VSTO"
)

foreach ($path in $registryPaths) {
    if (Test-Path $path) {
        try {
            # For VSTO path, only remove specific values related to our add-in
            if ($path -like "*VSTO*") {
                $securityPath = "$path\Security\Inclusion"
                if (Test-Path $securityPath) {
                    Get-ItemProperty -Path $securityPath | Get-Member -MemberType NoteProperty |
                    Where-Object { $_.Name -like "*ExcelPerplexityVSTO*" } |
                    ForEach-Object {
                        Remove-ItemProperty -Path $securityPath -Name $_.Name -ErrorAction SilentlyContinue
                    }
                }
            } else {
                Remove-Item -Path $path -Recurse -Force -ErrorAction Stop
            }
            Write-Host "✓ Removed registry entry: $path" -ForegroundColor Green
        } catch {
            Write-Host "⚠ Could not remove $path : $_" -ForegroundColor Yellow
        }
    }
}

# Step 4: Clean project output
Write-Host "`nStep 4: Cleaning project output..." -ForegroundColor Green

$projectPath = "C:\Users\naray\Documents\Projects\ExcelClaude\ExcelPerplexityVSTO\ExcelPerplexityVSTO"
$foldersToClean = @(
    "$projectPath\bin\Debug",
    "$projectPath\bin\Release",
    "$projectPath\obj\Debug",
    "$projectPath\obj\Release"
)

foreach ($folder in $foldersToClean) {
    if (Test-Path $folder) {
        try {
            Remove-Item -Path "$folder\*" -Recurse -Force -ErrorAction Stop
            Write-Host "✓ Cleaned: $folder" -ForegroundColor Green
        } catch {
            Write-Host "⚠ Could not clean $folder : $_" -ForegroundColor Yellow
        }
    }
}

# Summary
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "Cleanup Complete!" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Next steps:" -ForegroundColor Yellow
Write-Host "1. Open Visual Studio" -ForegroundColor White
Write-Host "2. Build → Clean Solution" -ForegroundColor White
Write-Host "3. Build → Rebuild Solution" -ForegroundColor White
Write-Host "4. Press F5 to debug" -ForegroundColor White
Write-Host ""
Write-Host "If you still have issues, restart your computer." -ForegroundColor Yellow
Write-Host ""
Write-Host "Press any key to exit..." -ForegroundColor Gray
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
