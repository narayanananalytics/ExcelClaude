# VSTO Add-in Troubleshooting Guide

## Error: "Another version is currently installed and cannot be upgraded"

This is the most common VSTO deployment error. Here are multiple solutions:

---

## 🚀 **Quick Solutions (Try These First)**

### **Solution 1: Automated Cleanup Script** ⭐ **RECOMMENDED**

**Using PowerShell (Recommended):**
1. Right-click on `CleanVSTO.ps1`
2. Select "Run with PowerShell"
3. If you get an execution policy error, run PowerShell as Administrator and execute:
   ```powershell
   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
   ```
4. Run the script again

**Using Batch File:**
1. Right-click on `CleanVSTO.bat`
2. Select "Run as Administrator"
3. Follow the prompts

**After running the script:**
1. Open Visual Studio
2. Build → Clean Solution
3. Build → Rebuild Solution
4. Press F5 to debug

---

### **Solution 2: Manual Uninstall**

1. **Close Excel completely**
   - Close all Excel windows
   - Open Task Manager (Ctrl+Shift+Esc)
   - End any EXCEL.EXE processes

2. **Uninstall via Control Panel:**
   - Windows 10/11: Settings → Apps → Apps & features
   - Search for "ExcelPerplexityVSTO"
   - Click Uninstall

   **OR**

   - Control Panel → Programs and Features
   - Find "ExcelPerplexityVSTO"
   - Uninstall

3. **Clean and Rebuild:**
   - Visual Studio → Build → Clean Solution
   - Build → Rebuild Solution
   - Press F5

---

### **Solution 3: Clear ClickOnce Cache Manually**

If the add-in doesn't appear in Control Panel:

1. **Close Excel**

2. **Open Command Prompt as Administrator**

3. **Clear ClickOnce cache:**
   ```cmd
   rundll32 dfshim CleanOnlineAppCache
   ```

4. **Delete cache folders:**
   ```cmd
   rd /s /q "%LocalAppData%\Apps\2.0"
   ```

5. **Clean registry** (optional but recommended):
   - Press Win+R, type `regedit`, press Enter
   - Navigate to: `HKEY_CURRENT_USER\Software\Microsoft\Office\Excel\Addins`
   - Delete `ExcelPerplexityVSTO` key if it exists
   - Navigate to: `HKEY_CURRENT_USER\Software\Microsoft\VSTO\Security\Inclusion`
   - Delete any entries containing "ExcelPerplexityVSTO"

6. **Restart Visual Studio**

7. **Clean and Rebuild**

---

## 🔧 **Advanced Solutions**

### **Solution 4: Change Deployment Location**

This tricks VSTO into thinking it's a new installation:

1. **Open your project in Visual Studio**

2. **Open the .csproj file** (ExcelPerplexityVSTO.csproj)

3. **Add/Change the PublishUrl property:**
   ```xml
   <PropertyGroup>
     <PublishUrl>publish\v2\</PublishUrl>
   </PropertyGroup>
   ```

4. **Save and rebuild**

---

### **Solution 5: Disable ClickOnce Security (Development Only)**

**⚠️ Use only for development/debugging:**

1. **Open Project Properties** (Right-click project → Properties)

2. **Go to "Signing" tab**

3. **Uncheck** "Sign the ClickOnce manifests" (temporarily)

4. **Go to "Security" tab**

5. **Uncheck** "Enable ClickOnce security settings"

6. **Save and rebuild**

**Note:** Re-enable before publishing!

---

### **Solution 6: Use Debug Without ClickOnce**

Bypass ClickOnce deployment entirely during development:

1. **Open Project Properties**

2. **Go to "Debug" tab**

3. **Change "Start Action":**
   - Select "Start external program"
   - Browse to Excel.exe: `C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE`
   (Adjust path for your Office version)

4. **Add Command Line Arguments:**
   ```
   /x "$(TargetPath)"
   ```

5. **Press F5** - Excel will start and load the add-in without installation

---

## 🛠 **Prevention Tips**

### **Best Practices to Avoid This Issue:**

1. **Always use "Build → Clean Solution" before rebuilding**

2. **Close Excel before building in Visual Studio**

3. **During development, consider:**
   - Using "Debug Without ClickOnce" method (Solution 6)
   - Setting up a separate test user account
   - Using a VM for testing

4. **Registry cleanup script:**
   Add this as a Pre-Build event:
   ```cmd
   taskkill /F /IM EXCEL.EXE 2>nul
   reg delete "HKCU\Software\Microsoft\Office\Excel\Addins\ExcelPerplexityVSTO" /f 2>nul
   ```

5. **Increment version number** when testing new builds:
   - Open `Properties\AssemblyInfo.cs`
   - Change `AssemblyVersion` and `AssemblyFileVersion`

---

## 🐛 **Still Not Working?**

### **Nuclear Option - Complete Removal:**

Run this PowerShell script as Administrator:

```powershell
# Kill Excel
Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force

# Clear ClickOnce
rundll32 dfshim CleanOnlineAppCache
Remove-Item "$env:LocalAppData\Apps\2.0" -Recurse -Force -ErrorAction SilentlyContinue

# Clear all VSTO registry entries
Remove-Item "HKCU:\Software\Microsoft\Office\Excel\Addins\ExcelPerplexityVSTO" -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item "HKCU:\Software\Microsoft\VSTO" -Recurse -Force -ErrorAction SilentlyContinue

# Clear project output
$projectPath = "C:\Users\naray\Documents\Projects\ExcelClaude\ExcelPerplexityVSTO\ExcelPerplexityVSTO"
Remove-Item "$projectPath\bin" -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item "$projectPath\obj" -Recurse -Force -ErrorAction SilentlyContinue

# Restart computer
Write-Host "Please restart your computer, then rebuild in Visual Studio"
```

---

## 📋 **Checklist**

Before asking for help, verify:

- [ ] Excel is completely closed (check Task Manager)
- [ ] Visual Studio is running as Administrator
- [ ] .NET Framework 4.8 is installed
- [ ] VSTO Runtime is installed
- [ ] You've run Clean Solution
- [ ] You've cleared ClickOnce cache
- [ ] You've checked registry entries
- [ ] You've restarted Visual Studio
- [ ] You've restarted your computer

---

## 🔍 **Diagnostic Commands**

**Check if add-in is registered:**
```powershell
Get-ItemProperty "HKCU:\Software\Microsoft\Office\Excel\Addins\ExcelPerplexityVSTO"
```

**Check ClickOnce installations:**
```powershell
Get-ChildItem "$env:LocalAppData\Apps\2.0" -Recurse | Where-Object { $_.Name -like "*Perplexity*" }
```

**Check VSTO Runtime:**
```powershell
Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\VSTO Runtime Setup\*"
```

---

## 📞 **Getting More Help**

If none of these solutions work:

1. **Check Windows Event Viewer:**
   - Windows Logs → Application
   - Look for errors from "VSTO 4.0" or "Excel"

2. **Enable VSTO Logging:**
   - Set environment variable: `VSTO_LOGALERTS=1`
   - Restart Visual Studio
   - Reproduce the error
   - Check for log files

3. **Check Excel Trust Center:**
   - Excel → File → Options → Trust Center → Trust Center Settings
   - Macro Settings → Enable all macros (temporarily)
   - Add-in Security → Uncheck "Require application extensions to be signed"

---

## 🎯 **Quick Reference**

| Issue | Solution |
|-------|----------|
| "Another version installed" | Run CleanVSTO.ps1 |
| Can't uninstall from Control Panel | Clear ClickOnce cache manually |
| Registry errors | Delete registry keys manually |
| Certificate errors | Install certificate or lower security |
| LoadBehavior = 2 (not loading) | Set to 3 in registry |
| Manifest errors | Rebuild with Clean Solution |
| Excel crashes on startup | Disable add-in in Safe Mode |

---

**Last Updated:** 2025-01-22
