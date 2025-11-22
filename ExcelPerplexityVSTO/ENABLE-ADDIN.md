# How to Enable the Excel VSTO Add-in

## Problem Identified

The add-in has **LoadBehavior = 2**, which means **Excel disabled it due to an error**.

## Root Causes Found

1. **VSTO Runtime Not Installed** (Primary Issue)
2. Possible code initialization issues

## Solutions

### Option 1: Install VSTO Runtime (For Standalone Deployment)

**Required for the add-in to work outside of Visual Studio.**

1. **Download VSTO Runtime:**
   - Visit: https://www.microsoft.com/en-us/download/details.aspx?id=105522
   - Latest version: Visual Studio 2010 Tools for Office Runtime (v10.0.60917)
   - Note: Despite the "2010" name, this works with Visual Studio 2013-2022

2. **Install:**
   - Run `vstor_redist.exe`
   - Follow the installation wizard
   - Restart your computer

3. **Verify Installation:**
   ```powershell
   Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\VSTO Runtime Setup\v4R"
   ```

4. **Re-enable the add-in:**
   ```powershell
   cd C:\Users\naray\Documents\Projects\ExcelClaude\ExcelPerplexityVSTO
   .\FixAddin.ps1
   ```

5. **Close all Excel instances and reopen Excel**

---

### Option 2: Debug in Visual Studio (For Development)

**Easier option if you have Visual Studio installed.**

Visual Studio includes the VSTO Runtime, so the add-in will work when debugging.

1. **Open the solution in Visual Studio:**
   ```powershell
   cd C:\Users\naray\Documents\Projects\ExcelClaude\ExcelPerplexityVSTO
   start ExcelPerplexityVSTO.sln
   ```

2. **Press F5 to start debugging**
   - Visual Studio will automatically:
     - Build the project
     - Register the add-in
     - Launch Excel with the add-in loaded
     - Attach the debugger

3. **If you see errors:**
   - Check the Output window (View > Output)
   - Look for exceptions in the Debug Console
   - Check if `ThisAddIn_Startup` is throwing errors

4. **Common startup errors to check:**
   - Missing UI files (TaskPane.Designer.cs)
   - Ribbon XML issues
   - API key initialization errors

---

### Option 3: Fix the Code Issues

The custom `ThisAddIn.Designer.cs` we created might not properly initialize the VSTO infrastructure. Let me check if there are missing components:

**Potential Issues:**

1. **Missing TaskPane.Designer.cs:**
   The TaskPane user control needs a designer file.

2. **Missing VSTO initialization:**
   The add-in might not be properly hooking into Excel's COM interface.

3. **Ribbon loading errors:**
   The Ribbon.xml might not be properly embedded.

**To diagnose:**

1. **Enable detailed logging:**
   ```powershell
   # Already done by FixAddin.ps1
   # Logs will be in: C:\Users\naray\AppData\Local\Temp\VSTO\
   ```

2. **Check Event Viewer:**
   ```powershell
   # Open Event Viewer
   eventvwr.msc

   # Navigate to: Windows Logs > Application
   # Look for errors from "VSTO 4.0" or "ExcelPerplexityVSTO"
   ```

3. **Try opening Excel and then check the logs**

---

## Quick Test Steps

### After Installing VSTO Runtime:

1. **Close ALL Excel instances:**
   ```powershell
   Get-Process excel -ErrorAction SilentlyContinue | Stop-Process -Force
   ```

2. **Verify LoadBehavior is set to 3:**
   ```powershell
   Get-ItemProperty "HKCU:\Software\Microsoft\Office\Excel\Addins\ExcelPerplexityVSTO"
   ```
   Should show `LoadBehavior : 3`

3. **Open Excel**

4. **Check if add-in loaded:**
   - Go to: File > Options > Add-ins
   - Manage: COM Add-ins > Go...
   - Look for "Excel Perplexity VSTO Add-in"
   - It should be **checked** and **enabled**

5. **If LoadBehavior changes back to 2:**
   - The add-in crashed during startup
   - Check Event Viewer and VSTO logs
   - Use Visual Studio debugging (Option 2) to see the error

---

## Current Status

| Item | Status |
|------|--------|
| Registry | ✅ Registered |
| LoadBehavior | ✅ Reset to 3 |
| Certificate | ✅ Trusted |
| Dependencies | ✅ All files present |
| VSTO Runtime | ❌ **NOT INSTALLED** |

**Next Action:** Install VSTO Runtime OR use Visual Studio debugging

---

## Alternative: Use Visual Studio Immediately

If you want to test the add-in right now without installing VSTO Runtime:

```powershell
# 1. Open in Visual Studio
cd C:\Users\naray\Documents\Projects\ExcelClaude\ExcelPerplexityVSTO
start ExcelPerplexityVSTO.sln

# 2. In Visual Studio, press F5
# Excel will open with the add-in loaded
```

This is the **fastest way** to test if the add-in works, because Visual Studio handles all the runtime requirements automatically.

---

## Troubleshooting

### If LoadBehavior keeps changing to 2:

This means there's a **code error** in the add-in startup. Common issues:

1. **Missing UI Designer Files:**
   - TaskPane.Designer.cs
   - Settings form designer

2. **Exception in ThisAddIn_Startup:**
   - Check for null references
   - Verify UI controls are properly initialized
   - Make sure Settings.Default is accessible

3. **Ribbon XML not found:**
   - Verify Ribbon.xml is set as Embedded Resource
   - Check the resource name matches the code

### To get detailed error information:

**Method 1: Event Viewer (after enabling logging)**
```powershell
eventvwr.msc
# Navigate to: Windows Logs > Application
# Filter by: Source = "VSTO 4.0"
```

**Method 2: VSTO Log Files**
```powershell
# View latest log
Get-ChildItem "$env:TEMP\VSTO" | Sort-Object LastWriteTime -Descending | Select-Object -First 1 | Get-Content
```

**Method 3: Visual Studio Debugging (Best Option)**
- Press F5 in Visual Studio
- Watch the Output window for errors
- Exceptions will be caught by the debugger

---

## Recommended Approach

**For Immediate Testing:**
1. Open Visual Studio
2. Press F5
3. Test the add-in in the Excel instance that opens

**For Production Deployment:**
1. Install VSTO Runtime
2. Fix any code issues found during Visual Studio testing
3. Rebuild the project
4. Re-deploy using Deploy.ps1

**Current Best Next Step:**
- **Use Visual Studio F5** to see if the add-in has code errors
- This will give you immediate feedback without installing VSTO Runtime
- Once it works in debug mode, then install VSTO Runtime for standalone use
