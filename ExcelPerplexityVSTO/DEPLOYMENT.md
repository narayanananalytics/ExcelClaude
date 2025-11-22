# Excel Perplexity VSTO Add-in - Deployment Guide

## Prerequisites

Before deploying the add-in, ensure you have:

1. **Visual Studio Tools for Office Runtime (VSTO Runtime)**
   - Download: https://www.microsoft.com/en-us/download/details.aspx?id=48217
   - Usually pre-installed with Visual Studio or Office

2. **.NET Framework 4.8**
   - Download: https://dotnet.microsoft.com/download/dotnet-framework/net48
   - Usually included with Windows 10/11

3. **Microsoft Excel** (Office 2013 or later)

## Deployment Methods

### Method 1: Quick Deployment (PowerShell - Recommended)

This is the easiest method for local development and testing.

**Steps:**

1. **Close all Excel instances**

2. **Run the deployment script as Administrator:**
   ```powershell
   cd C:\Users\naray\Documents\Projects\ExcelClaude\ExcelPerplexityVSTO
   .\Deploy.ps1
   ```

3. **Open Excel**
   - The add-in should load automatically
   - Look for the custom ribbon tab

4. **Verify installation:**
   - Go to `File > Options > Add-ins`
   - Select `COM Add-ins` from the Manage dropdown
   - Click `Go...`
   - Check that "Excel Perplexity VSTO Add-in" is listed and checked

**To Uninstall:**
```powershell
.\Undeploy.ps1
```

---

### Method 2: Manual Registry Deployment

If you prefer manual control or the PowerShell script doesn't work:

1. **Close all Excel instances**

2. **Double-click `RegisterAddIn.reg`**
   - Click "Yes" to allow registry changes
   - Click "OK" to confirm

3. **Trust the certificate:**
   - Open Certificate Manager: Press `Win+R`, type `certmgr.msc`, press Enter
   - Navigate to `Trusted Publishers > Certificates`
   - Right-click `Certificates` > `All Tasks` > `Import...`
   - Browse to: `ExcelPerplexityVSTO\ExcelPerplexityVSTO_TemporaryKey.pfx`
   - Password: `password`
   - Place in: `Trusted Publishers`

4. **Open Excel** and verify the add-in loads

**To Uninstall:**
- Double-click `UnregisterAddIn.reg`

---

### Method 3: ClickOnce Deployment (For Distribution)

For deploying to multiple users or computers:

**Prerequisites:**
- A web server or network share
- A code signing certificate (for production)

**Steps:**

1. **Configure ClickOnce publishing in Visual Studio:**
   - Right-click project > Properties > Publish
   - Set Publishing Folder (network path or web URL)
   - Set Installation Folder URL
   - Click "Prerequisites" and ensure VSTO Runtime is selected
   - Click "Publish Now"

2. **Share the publish URL with users**
   - Users click the `.vsto` file to install
   - The add-in auto-updates when you republish

**Command Line Publishing:**
```powershell
msbuild ExcelPerplexityVSTO.csproj /t:Publish /p:Configuration=Release /p:PublishUrl="\\server\share\ExcelAddIn\"
```

---

### Method 4: MSI Installer (Enterprise Deployment)

For enterprise environments with Group Policy:

**Tools Required:**
- WiX Toolset (https://wixtoolset.org/)
- Visual Studio with WiX extension

**Steps:**

1. **Create a WiX Setup Project**
2. **Configure to install:**
   - VSTO Runtime (prerequisite)
   - Add-in DLL and manifest
   - Registry entries
   - Certificate to Trusted Publishers
3. **Build MSI**
4. **Deploy via Group Policy or SCCM**

---

## Troubleshooting

### Add-in doesn't appear in Excel

1. **Check if it's registered:**
   ```powershell
   Get-ItemProperty "HKCU:\Software\Microsoft\Office\Excel\Addins\ExcelPerplexityVSTO"
   ```

2. **Check LoadBehavior value:**
   - Should be `3` (load on startup)
   - If it's `2`, Excel disabled it due to an error

3. **Check VSTO Runtime:**
   ```powershell
   Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\VSTO Runtime Setup\v4R"
   ```

4. **Check Excel Trust Center:**
   - File > Options > Trust Center > Trust Center Settings
   - Macro Settings > Enable all macros (for testing)
   - Message Bar > Never show info about blocked content (for testing)

### Certificate Trust Issues

If Excel shows security warnings:

1. **Re-run Deploy.ps1 as Administrator**

2. **Manually trust certificate:**
   - Open `certmgr.msc`
   - Import `ExcelPerplexityVSTO_TemporaryKey.pfx` to Trusted Publishers

3. **For production, use a proper code signing certificate from:**
   - DigiCert
   - Sectigo
   - GlobalSign

### Add-in loads but crashes

1. **Check Event Viewer:**
   - Windows Logs > Application
   - Look for VSTO 4.0 errors

2. **Enable VSTO logging:**
   ```powershell
   New-Item -Path "HKCU:\Software\Microsoft\VSTO\Security" -Force
   Set-ItemProperty -Path "HKCU:\Software\Microsoft\VSTO\Security" -Name "LoadBehavior" -Value 2
   ```
   - Logs will be in `%TEMP%\VSTO\`

3. **Check dependencies:**
   - Ensure `Newtonsoft.Json.dll` is in the output folder

### LoadBehavior keeps changing to 2

This means Excel is disabling the add-in due to errors:

1. **Check the startup code** (`ThisAddIn_Startup`)
2. **Ensure all dependencies are present**
3. **Check for unhandled exceptions**
4. **Reset LoadBehavior:**
   ```powershell
   Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\Excel\Addins\ExcelPerplexityVSTO" -Name "LoadBehavior" -Value 3 -Type DWord
   ```

---

## Security Notes

### Development Certificate

The included certificate (`ExcelPerplexityVSTO_TemporaryKey.pfx`) is for **development only**:
- Password: `password`
- Not suitable for production
- Users will see security warnings

### Production Certificate

For production deployment:

1. **Purchase a code signing certificate**
2. **Re-sign the manifest:**
   ```powershell
   mage -Update ExcelPerplexityVSTO.vsto -CertFile YourCert.pfx -Password YourPassword
   ```

---

## Files Installed

After deployment, the following files are used:

```
ExcelPerplexityVSTO\bin\Release\
├── ExcelPerplexityVSTO.dll          # Main add-in assembly
├── ExcelPerplexityVSTO.dll.manifest # Assembly manifest
├── ExcelPerplexityVSTO.vsto         # VSTO manifest (entry point)
├── ExcelPerplexityVSTO.dll.config   # Configuration file
├── Newtonsoft.Json.dll              # JSON library dependency
└── ExcelPerplexityVSTO.pdb          # Debug symbols (optional)
```

**Registry Location:**
```
HKEY_CURRENT_USER\Software\Microsoft\Office\Excel\Addins\ExcelPerplexityVSTO
```

---

## Next Steps After Deployment

1. **Configure Perplexity API Key:**
   - Open Excel
   - Look for the add-in ribbon tab
   - Click "Settings"
   - Enter your Perplexity API key
   - Get your API key from: https://perplexity.ai/settings/api

2. **Test the add-in:**
   - Click "Insert Sample Data" to test OHLC chart creation
   - Open the AI Assistant task pane
   - Try asking questions about your data

3. **Share with users:**
   - Distribute the installer or publish location
   - Provide the API key setup instructions

---

## Support

For issues or questions:
- Check the Troubleshooting section above
- Review Excel's Add-in Manager (File > Options > Add-ins)
- Check Windows Event Viewer for VSTO errors
- Enable VSTO logging for detailed diagnostics
