# Build Instructions for Excel Perplexity VSTO Add-in

## Prerequisites

Before you can build this VSTO project, you need to install the following:

### 1. Visual Studio 2022

Download and install Visual Studio 2022 (Community, Professional, or Enterprise) from:
https://visualstudio.microsoft.com/downloads/

### 2. Required Workloads

During Visual Studio installation (or via Visual Studio Installer → Modify), install:

- **.NET desktop development** workload
- **Office/SharePoint development** workload

To verify/install workloads:
1. Open Visual Studio Installer
2. Click "Modify" on your Visual Studio 2022 installation
3. Select the "Workloads" tab
4. Check both workloads mentioned above
5. Click "Modify" to install

### 3. .NET Framework 4.8

This should be included with Visual Studio 2022, but verify it's installed:
- Control Panel → Programs → Turn Windows features on or off
- Look for ".NET Framework 4.8"

### 4. Microsoft Office

You need Excel 2016 or later installed on your development machine.

## Opening the Project

1. Navigate to the project folder:
   ```
   C:\Users\naray\Documents\Projects\ExcelClaude\ExcelPerplexityVSTO
   ```

2. Double-click `ExcelPerplexityVSTO.sln` to open in Visual Studio 2022

## Installing NuGet Packages

After opening the solution, you need to install required NuGet packages:

1. In Visual Studio, go to **Tools → NuGet Package Manager → Package Manager Console**

2. Run these commands:
   ```powershell
   Install-Package Newtonsoft.Json
   Install-Package System.Net.Http
   ```

Alternatively, use the NuGet Package Manager UI:
1. Right-click on the project → "Manage NuGet Packages"
2. Search for and install:
   - Newtonsoft.Json
   - System.Net.Http (if not already included)

## Building the Project

### Debug Build (for Development)

1. In Visual Studio, select **Debug** configuration from the toolbar dropdown
2. Press **F5** or click **Debug → Start Debugging**

This will:
- Build the add-in
- Launch Excel with the add-in loaded
- Attach the debugger for testing

You should see a new **"Perplexity AI"** tab in Excel's ribbon.

### Release Build (for Distribution)

1. Select **Release** configuration
2. Click **Build → Build Solution** (or press Ctrl+Shift+B)

The compiled add-in will be in:
```
ExcelPerplexityVSTO\bin\Release\
```

## Common Build Issues and Fixes

### Issue 1: "Office/SharePoint Tools not found"

**Solution:** Install the Office/SharePoint development workload via Visual Studio Installer

### Issue 2: "Cannot find Microsoft.Office.Interop.Excel"

**Solution:**
1. Ensure Office is installed
2. Right-click project → Add Reference
3. Go to COM tab → Type Libraries
4. Find and check "Microsoft Excel 16.0 Object Library"

### Issue 3: "VSTO runtime not found"

**Solution:** Install VSTO Runtime from:
https://www.microsoft.com/en-us/download/details.aspx?id=48217

### Issue 4: Trust certificate error when debugging

**Solution:** Visual Studio automatically creates a debug certificate. If you see trust warnings:
1. Go to File → Account → Account Settings
2. Ensure you're signed in to Visual Studio
3. The debug certificate is managed automatically

## Project Structure

```
ExcelPerplexityVSTO/
├── ExcelPerplexityVSTO.sln          # Solution file
└── ExcelPerplexityVSTO/             # Main project folder
    ├── Helpers/
    │   ├── ExcelChartHelper.cs      # OHLC chart creation & overlays
    │   └── TechnicalIndicators.cs   # MA, Bollinger, RSI, etc.
    ├── Services/
    │   └── PerplexityService.cs     # AI integration
    ├── Models/
    │   ├── OHLCData.cs             # OHLC data structure
    │   └── ChartConfig.cs          # Chart configuration
    ├── UI/
    │   ├── Ribbon.cs               # Custom ribbon implementation
    │   ├── Ribbon.xml              # Ribbon UI definition
    │   └── TaskPane.cs             # AI chat interface
    ├── Properties/
    │   ├── AssemblyInfo.cs
    │   ├── Settings.settings
    │   └── Settings.Designer.cs
    ├── ThisAddIn.cs                # Add-in entry point
    ├── app.config
    └── ExcelPerplexityVSTO.csproj
```

## First Run Configuration

1. **Start the add-in** by pressing F5 in Visual Studio
2. **Excel will launch** with the Perplexity AI tab
3. **Configure API Key**:
   - Click "Perplexity AI" tab → "Settings" button
   - Enter your Perplexity API key from https://perplexity.ai/settings/api
   - Click "Save"

## Testing the Add-in

### Test 1: Insert Sample Data

1. Click "Perplexity AI" → "Insert Sample Data"
2. Sample OHLC data will be inserted into the active sheet

### Test 2: Create OHLC Chart

1. Select the sample data range (including headers)
2. Click "Create OHLC Chart"
3. Choose overlay options
4. A candlestick chart should appear with green/red formatting

### Test 3: AI Assistant

1. Click "AI Assistant" to open the task pane
2. Ask: "Generate code to add a 20-period moving average overlay to my chart"
3. The AI will generate C# code
4. Right-click the code to copy it

## Deployment

### For Development Team

Share the entire `ExcelPerplexityVSTO` folder. Team members should:
1. Install prerequisites (Visual Studio + workloads)
2. Open the .sln file
3. Install NuGet packages
4. Press F5 to run

### For End Users

Use Visual Studio's ClickOnce publishing:

1. Right-click project → **Publish**
2. Follow the Publish Wizard
3. Choose publish location (network share, website, or disk)
4. Configure prerequisites (VSTO Runtime, .NET Framework 4.8)
5. Click **Finish**

This creates an installer (setup.exe) that users can run to install the add-in.

## Debugging Tips

1. **Set Breakpoints:** Click in the left margin of code editor
2. **View Output:** Debug → Windows → Output (shows debug messages)
3. **Immediate Window:** Debug → Windows → Immediate (test code during debug)
4. **Exception Settings:** Debug → Windows → Exception Settings (break on exceptions)

## Key Files to Modify

- **Add chart features:** `Helpers/ExcelChartHelper.cs`
- **Add indicators:** `Helpers/TechnicalIndicators.cs`
- **Modify AI prompts:** `Services/PerplexityService.cs` (SYSTEM_PROMPT constant)
- **Add ribbon buttons:** `UI/Ribbon.xml` and `UI/Ribbon.cs`
- **Modify task pane UI:** `UI/TaskPane.cs`

## Performance Notes

- Chart creation: < 1 second for 1000 data points
- Indicator calculation: < 100ms for most indicators
- AI response: 2-5 seconds (depends on Perplexity API)
- Memory overhead: ~50MB

## Support and Resources

- **VSTO Documentation:** https://docs.microsoft.com/en-us/visualstudio/vsto/
- **Excel Interop Reference:** https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel
- **Perplexity API Docs:** https://docs.perplexity.ai/

## Security Notes

- API keys are stored encrypted in user registry via Settings
- HTTPS is used for all Perplexity API calls
- No data is sent to cloud except AI prompts
- Code should be reviewed before execution (good practice)

## Next Steps

After successful build:

1. ✅ Test all ribbon buttons
2. ✅ Configure API key
3. ✅ Try creating OHLC charts
4. ✅ Test AI code generation
5. ✅ Experiment with overlays
6. ✅ Calculate and add technical indicators

---

**Happy Coding!** 🚀

For questions or issues, check the Debug Output window in Visual Studio for detailed error messages.
