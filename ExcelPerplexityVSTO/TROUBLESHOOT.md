# Add-in Not Loading - Troubleshooting Guide

## Current Status

**LoadBehavior = 2** - This means Excel is **disabling the add-in** because it's **crashing on startup**.

## Root Cause

The custom `ThisAddIn.Designer.cs` file we created is **not properly initializing** the VSTO infrastructure. VSTO add-ins require special initialization that we didn't fully implement.

## Solutions (In Order of Recommendation)

### ✅ Solution 1: Use Visual Studio to Debug (BEST OPTION)

This is the **fastest way** to see the actual error and fix it.

**Steps:**

1. **Open the solution in Visual Studio:**
   ```powershell
   cd C:\Users\naray\Documents\Projects\ExcelClaude\ExcelPerplexityVSTO
   start ExcelPerplexityVSTO.sln
   ```

2. **Press F5** to start debugging

3. **Visual Studio will:**
   - Build the project
   - Launch Excel with the add-in
   - Show you the **exact error** in the Output window

4. **Check for errors:**
   - Look at the Output window (View > Output)
   - Check the Error List window
   - Any exceptions will be caught by the debugger

5. **Common errors you might see:**
   - Missing `TaskPane.Designer.cs`
   - Null reference in `ThisAddIn_Startup`
   - Missing Settings file
   - Ribbon XML not found

**Advantages:**
- ✅ Shows exact error messages
- ✅ Allows you to set breakpoints
- ✅ No VSTO Runtime installation required
- ✅ Works immediately

---

### Solution 2: Create Missing UI Designer Files

The issue might be missing designer files for UI components.

**Missing Files:**

1. **TaskPane.Designer.cs** - The task pane user control needs a designer file
2. **Settings Designer** - Might be incomplete

Let me create these:

<function_calls>
<invoke name="Read">
<parameter name="file_path">C:\Users\naray\Documents\Projects\ExcelClaude\ExcelPerplexityVSTO\ExcelPerplexityVSTO\UI\TaskPane.cs