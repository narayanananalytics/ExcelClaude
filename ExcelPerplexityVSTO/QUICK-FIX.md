# Quick Fix for Add-in Not Loading

## The Problem

Line 25 in `TaskPane.cs` is causing a **NullReferenceException**:

```csharp
perplexityService = Globals.ThisAddIn.PerplexityService;  // ❌ CRASHES if Globals.ThisAddIn is null
```

This happens in the constructor **before** the VSTO infrastructure is fully initialized.

## The Fix

### Option 1: Use Visual Studio F5 (Recommended - 2 minutes)

**This is THE FASTEST way to get the add-in working:**

1. **Open Visual Studio:**
   ```powershell
   cd C:\Users\naray\Documents\Projects\ExcelClaude\ExcelPerplexityVSTO
   start ExcelPerplexityVSTO.sln
   ```

2. **Press F5**
   - Visual Studio will build and launch Excel with the add-in
   - You'll see the exact error
   - We can fix it from there

**Why this works:**
- Visual Studio has the VSTO Runtime built-in
- The debugger will show the exact line causing the crash
- We can fix code issues immediately

---

### Option 2: Fix the Code Manually (If you want to try)

The issue is in `TaskPane.cs` line 25. We need to delay the service initialization:

**Change this:**
```csharp
public TaskPane()
{
    InitializeComponent();
    perplexityService = Globals.ThisAddIn.PerplexityService;  // ❌ CRASHES
}
```

**To this:**
```csharp
public TaskPane()
{
    InitializeComponent();
    // Don't access Globals.ThisAddIn in constructor
}

protected override void OnLoad(EventArgs e)
{
    base.OnLoad(e);
    // Initialize service after the control is loaded
    if (Globals.ThisAddIn != null)
    {
        perplexityService = Globals.ThisAddIn.PerplexityService;
    }
}
```

Then rebuild and redeploy.

---

### Option 3: Simplify for Testing

Create a minimal version that doesn't use Perplexity Service initially:

**Change TaskPane.cs constructor:**
```csharp
public TaskPane()
{
    InitializeComponent();
    // Initialize service later when actually needed
    perplexityService = null;
}
```

And in `BtnSend_Click`, initialize if null:
```csharp
if (perplexityService == null && Globals.ThisAddIn != null)
{
    perplexityService = Globals.ThisAddIn.PerplexityService;
}
```

---

## Recommended Action

**Use Visual Studio F5** - it's BY FAR the fastest approach:

1. Takes 2 minutes
2. Shows exact errors
3. No VSTO Runtime needed
4. Immediate feedback

**Steps:**
```powershell
# 1. Open solution
cd C:\Users\naray\Documents\Projects\ExcelClaude\ExcelPerplexityVSTO
start ExcelPerplexityVSTO.sln

# 2. In Visual Studio, press F5
# 3. Excel opens with add-in loaded
# 4. Check Output window for any errors
```

Once it works in debug mode, we can address the VSTO Runtime for standalone deployment.

---

## Why LoadBehavior Keeps Changing to 2

**LoadBehavior = 2** means the add-in **crashed during initialization**.

**The crash happens because:**
1. `TaskPane` constructor runs
2. It tries to access `Globals.ThisAddIn.PerplexityService`
3. But `Globals.ThisAddIn` is `null` (not initialized yet)
4. **NullReferenceException** → Excel disables the add-in

**The fix:**
- Don't access `Globals.ThisAddIn` in constructors
- Wait until `OnLoad` or lazily initialize when needed

---

## Next Steps

1. **Open Visual Studio**
2. **Press F5**
3. **Check Output window** for the error message
4. **We'll fix it from there**

This is much faster than trying to fix blindly!
