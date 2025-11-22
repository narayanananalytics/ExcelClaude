# Code Examples - Excel Perplexity VSTO

Quick reference for common chart and indicator operations.

## Table of Contents

- [Basic Chart Creation](#basic-chart-creation)
- [Chart Overlays](#chart-overlays)
- [Technical Indicators](#technical-indicators)
- [AI Integration](#ai-integration)
- [Data Manipulation](#data-manipulation)
- [Complete Examples](#complete-examples)

---

## Basic Chart Creation

### Create OHLC Chart

```csharp
using ExcelPerplexityVSTO.Helpers;
using Excel = Microsoft.Office.Interop.Excel;

// Get active worksheet
Excel.Worksheet worksheet = Globals.ThisAddIn.Application.ActiveSheet as Excel.Worksheet;

// Define data range (A1:E31 = Date, Open, High, Low, Close)
Excel.Range dataRange = worksheet.Range["A1:E31"];

// Create chart
var chartHelper = new ExcelChartHelper();
Excel.Chart chart = chartHelper.CreateOHLCChart(
    worksheet,
    dataRange,
    "AAPL Stock Price"
);
```

### Create Chart at Specific Position

```csharp
// Create chart objects collection
Excel.ChartObjects chartObjects = (Excel.ChartObjects)worksheet.ChartObjects();

// Add chart at specific position
Excel.ChartObject chartObject = chartObjects.Add(
    Left: 100,      // X position
    Top: 50,        // Y position
    Width: 800,     // Width in points
    Height: 400     // Height in points
);

Excel.Chart chart = chartObject.Chart;
chart.ChartType = Excel.XlChartType.xlStockOHLC;
chart.SetSourceData(dataRange);
```

---

## Chart Overlays

### Add Simple Moving Average

```csharp
using System.Drawing;

// Assume you have calculated MA values in column F
Excel.Range dateRange = worksheet.Range["A2:A31"];
Excel.Range ma20Values = worksheet.Range["F2:F31"];

var chartHelper = new ExcelChartHelper();
Excel.Series maSeries = chartHelper.AddMovingAverageOverlay(
    chart,
    dateRange,
    ma20Values,
    period: 20,
    color: Color.Blue,
    lineWeight: 2f
);
```

### Add Multiple Moving Averages

```csharp
// Add 20-period MA (blue)
chartHelper.AddMovingAverageOverlay(
    chart,
    dateRange,
    worksheet.Range["F2:F31"],
    period: 20,
    color: Color.Blue,
    lineWeight: 2f
);

// Add 50-period MA (orange)
chartHelper.AddMovingAverageOverlay(
    chart,
    dateRange,
    worksheet.Range["G2:G31"],
    period: 50,
    color: Color.Orange,
    lineWeight: 2.5f
);

// Add 200-period MA (red)
chartHelper.AddMovingAverageOverlay(
    chart,
    dateRange,
    worksheet.Range["H2:H31"],
    period: 200,
    color: Color.Red,
    lineWeight: 3f
);
```

### Add Bollinger Bands

```csharp
Excel.Range upperBandRange = worksheet.Range["F2:F31"];
Excel.Range middleBandRange = worksheet.Range["G2:G31"];
Excel.Range lowerBandRange = worksheet.Range["H2:H31"];

chartHelper.AddBollingerBandsOverlay(
    chart,
    dateRange,
    upperBandRange,
    middleBandRange,
    lowerBandRange,
    color: Color.Purple
);
```

### Add Volume Overlay (Secondary Axis)

```csharp
Excel.Range volumeRange = worksheet.Range["F2:F31"];

Excel.Series volumeSeries = chartHelper.AddVolumeOverlay(
    chart,
    dateRange,
    volumeRange,
    upColor: Color.FromArgb(0, 200, 5),      // Green
    downColor: Color.FromArgb(239, 83, 80)   // Red
);

// Volume will be displayed on secondary Y-axis at bottom 25% of chart
```

---

## Technical Indicators

### Calculate Simple Moving Average

```csharp
using ExcelPerplexityVSTO.Helpers;

// Get close prices from Excel
double[] closePrices = new double[30];
for (int i = 0; i < 30; i++)
{
    closePrices[i] = (double)worksheet.Cells[i + 2, 5].Value; // Column E
}

// Calculate 20-period SMA
double?[] sma20 = TechnicalIndicators.CalculateSMA(closePrices, 20);

// Write results to worksheet (column F)
TechnicalIndicators.WriteIndicatorToWorksheet(
    worksheet,
    startRow: 1,
    column: 6,  // Column F
    values: sma20,
    headerName: "MA(20)"
);
```

### Calculate Exponential Moving Average

```csharp
// Calculate 12-period EMA
double?[] ema12 = TechnicalIndicators.CalculateEMA(closePrices, 12);

// Write to worksheet
for (int i = 0; i < ema12.Length; i++)
{
    if (ema12[i].HasValue)
    {
        worksheet.Cells[i + 2, 6] = ema12[i].Value;
    }
}
```

### Calculate Bollinger Bands

```csharp
// Calculate Bollinger Bands (20-period, 2 std dev)
var (upper, middle, lower) = TechnicalIndicators.CalculateBollingerBands(
    closePrices,
    period: 20,
    standardDeviations: 2.0
);

// Write to worksheet
TechnicalIndicators.WriteIndicatorToWorksheet(worksheet, 1, 6, upper, "BB Upper");
TechnicalIndicators.WriteIndicatorToWorksheet(worksheet, 1, 7, middle, "BB Middle");
TechnicalIndicators.WriteIndicatorToWorksheet(worksheet, 1, 8, lower, "BB Lower");
```

### Calculate RSI

```csharp
// Calculate 14-period RSI
double?[] rsi = TechnicalIndicators.CalculateRSI(closePrices, 14);

// Write to worksheet
TechnicalIndicators.WriteIndicatorToWorksheet(worksheet, 1, 6, rsi, "RSI(14)");

// Create separate RSI chart
var rsiChart = chartHelper.CreateRSIChart(
    worksheet,
    dateRange,
    worksheet.Range["F2:F31"],
    leftPosition: 50,
    topPosition: 500  // Below main chart
);
```

### Calculate MACD

```csharp
// Calculate MACD (12, 26, 9)
var (macdLine, signalLine, histogram) = TechnicalIndicators.CalculateMACD(
    closePrices,
    fastPeriod: 12,
    slowPeriod: 26,
    signalPeriod: 9
);

// Write to worksheet
TechnicalIndicators.WriteIndicatorToWorksheet(worksheet, 1, 6, macdLine, "MACD");
TechnicalIndicators.WriteIndicatorToWorksheet(worksheet, 1, 7, signalLine, "Signal");
TechnicalIndicators.WriteIndicatorToWorksheet(worksheet, 1, 8, histogram, "Histogram");
```

### Calculate ATR (Average True Range)

```csharp
// Get high, low, close prices
double[] highs = GetColumnValues(worksheet, 3, 30);   // Column C
double[] lows = GetColumnValues(worksheet, 4, 30);    // Column D
double[] closes = GetColumnValues(worksheet, 5, 30);  // Column E

// Calculate 14-period ATR
double?[] atr = TechnicalIndicators.CalculateATR(highs, lows, closes, 14);

// Write to worksheet
TechnicalIndicators.WriteIndicatorToWorksheet(worksheet, 1, 6, atr, "ATR(14)");
```

### Calculate Stochastic Oscillator

```csharp
// Calculate Stochastic (14, 3)
var (percentK, percentD) = TechnicalIndicators.CalculateStochastic(
    highs,
    lows,
    closes,
    kPeriod: 14,
    dPeriod: 3
);

// Write to worksheet
TechnicalIndicators.WriteIndicatorToWorksheet(worksheet, 1, 6, percentK, "%K");
TechnicalIndicators.WriteIndicatorToWorksheet(worksheet, 1, 7, percentD, "%D");
```

---

## AI Integration

### Basic AI Chat

```csharp
using ExcelPerplexityVSTO.Services;

var perplexity = Globals.ThisAddIn.PerplexityService;

// Send message and get response
string response = await perplexity.SendMessageAsync(
    "How do I create a candlestick chart in Excel using C#?"
);

// Display response
MessageBox.Show(response);
```

### Generate Chart Code

```csharp
// Generate code for OHLC chart
string code = await perplexity.GenerateChartCodeAsync(
    chartType: "OHLC",
    dataRange: "A1:E31",
    customization: "Add 20-period moving average overlay"
);

// Extract C# code from response
string csharpCode = PerplexityService.ExtractCSharpCode(code);

// Display in task pane or message box
MessageBox.Show(csharpCode, "Generated Code");
```

### Generate Overlay Code

```csharp
// Generate code to add Bollinger Bands
var parameters = new Dictionary<string, string>
{
    { "period", "20" },
    { "stdDev", "2" },
    { "color", "purple" }
};

string overlayCode = await perplexity.GenerateOverlayCodeAsync(
    overlayType: "Bollinger Bands",
    dataRange: "F1:H31",
    parameters: parameters
);
```

### Generate Indicator Calculation Code

```csharp
// Generate RSI calculation code
string indicatorCode = await perplexity.GenerateIndicatorCodeAsync(
    indicatorType: "RSI",
    period: 14,
    sourceRange: "E2:E31"
);
```

---

## Data Manipulation

### Read Data from Excel

```csharp
// Read range into array
Excel.Range dataRange = worksheet.Range["E2:E31"];
object[,] values = dataRange.Value2 as object[,];

double[] closePrices = new double[values.GetLength(0)];
for (int i = 0; i < values.GetLength(0); i++)
{
    closePrices[i] = Convert.ToDouble(values[i + 1, 1]);
}
```

### Helper Method: Get Column Values

```csharp
private double[] GetColumnValues(Excel.Worksheet worksheet, int column, int count)
{
    double[] values = new double[count];
    for (int i = 0; i < count; i++)
    {
        values[i] = (double)worksheet.Cells[i + 2, column].Value;
    }
    return values;
}
```

### Write Data to Excel

```csharp
// Write array to column
for (int i = 0; i < sma20.Length; i++)
{
    if (sma20[i].HasValue)
    {
        worksheet.Cells[i + 2, 6] = sma20[i].Value;
    }
}

// Or use Range for bulk write
Excel.Range outputRange = worksheet.Range[
    worksheet.Cells[2, 6],
    worksheet.Cells[31, 6]
];

object[,] outputValues = new object[30, 1];
for (int i = 0; i < 30; i++)
{
    outputValues[i, 0] = sma20[i].HasValue ? sma20[i].Value : (object)"";
}
outputRange.Value2 = outputValues;
```

---

## Complete Examples

### Example 1: Full Trading Chart

```csharp
using System.Drawing;
using ExcelPerplexityVSTO.Helpers;
using Excel = Microsoft.Office.Interop.Excel;

public void CreateFullTradingChart()
{
    Excel.Worksheet worksheet = Globals.ThisAddIn.Application.ActiveSheet as Excel.Worksheet;

    // 1. Get close prices
    double[] closes = new double[30];
    for (int i = 0; i < 30; i++)
    {
        closes[i] = (double)worksheet.Cells[i + 2, 5].Value;
    }

    // 2. Calculate indicators
    double?[] ma20 = TechnicalIndicators.CalculateSMA(closes, 20);
    double?[] ma50 = TechnicalIndicators.CalculateSMA(closes, 50);
    var (bbUpper, bbMiddle, bbLower) = TechnicalIndicators.CalculateBollingerBands(closes, 20, 2.0);

    // 3. Write indicators to worksheet
    TechnicalIndicators.WriteIndicatorToWorksheet(worksheet, 1, 7, ma20, "MA(20)");
    TechnicalIndicators.WriteIndicatorToWorksheet(worksheet, 1, 8, ma50, "MA(50)");
    TechnicalIndicators.WriteIndicatorToWorksheet(worksheet, 1, 9, bbUpper, "BB Upper");
    TechnicalIndicators.WriteIndicatorToWorksheet(worksheet, 1, 10, bbMiddle, "BB Mid");
    TechnicalIndicators.WriteIndicatorToWorksheet(worksheet, 1, 11, bbLower, "BB Lower");

    // 4. Create OHLC chart
    var chartHelper = new ExcelChartHelper();
    Excel.Range ohlcRange = worksheet.Range["A1:E31"];
    Excel.Chart chart = chartHelper.CreateOHLCChart(worksheet, ohlcRange, "Trading Chart");

    // 5. Add overlays
    Excel.Range dates = worksheet.Range["A2:A31"];

    chartHelper.AddMovingAverageOverlay(chart, dates, worksheet.Range["G2:G31"], 20, Color.Blue, 2f);
    chartHelper.AddMovingAverageOverlay(chart, dates, worksheet.Range["H2:H31"], 50, Color.Orange, 2.5f);

    chartHelper.AddBollingerBandsOverlay(
        chart, dates,
        worksheet.Range["I2:I31"],
        worksheet.Range["J2:J31"],
        worksheet.Range["K2:K31"],
        Color.Purple
    );

    // 6. Add volume
    Excel.Range volumeRange = worksheet.Range["F2:F31"];
    chartHelper.AddVolumeOverlay(chart, dates, volumeRange);

    MessageBox.Show("Trading chart created successfully!");
}
```

### Example 2: Batch Indicator Calculation

```csharp
public void CalculateAllIndicators()
{
    Excel.Worksheet worksheet = Globals.ThisAddIn.Application.ActiveSheet as Excel.Worksheet;

    // Read price data
    int rowCount = 30;
    double[] highs = new double[rowCount];
    double[] lows = new double[rowCount];
    double[] closes = new double[rowCount];

    for (int i = 0; i < rowCount; i++)
    {
        highs[i] = (double)worksheet.Cells[i + 2, 3].Value;
        lows[i] = (double)worksheet.Cells[i + 2, 4].Value;
        closes[i] = (double)worksheet.Cells[i + 2, 5].Value;
    }

    // Calculate all indicators
    double?[] sma20 = TechnicalIndicators.CalculateSMA(closes, 20);
    double?[] ema12 = TechnicalIndicators.CalculateEMA(closes, 12);
    var (bbU, bbM, bbL) = TechnicalIndicators.CalculateBollingerBands(closes, 20, 2.0);
    double?[] rsi = TechnicalIndicators.CalculateRSI(closes, 14);
    var (macd, signal, hist) = TechnicalIndicators.CalculateMACD(closes, 12, 26, 9);
    double?[] atr = TechnicalIndicators.CalculateATR(highs, lows, closes, 14);

    // Write all to worksheet
    int col = 7; // Start at column G
    TechnicalIndicators.WriteIndicatorToWorksheet(worksheet, 1, col++, sma20, "SMA(20)");
    TechnicalIndicators.WriteIndicatorToWorksheet(worksheet, 1, col++, ema12, "EMA(12)");
    TechnicalIndicators.WriteIndicatorToWorksheet(worksheet, 1, col++, bbU, "BB Upper");
    TechnicalIndicators.WriteIndicatorToWorksheet(worksheet, 1, col++, bbM, "BB Mid");
    TechnicalIndicators.WriteIndicatorToWorksheet(worksheet, 1, col++, bbL, "BB Lower");
    TechnicalIndicators.WriteIndicatorToWorksheet(worksheet, 1, col++, rsi, "RSI(14)");
    TechnicalIndicators.WriteIndicatorToWorksheet(worksheet, 1, col++, macd, "MACD");
    TechnicalIndicators.WriteIndicatorToWorksheet(worksheet, 1, col++, signal, "Signal");
    TechnicalIndicators.WriteIndicatorToWorksheet(worksheet, 1, col++, hist, "Histogram");
    TechnicalIndicators.WriteIndicatorToWorksheet(worksheet, 1, col++, atr, "ATR(14)");

    MessageBox.Show($"Calculated {col - 7} indicators!");
}
```

### Example 3: AI-Powered Chart Generation

```csharp
public async void GenerateChartFromPrompt(string userPrompt)
{
    var perplexity = Globals.ThisAddIn.PerplexityService;

    try
    {
        // Get selected range
        Excel.Range selection = Globals.ThisAddIn.Application.Selection as Excel.Range;
        string rangeAddress = selection.Address;

        // Build prompt
        string fullPrompt = $"{userPrompt}\n\nData is in range: {rangeAddress}";

        // Get AI response
        string response = await perplexity.SendMessageAsync(fullPrompt);

        // Extract code
        string code = PerplexityService.ExtractCSharpCode(response);

        // Display code for user review
        using (var codeViewer = new Form())
        {
            codeViewer.Text = "Generated Code";
            codeViewer.Size = new Size(800, 600);

            var textBox = new TextBox
            {
                Multiline = true,
                Dock = DockStyle.Fill,
                Font = new Font("Consolas", 10),
                Text = code,
                ScrollBars = ScrollBars.Both
            };

            var btnCopy = new Button
            {
                Text = "Copy to Clipboard",
                Dock = DockStyle.Bottom,
                Height = 40
            };
            btnCopy.Click += (s, e) =>
            {
                Clipboard.SetText(code);
                MessageBox.Show("Code copied!");
            };

            codeViewer.Controls.Add(textBox);
            codeViewer.Controls.Add(btnCopy);
            codeViewer.ShowDialog();
        }
    }
    catch (Exception ex)
    {
        MessageBox.Show($"Error: {ex.Message}", "AI Generation Failed");
    }
}
```

### Example 4: Custom Chart with Multiple Panes

```csharp
public void CreateMultiPaneChart()
{
    Excel.Worksheet worksheet = Globals.ThisAddIn.Application.ActiveSheet as Excel.Worksheet;
    var chartHelper = new ExcelChartHelper();

    // Read data
    double[] closes = GetColumnValues(worksheet, 5, 30);
    double[] highs = GetColumnValues(worksheet, 3, 30);
    double[] lows = GetColumnValues(worksheet, 4, 30);

    // Calculate indicators
    double?[] rsi = TechnicalIndicators.CalculateRSI(closes, 14);
    var (macd, signal, hist) = TechnicalIndicators.CalculateMACD(closes, 12, 26, 9);

    // Write to worksheet
    TechnicalIndicators.WriteIndicatorToWorksheet(worksheet, 1, 7, rsi, "RSI");
    TechnicalIndicators.WriteIndicatorToWorksheet(worksheet, 1, 8, macd, "MACD");

    // 1. Main OHLC chart (top)
    Excel.Range ohlcRange = worksheet.Range["A1:E31"];
    Excel.Chart mainChart = chartHelper.CreateOHLCChart(
        worksheet, ohlcRange, "Price Chart"
    );

    // Position main chart
    ((Excel.ChartObject)mainChart.Parent).Top = 50;
    ((Excel.ChartObject)mainChart.Parent).Left = 50;
    ((Excel.ChartObject)mainChart.Parent).Height = 300;

    // 2. RSI chart (middle)
    Excel.Range dates = worksheet.Range["A2:A31"];
    Excel.Chart rsiChart = chartHelper.CreateRSIChart(
        worksheet,
        dates,
        worksheet.Range["G2:G31"],
        leftPosition: 50,
        topPosition: 375
    );

    // 3. MACD chart (bottom) - Create manually
    Excel.ChartObjects chartObjects = (Excel.ChartObjects)worksheet.ChartObjects();
    Excel.ChartObject macdChartObj = chartObjects.Add(50, 550, 800, 150);
    Excel.Chart macdChart = macdChartObj.Chart;

    macdChart.ChartType = Excel.XlChartType.xlLine;
    macdChart.HasTitle = true;
    macdChart.ChartTitle.Text = "MACD";

    // Add MACD line
    Excel.SeriesCollection series = macdChart.SeriesCollection();
    Excel.Series macdSeries = series.NewSeries();
    macdSeries.XValues = dates;
    macdSeries.Values = worksheet.Range["H2:H31"];
    macdSeries.Name = "MACD";
    macdSeries.Format.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.Blue);

    MessageBox.Show("Multi-pane chart created!");
}

private double[] GetColumnValues(Excel.Worksheet ws, int col, int count)
{
    double[] values = new double[count];
    for (int i = 0; i < count; i++)
    {
        values[i] = (double)ws.Cells[i + 2, col].Value;
    }
    return values;
}
```

---

## Tips and Best Practices

### 1. Always Use Using Statements (When Needed)

```csharp
// For COM objects that need cleanup
var worksheet = Globals.ThisAddIn.Application.ActiveSheet as Excel.Worksheet;
// Work with worksheet
// VSTO handles cleanup automatically for most objects
```

### 2. Error Handling

```csharp
try
{
    var chart = chartHelper.CreateOHLCChart(worksheet, dataRange, "Chart");
}
catch (Exception ex)
{
    MessageBox.Show($"Error creating chart: {ex.Message}", "Error");
    System.Diagnostics.Debug.WriteLine($"Chart error: {ex.StackTrace}");
}
```

### 3. Check for Null Values

```csharp
if (sma20[i].HasValue)
{
    worksheet.Cells[i + 2, 6] = sma20[i].Value;
}
else
{
    worksheet.Cells[i + 2, 6] = "";
}
```

### 4. Performance Optimization

```csharp
// Turn off screen updating for batch operations
Excel.Application app = Globals.ThisAddIn.Application;
app.ScreenUpdating = false;

try
{
    // Perform bulk operations
    for (int i = 0; i < 1000; i++)
    {
        // Write data
    }
}
finally
{
    app.ScreenUpdating = true;
}
```

### 5. Use ColorTranslator for Colors

```csharp
using System.Drawing;

// Convert System.Drawing.Color to Excel OLE color
series.Format.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.Blue);
```

---

Happy coding! For more examples, check the `Helpers/` and `Services/` folders in the project.
