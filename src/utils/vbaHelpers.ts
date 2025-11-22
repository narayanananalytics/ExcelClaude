/**
 * VBA execution helpers for Excel
 * Note: Office.js doesn't support direct VBA execution
 * These helpers prepare VBA code for manual execution or clipboard copy
 */

export interface VBACode {
  code: string;
  description: string;
}

export class VBAHelpers {
  /**
   * Extract VBA code from Perplexity response
   */
  static extractVBACode(response: string): VBACode | null {
    // Look for VBA code blocks in markdown format
    const vbaMatch = response.match(/```vba\n([\s\S]*?)```/i) ||
                     response.match(/```VBA\n([\s\S]*?)```/i) ||
                     response.match(/```\n(Sub [\s\S]*?End Sub)```/i);

    if (vbaMatch && vbaMatch[1]) {
      const code = vbaMatch[1].trim();

      // Extract description from comments at the top
      const descMatch = code.match(/^\s*'\s*(.+)/);
      const description = descMatch ? descMatch[1] : 'Generated VBA macro';

      return {
        code,
        description
      };
    }

    // Fallback: try to find Sub...End Sub directly
    const directMatch = response.match(/(Sub [\s\S]*?End Sub)/i);
    if (directMatch) {
      return {
        code: directMatch[1].trim(),
        description: 'Generated VBA macro'
      };
    }

    return null;
  }

  /**
   * Copy VBA code to clipboard
   * Note: Clipboard API requires user interaction and HTTPS
   */
  static async copyToClipboard(code: string): Promise<boolean> {
    try {
      if (navigator.clipboard && navigator.clipboard.writeText) {
        await navigator.clipboard.writeText(code);
        return true;
      }

      // Fallback for older browsers
      const textArea = document.createElement('textarea');
      textArea.value = code;
      textArea.style.position = 'fixed';
      textArea.style.left = '-999999px';
      document.body.appendChild(textArea);
      textArea.focus();
      textArea.select();

      const successful = document.execCommand('copy');
      document.body.removeChild(textArea);
      return successful;
    } catch (error) {
      console.error('Failed to copy to clipboard:', error);
      return false;
    }
  }

  /**
   * Generate OHLC chart VBA code
   */
  static generateOHLCChartVBA(rangeAddress: string): string {
    return `Sub CreateOHLCChart()
    ' Creates a properly formatted OHLC candlestick chart
    ' Data should be in format: Date, Open, High, Low, Close

    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim dataRange As Range
    Dim cht As Chart

    Set ws = ActiveSheet
    Set dataRange = ws.Range("${rangeAddress}")

    ' Create new chart
    Set chartObj = ws.ChartObjects.Add( _
        Left:=400, _
        Width:=600, _
        Top:=20, _
        Height:=400)

    Set cht = chartObj.Chart

    ' Set chart type to Stock OHLC
    cht.SetSourceData Source:=dataRange
    cht.ChartType = xlStockOHLC

    ' Format the chart
    cht.HasTitle = True
    cht.ChartTitle.Text = "OHLC Chart"

    ' Format Up Bars (bullish - green)
    With cht.ChartGroups(1)
        .UpBars.Interior.Color = RGB(0, 200, 5)
        .UpBars.Border.Color = RGB(0, 200, 5)

        ' Format Down Bars (bearish - red)
        .DownBars.Interior.Color = RGB(239, 83, 80)
        .DownBars.Border.Color = RGB(239, 83, 80)

        ' Gap width between bars
        .GapWidth = 50
    End With

    ' Format axes
    With cht.Axes(xlValue)
        .HasTitle = True
        .AxisTitle.Text = "Price"
    End With

    With cht.Axes(xlCategory)
        .HasTitle = True
        .AxisTitle.Text = "Date"
    End With

    ' Add legend
    cht.HasLegend = True
    cht.Legend.Position = xlLegendPositionBottom

    MsgBox "OHLC chart created successfully!", vbInformation

    ' Cleanup
    Set cht = Nothing
    Set chartObj = Nothing
    Set dataRange = Nothing
    Set ws = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Error creating chart: " & Err.Description, vbCritical
    If Not cht Is Nothing Then Set cht = Nothing
    If Not chartObj Is Nothing Then Set chartObj = Nothing
    If Not dataRange Is Nothing Then Set dataRange = Nothing
    If Not ws Is Nothing Then Set ws = Nothing
End Sub`;
  }

  /**
   * Format VBA code for display
   */
  static formatForDisplay(code: string): string {
    return code
      .split('\n')
      .map(line => line)
      .join('\n');
  }

  /**
   * Get instructions for executing VBA in Excel
   */
  static getExecutionInstructions(): string {
    return `To execute this VBA macro:

1. Copy the VBA code (click "Copy Code" button)
2. In Excel, press Alt+F11 to open the VBA Editor
3. In the VBA Editor, go to Insert → Module
4. Paste the code into the module window
5. Close the VBA Editor (Alt+Q)
6. Press Alt+F8 to open the Macro dialog
7. Select "GeneratedMacro" (or the macro name) and click "Run"

Note: If macros are disabled, go to:
File → Options → Trust Center → Trust Center Settings → Macro Settings
→ Enable "Enable all macros" (or "Disable all macros with notification")`;
  }
}
