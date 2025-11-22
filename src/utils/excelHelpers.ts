/**
 * Excel helper utilities for common operations
 */

export interface OHLCData {
  date: Date | string;
  open: number;
  high: number;
  low: number;
  close: number;
  volume?: number;
}

export class ExcelHelpers {
  /**
   * Get the currently selected range address
   */
  static async getSelectedRange(): Promise<string> {
    return Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load('address');
      await context.sync();
      return range.address;
    });
  }

  /**
   * Get data from a specified range
   */
  static async getRangeData(rangeAddress: string): Promise<any[][]> {
    return Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange(rangeAddress);
      range.load('values');
      await context.sync();
      return range.values as any[][];
    });
  }

  /**
   * Create an OHLC (Stock) chart with proper formatting
   * Note: Excel's stockOHLC requires data in columns with specific order
   */
  static async createOHLCChart(rangeAddress: string, chartTitle?: string): Promise<void> {
    return Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();

      // Parse the range to get row information
      const range = sheet.getRange(rangeAddress);
      range.load('address');
      await context.sync();

      const addressParts = range.address.split('!');
      const rangePart = addressParts[1];

      // Parse range like "A1:F9" to get row numbers
      const match = rangePart.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
      if (!match) {
        throw new Error('Invalid range format');
      }

      const startRow = parseInt(match[2]);
      const endRow = parseInt(match[4]);

      // For proper OHLC/Candlestick chart, we need to include Date + OHLC columns
      // Structure should be: Date (A), Open (B), High (C), Low (D), Close (E)
      const chartDataRange = sheet.getRange(`A${startRow}:E${endRow}`);

      // Try stockOHLC first (shows OHLC bars)
      // Note: Excel may render this differently depending on version
      let chart;
      try {
        chart = sheet.charts.add(
          'StockOHLC' as any, // Type assertion for compatibility
          chartDataRange,
          Excel.ChartSeriesBy.columns
        );
      } catch (error) {
        // Fallback to columnClustered if stockOHLC fails
        console.log('StockOHLC not supported, using clustered column chart');
        chart = sheet.charts.add(
          Excel.ChartType.columnClustered,
          chartDataRange,
          Excel.ChartSeriesBy.columns
        );
      }

      // Set chart title
      chart.title.text = chartTitle || 'OHLC Chart';

      // Format the chart
      chart.legend.position = Excel.ChartLegendPosition.bottom;
      chart.legend.visible = true;

      // Position the chart
      chart.top = 20;
      chart.left = 400;
      chart.height = 400;
      chart.width = 600;

      await context.sync();

      // Apply custom formatting
      try {
        chart.series.load('items');
        await context.sync();

        // Color the series: Open=green, High=red, Low=green, Close=red
        const colorMap = ['#00C805', '#EF5350', '#00C805', '#EF5350'];
        for (let i = 0; i < Math.min(chart.series.items.length, 4); i++) {
          const series = chart.series.items[i];
          if (series.format && series.format.fill) {
            series.format.fill.setSolidColor(colorMap[i]);
          }
        }
        await context.sync();
      } catch (error) {
        console.log('Could not apply custom colors:', error);
      }
    });
  }

  /**
   * Create a basic chart of any type
   */
  static async createChart(
    rangeAddress: string,
    chartType: Excel.ChartType,
    chartTitle?: string
  ): Promise<void> {
    return Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange(rangeAddress);

      const chart = sheet.charts.add(chartType, range, Excel.ChartSeriesBy.auto);

      if (chartTitle) {
        chart.title.text = chartTitle;
      }

      chart.legend.position = Excel.ChartLegendPosition.right;
      chart.legend.visible = true;

      // Position the chart
      chart.top = 20;
      chart.left = 400;
      chart.height = 400;
      chart.width = 600;

      await context.sync();
    });
  }

  /**
   * Validate if a range contains OHLC data structure
   */
  static async validateOHLCData(rangeAddress: string): Promise<{ valid: boolean; message: string }> {
    try {
      const data = await this.getRangeData(rangeAddress);

      if (data.length < 2) {
        return { valid: false, message: 'Data must have at least a header row and one data row' };
      }

      const headerRow = data[0];
      if (headerRow.length < 5) {
        return {
          valid: false,
          message: 'OHLC data requires at least 5 columns: Date, Open, High, Low, Close'
        };
      }

      // Check if headers contain expected values (case-insensitive)
      const headerStr = headerRow.map(h => String(h).toLowerCase()).join(',');
      const hasRequiredHeaders =
        headerStr.includes('open') &&
        headerStr.includes('high') &&
        headerStr.includes('low') &&
        headerStr.includes('close');

      if (!hasRequiredHeaders) {
        return {
          valid: false,
          message: 'Headers must include: Open, High, Low, Close (Date column recommended)'
        };
      }

      return { valid: true, message: 'Data structure is valid for OHLC chart' };
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      return { valid: false, message: `Error validating data: ${errorMessage}` };
    }
  }

  /**
   * Format a range as a table
   */
  static async formatAsTable(rangeAddress: string, hasHeaders: boolean = true): Promise<void> {
    return Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange(rangeAddress);

      const table = sheet.tables.add(range, hasHeaders);
      table.style = 'TableStyleMedium2';

      await context.sync();
    });
  }

  /**
   * Apply conditional formatting to highlight positive/negative values
   */
  static async applyConditionalFormatting(rangeAddress: string): Promise<void> {
    return Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange(rangeAddress);

      // Clear existing conditional formatting
      range.conditionalFormats.clearAll();

      // Add formatting for positive values (green)
      const positiveFormat = range.conditionalFormats.add(
        Excel.ConditionalFormatType.cellValue
      );
      positiveFormat.cellValue.rule = { formula1: '0', operator: 'GreaterThan' };
      positiveFormat.cellValue.format.font.color = '#00C805';

      // Add formatting for negative values (red)
      const negativeFormat = range.conditionalFormats.add(
        Excel.ConditionalFormatType.cellValue
      );
      negativeFormat.cellValue.rule = { formula1: '0', operator: 'LessThan' };
      negativeFormat.cellValue.format.font.color = '#EF5350';

      await context.sync();
    });
  }

  /**
   * Insert sample OHLC data for testing
   */
  static async insertSampleOHLCData(): Promise<string> {
    return Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();

      const sampleData = [
        ['Date', 'Open', 'High', 'Low', 'Close', 'Volume'],
        ['2024-01-01', 100, 105, 98, 103, 1000000],
        ['2024-01-02', 103, 108, 102, 106, 1200000],
        ['2024-01-03', 106, 107, 101, 102, 900000],
        ['2024-01-04', 102, 110, 101, 109, 1500000],
        ['2024-01-05', 109, 112, 107, 108, 1100000],
        ['2024-01-06', 108, 115, 108, 114, 1300000],
        ['2024-01-07', 114, 116, 110, 111, 1000000],
        ['2024-01-08', 111, 118, 111, 117, 1400000]
      ];

      const startCell = sheet.getRange('A1');
      const dataRange = startCell.getResizedRange(
        sampleData.length - 1,
        sampleData[0].length - 1
      );

      dataRange.values = sampleData;
      dataRange.format.autofitColumns();

      // Format header row
      const headerRange = sheet.getRange('A1:F1');
      headerRange.format.font.bold = true;
      headerRange.format.fill.color = '#4472C4';
      headerRange.format.font.color = '#FFFFFF';

      // Load the address property before returning
      dataRange.load('address');
      await context.sync();

      return dataRange.address;
    });
  }
}
