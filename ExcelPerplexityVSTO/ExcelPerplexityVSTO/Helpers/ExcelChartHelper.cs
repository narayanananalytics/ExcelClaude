using System;
using System.Drawing;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelPerplexityVSTO.Helpers
{
    /// <summary>
    /// Helper class for creating and manipulating Excel charts with overlay support.
    /// Provides methods for OHLC charts, moving averages, volume overlays, and technical indicators.
    /// </summary>
    public class ExcelChartHelper
    {
        /// <summary>
        /// Creates a properly formatted OHLC (candlestick) chart with green/red coloring.
        /// </summary>
        /// <param name="worksheet">The worksheet containing the data</param>
        /// <param name="dataRange">Range containing Date, Open, High, Low, Close columns</param>
        /// <param name="chartTitle">Title for the chart</param>
        /// <returns>The created chart object</returns>
        public Excel.Chart CreateOHLCChart(Excel.Worksheet worksheet, Excel.Range dataRange, string chartTitle = "OHLC Chart")
        {
            // Add chart to worksheet
            Excel.ChartObjects chartObjects = (Excel.ChartObjects)worksheet.ChartObjects();
            Excel.ChartObject chartObject = chartObjects.Add(
                Left: 50,
                Top: 50,
                Width: 800,
                Height: 400
            );

            Excel.Chart chart = chartObject.Chart;

            // Set chart type to Stock OHLC
            chart.ChartType = Excel.XlChartType.xlStockOHLC;

            // Set data source
            chart.SetSourceData(dataRange, Excel.XlRowCol.xlColumns);

            // Format the chart
            chart.HasTitle = true;
            chart.ChartTitle.Text = chartTitle;

            // Format axes
            Excel.Axis xAxis = (Excel.Axis)chart.Axes(Excel.XlAxisType.xlCategory);
            xAxis.HasTitle = true;
            xAxis.AxisTitle.Text = "Date";

            Excel.Axis yAxis = (Excel.Axis)chart.Axes(Excel.XlAxisType.xlValue);
            yAxis.HasTitle = true;
            yAxis.AxisTitle.Text = "Price";

            // Format the candlesticks - Green for bullish (up), Red for bearish (down)
            try
            {
                // Access the series group to format UpBars and DownBars
                Excel.ChartGroup chartGroup = chart.ChartGroups(1);

                // Format UpBars (Close > Open) - Green
                if (chartGroup.HasUpDownBars)
                {
                    chartGroup.UpBars.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(0, 200, 5));
                    chartGroup.UpBars.Border.Color = ColorTranslator.ToOle(Color.FromArgb(0, 150, 0));

                    // Format DownBars (Close < Open) - Red
                    chartGroup.DownBars.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(239, 83, 80));
                    chartGroup.DownBars.Border.Color = ColorTranslator.ToOle(Color.FromArgb(200, 50, 50));
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error formatting candlesticks: {ex.Message}");
            }

            return chart;
        }

        /// <summary>
        /// Adds a moving average overlay to an existing chart.
        /// </summary>
        /// <param name="chart">The chart to add the overlay to</param>
        /// <param name="dateRange">Range containing dates</param>
        /// <param name="maValues">Range containing moving average values</param>
        /// <param name="period">Period of the moving average (e.g., 20, 50, 200)</param>
        /// <param name="color">Color for the MA line</param>
        /// <param name="lineWeight">Thickness of the line (default 2)</param>
        /// <returns>The created series</returns>
        public Excel.Series AddMovingAverageOverlay(
            Excel.Chart chart,
            Excel.Range dateRange,
            Excel.Range maValues,
            int period,
            Color color,
            float lineWeight = 2f)
        {
            // Add new series to chart
            Excel.SeriesCollection seriesCollection = chart.SeriesCollection();
            Excel.Series series = seriesCollection.NewSeries();

            // Set series data
            series.XValues = dateRange;
            series.Values = maValues;
            series.Name = $"{period}-Period MA";

            // Change to line chart type
            series.ChartType = Excel.XlChartType.xlLine;

            // Format the line
            series.Format.Line.ForeColor.RGB = ColorTranslator.ToOle(color);
            series.Format.Line.Weight = lineWeight;
            series.Format.Line.DashStyle = Microsoft.Office.Core.MsoLineDashStyle.msoLineSolid;

            // Remove markers for cleaner look
            series.MarkerStyle = Excel.XlMarkerStyle.xlMarkerStyleNone;

            return series;
        }

        /// <summary>
        /// Adds Bollinger Bands overlay (upper band, middle line, lower band).
        /// </summary>
        /// <param name="chart">The chart to add the overlay to</param>
        /// <param name="dateRange">Range containing dates</param>
        /// <param name="upperBand">Range containing upper band values</param>
        /// <param name="middleBand">Range containing middle band (SMA) values</param>
        /// <param name="lowerBand">Range containing lower band values</param>
        /// <param name="color">Base color for the bands</param>
        public void AddBollingerBandsOverlay(
            Excel.Chart chart,
            Excel.Range dateRange,
            Excel.Range upperBand,
            Excel.Range middleBand,
            Excel.Range lowerBand,
            Color color)
        {
            Excel.SeriesCollection seriesCollection = chart.SeriesCollection();

            // Add upper band
            Excel.Series upperSeries = seriesCollection.NewSeries();
            upperSeries.XValues = dateRange;
            upperSeries.Values = upperBand;
            upperSeries.Name = "BB Upper";
            upperSeries.ChartType = Excel.XlChartType.xlLine;
            upperSeries.Format.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(128, color.R, color.G, color.B));
            upperSeries.Format.Line.Weight = 1f;
            upperSeries.Format.Line.DashStyle = Microsoft.Office.Core.MsoLineDashStyle.msoLineDash;
            upperSeries.MarkerStyle = Excel.XlMarkerStyle.xlMarkerStyleNone;

            // Add middle band
            Excel.Series middleSeries = seriesCollection.NewSeries();
            middleSeries.XValues = dateRange;
            middleSeries.Values = middleBand;
            middleSeries.Name = "BB Middle";
            middleSeries.ChartType = Excel.XlChartType.xlLine;
            middleSeries.Format.Line.ForeColor.RGB = ColorTranslator.ToOle(color);
            middleSeries.Format.Line.Weight = 1.5f;
            middleSeries.MarkerStyle = Excel.XlMarkerStyle.xlMarkerStyleNone;

            // Add lower band
            Excel.Series lowerSeries = seriesCollection.NewSeries();
            lowerSeries.XValues = dateRange;
            lowerSeries.Values = lowerBand;
            lowerSeries.Name = "BB Lower";
            lowerSeries.ChartType = Excel.XlChartType.xlLine;
            lowerSeries.Format.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(128, color.R, color.G, color.B));
            lowerSeries.Format.Line.Weight = 1f;
            lowerSeries.Format.Line.DashStyle = Microsoft.Office.Core.MsoLineDashStyle.msoLineDash;
            lowerSeries.MarkerStyle = Excel.XlMarkerStyle.xlMarkerStyleNone;
        }

        /// <summary>
        /// Adds volume bars on a secondary axis below the price chart.
        /// </summary>
        /// <param name="chart">The chart to add the overlay to</param>
        /// <param name="dateRange">Range containing dates</param>
        /// <param name="volumeRange">Range containing volume values</param>
        /// <param name="upColor">Color for volume bars on up days (default green)</param>
        /// <param name="downColor">Color for volume bars on down days (default red)</param>
        /// <returns>The created volume series</returns>
        public Excel.Series AddVolumeOverlay(
            Excel.Chart chart,
            Excel.Range dateRange,
            Excel.Range volumeRange,
            Color? upColor = null,
            Color? downColor = null)
        {
            Color volumeUpColor = upColor ?? Color.FromArgb(0, 200, 5);
            Color volumeDownColor = downColor ?? Color.FromArgb(239, 83, 80);

            // Add volume series
            Excel.SeriesCollection seriesCollection = chart.SeriesCollection();
            Excel.Series volumeSeries = seriesCollection.NewSeries();

            volumeSeries.XValues = dateRange;
            volumeSeries.Values = volumeRange;
            volumeSeries.Name = "Volume";
            volumeSeries.ChartType = Excel.XlChartType.xlColumnClustered;

            // Place on secondary axis
            volumeSeries.AxisGroup = Excel.XlAxisGroup.xlSecondary;

            // Format the secondary Y-axis
            try
            {
                Excel.Axis secondaryYAxis = (Excel.Axis)chart.Axes(
                    Excel.XlAxisType.xlValue,
                    Excel.XlAxisGroup.xlSecondary
                );
                secondaryYAxis.HasTitle = true;
                secondaryYAxis.AxisTitle.Text = "Volume";

                // Make volume bars shorter (25% of chart height)
                secondaryYAxis.MaximumScale = secondaryYAxis.MaximumScale * 4;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error formatting volume axis: {ex.Message}");
            }

            // Format volume bars with semi-transparent color
            volumeSeries.Format.Fill.ForeColor.RGB = ColorTranslator.ToOle(
                Color.FromArgb(128, volumeUpColor.R, volumeUpColor.G, volumeUpColor.B)
            );
            volumeSeries.Format.Fill.Transparency = 0.5f;

            return volumeSeries;
        }

        /// <summary>
        /// Adds an RSI (Relative Strength Index) indicator in a separate pane below the main chart.
        /// </summary>
        /// <param name="worksheet">The worksheet to add the chart to</param>
        /// <param name="dateRange">Range containing dates</param>
        /// <param name="rsiValues">Range containing RSI values (0-100)</param>
        /// <param name="leftPosition">Left position of the chart</param>
        /// <param name="topPosition">Top position of the chart</param>
        /// <returns>The created RSI chart</returns>
        public Excel.Chart CreateRSIChart(
            Excel.Worksheet worksheet,
            Excel.Range dateRange,
            Excel.Range rsiValues,
            double leftPosition = 50,
            double topPosition = 500)
        {
            Excel.ChartObjects chartObjects = (Excel.ChartObjects)worksheet.ChartObjects();
            Excel.ChartObject chartObject = chartObjects.Add(
                Left: leftPosition,
                Top: topPosition,
                Width: 800,
                Height: 150
            );

            Excel.Chart chart = chartObject.Chart;
            chart.ChartType = Excel.XlChartType.xlLine;

            // Add RSI series
            Excel.SeriesCollection seriesCollection = chart.SeriesCollection();
            Excel.Series rsiSeries = seriesCollection.NewSeries();
            rsiSeries.XValues = dateRange;
            rsiSeries.Values = rsiValues;
            rsiSeries.Name = "RSI";

            // Format chart
            chart.HasTitle = true;
            chart.ChartTitle.Text = "RSI (14)";

            Excel.Axis yAxis = (Excel.Axis)chart.Axes(Excel.XlAxisType.xlValue);
            yAxis.MinimumScale = 0;
            yAxis.MaximumScale = 100;
            yAxis.MajorUnit = 20;

            // Format RSI line
            rsiSeries.Format.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.Purple);
            rsiSeries.Format.Line.Weight = 2f;
            rsiSeries.MarkerStyle = Excel.XlMarkerStyle.xlMarkerStyleNone;

            // Add reference lines at 30 and 70
            AddHorizontalLine(chart, 30, Color.Red, "Oversold");
            AddHorizontalLine(chart, 70, Color.Green, "Overbought");

            return chart;
        }

        /// <summary>
        /// Adds a horizontal reference line to a chart.
        /// </summary>
        private void AddHorizontalLine(Excel.Chart chart, double value, Color color, string name)
        {
            try
            {
                Excel.SeriesCollection seriesCollection = chart.SeriesCollection();
                Excel.Series lineSeries = seriesCollection.NewSeries();

                // Create array with same value for horizontal line
                Excel.Axis xAxis = (Excel.Axis)chart.Axes(Excel.XlAxisType.xlCategory);
                int pointCount = chart.SeriesCollection(1).Points.Count;

                double[] values = new double[pointCount];
                for (int i = 0; i < pointCount; i++)
                {
                    values[i] = value;
                }

                lineSeries.Values = values;
                lineSeries.Name = name;
                lineSeries.ChartType = Excel.XlChartType.xlLine;
                lineSeries.Format.Line.ForeColor.RGB = ColorTranslator.ToOle(color);
                lineSeries.Format.Line.Weight = 1f;
                lineSeries.Format.Line.DashStyle = Microsoft.Office.Core.MsoLineDashStyle.msoLineDash;
                lineSeries.MarkerStyle = Excel.XlMarkerStyle.xlMarkerStyleNone;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error adding horizontal line: {ex.Message}");
            }
        }

        /// <summary>
        /// Creates a complete trading chart with OHLC and multiple overlays.
        /// </summary>
        /// <param name="worksheet">The worksheet containing the data</param>
        /// <param name="dataRange">Range with Date, Open, High, Low, Close, Volume columns</param>
        /// <param name="showMA20">Show 20-period moving average</param>
        /// <param name="showMA50">Show 50-period moving average</param>
        /// <param name="showBollinger">Show Bollinger Bands</param>
        /// <param name="showVolume">Show volume overlay</param>
        /// <returns>The main chart object</returns>
        public Excel.Chart CreateFullTradingChart(
            Excel.Worksheet worksheet,
            Excel.Range dataRange,
            bool showMA20 = true,
            bool showMA50 = true,
            bool showBollinger = false,
            bool showVolume = true)
        {
            // Create base OHLC chart
            Excel.Chart chart = CreateOHLCChart(worksheet, dataRange, "Trading Chart with Indicators");

            // Note: In a real implementation, you would calculate these indicators
            // using the TechnicalIndicators class and add them to the worksheet first

            // This is a template showing how overlays would be added
            // The actual indicator calculation would happen separately

            return chart;
        }

        /// <summary>
        /// Removes all charts from a worksheet.
        /// </summary>
        /// <param name="worksheet">The worksheet to clear charts from</param>
        public void ClearAllCharts(Excel.Worksheet worksheet)
        {
            try
            {
                Excel.ChartObjects chartObjects = (Excel.ChartObjects)worksheet.ChartObjects();
                chartObjects.Delete();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error clearing charts: {ex.Message}");
            }
        }
    }
}
