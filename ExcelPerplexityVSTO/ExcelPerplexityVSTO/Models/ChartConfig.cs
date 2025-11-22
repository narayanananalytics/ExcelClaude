using System;
using System.Collections.Generic;
using System.Drawing;

namespace ExcelPerplexityVSTO.Models
{
    /// <summary>
    /// Configuration options for creating charts with overlays.
    /// </summary>
    public class ChartConfig
    {
        /// <summary>
        /// The type of chart to create.
        /// </summary>
        public ChartType Type { get; set; }

        /// <summary>
        /// Title of the chart.
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        /// Label for the X-axis.
        /// </summary>
        public string XAxisTitle { get; set; }

        /// <summary>
        /// Label for the Y-axis.
        /// </summary>
        public string YAxisTitle { get; set; }

        /// <summary>
        /// Excel range address containing the data.
        /// </summary>
        public string DataRange { get; set; }

        /// <summary>
        /// Width of the chart in points.
        /// </summary>
        public double Width { get; set; }

        /// <summary>
        /// Height of the chart in points.
        /// </summary>
        public double Height { get; set; }

        /// <summary>
        /// Left position of the chart in points.
        /// </summary>
        public double Left { get; set; }

        /// <summary>
        /// Top position of the chart in points.
        /// </summary>
        public double Top { get; set; }

        /// <summary>
        /// List of overlay configurations to add to the chart.
        /// </summary>
        public List<OverlayConfig> Overlays { get; set; }

        /// <summary>
        /// Whether to show the chart legend.
        /// </summary>
        public bool ShowLegend { get; set; }

        /// <summary>
        /// Whether to show gridlines.
        /// </summary>
        public bool ShowGridlines { get; set; }

        /// <summary>
        /// Default constructor with sensible defaults.
        /// </summary>
        public ChartConfig()
        {
            Type = ChartType.OHLC;
            Title = "Chart";
            XAxisTitle = "Date";
            YAxisTitle = "Price";
            Width = 800;
            Height = 400;
            Left = 50;
            Top = 50;
            Overlays = new List<OverlayConfig>();
            ShowLegend = true;
            ShowGridlines = true;
        }
    }

    /// <summary>
    /// Types of charts that can be created.
    /// </summary>
    public enum ChartType
    {
        OHLC,
        Line,
        Column,
        Area,
        Scatter,
        Pie,
        Bar
    }

    /// <summary>
    /// Configuration for a chart overlay (e.g., Moving Average, Bollinger Bands).
    /// </summary>
    public class OverlayConfig
    {
        /// <summary>
        /// Type of overlay.
        /// </summary>
        public OverlayType Type { get; set; }

        /// <summary>
        /// Name/label for the overlay.
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Color for the overlay.
        /// </summary>
        public Color Color { get; set; }

        /// <summary>
        /// Line weight/thickness.
        /// </summary>
        public float LineWeight { get; set; }

        /// <summary>
        /// Excel range containing the overlay data.
        /// </summary>
        public string DataRange { get; set; }

        /// <summary>
        /// Whether to use secondary axis.
        /// </summary>
        public bool UseSecondaryAxis { get; set; }

        /// <summary>
        /// Period for indicators that require it (e.g., 20 for 20-period MA).
        /// </summary>
        public int? Period { get; set; }

        /// <summary>
        /// Additional parameters specific to the overlay type.
        /// </summary>
        public Dictionary<string, object> Parameters { get; set; }

        /// <summary>
        /// Default constructor.
        /// </summary>
        public OverlayConfig()
        {
            Color = Color.Blue;
            LineWeight = 2f;
            UseSecondaryAxis = false;
            Parameters = new Dictionary<string, object>();
        }
    }

    /// <summary>
    /// Types of overlays that can be added to charts.
    /// </summary>
    public enum OverlayType
    {
        /// <summary>
        /// Simple Moving Average
        /// </summary>
        SMA,

        /// <summary>
        /// Exponential Moving Average
        /// </summary>
        EMA,

        /// <summary>
        /// Bollinger Bands (upper, middle, lower)
        /// </summary>
        BollingerBands,

        /// <summary>
        /// Volume bars
        /// </summary>
        Volume,

        /// <summary>
        /// Relative Strength Index
        /// </summary>
        RSI,

        /// <summary>
        /// Moving Average Convergence Divergence
        /// </summary>
        MACD,

        /// <summary>
        /// Average True Range
        /// </summary>
        ATR,

        /// <summary>
        /// Stochastic Oscillator
        /// </summary>
        Stochastic,

        /// <summary>
        /// Custom overlay
        /// </summary>
        Custom
    }

    /// <summary>
    /// Result of a chart creation operation.
    /// </summary>
    public class ChartResult
    {
        /// <summary>
        /// Whether the operation was successful.
        /// </summary>
        public bool Success { get; set; }

        /// <summary>
        /// Error message if the operation failed.
        /// </summary>
        public string ErrorMessage { get; set; }

        /// <summary>
        /// Name of the created chart.
        /// </summary>
        public string ChartName { get; set; }

        /// <summary>
        /// Worksheet name where the chart was created.
        /// </summary>
        public string WorksheetName { get; set; }

        /// <summary>
        /// Number of overlays successfully added.
        /// </summary>
        public int OverlaysAdded { get; set; }

        public static ChartResult CreateSuccess(string chartName, string worksheetName, int overlaysAdded = 0)
        {
            return new ChartResult
            {
                Success = true,
                ChartName = chartName,
                WorksheetName = worksheetName,
                OverlaysAdded = overlaysAdded
            };
        }

        public static ChartResult CreateFailure(string errorMessage)
        {
            return new ChartResult
            {
                Success = false,
                ErrorMessage = errorMessage
            };
        }
    }
}
