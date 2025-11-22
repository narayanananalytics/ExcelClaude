using System;

namespace ExcelPerplexityVSTO.Models
{
    /// <summary>
    /// Represents a single OHLC (Open-High-Low-Close) data point for financial charts.
    /// </summary>
    public class OHLCData
    {
        /// <summary>
        /// The date/time of this data point.
        /// </summary>
        public DateTime Date { get; set; }

        /// <summary>
        /// The opening price.
        /// </summary>
        public double Open { get; set; }

        /// <summary>
        /// The highest price during the period.
        /// </summary>
        public double High { get; set; }

        /// <summary>
        /// The lowest price during the period.
        /// </summary>
        public double Low { get; set; }

        /// <summary>
        /// The closing price.
        /// </summary>
        public double Close { get; set; }

        /// <summary>
        /// The trading volume (optional).
        /// </summary>
        public double? Volume { get; set; }

        /// <summary>
        /// Indicates if this is a bullish period (Close > Open).
        /// </summary>
        public bool IsBullish => Close >= Open;

        /// <summary>
        /// Indicates if this is a bearish period (Close < Open).
        /// </summary>
        public bool IsBearish => Close < Open;

        /// <summary>
        /// The body size (absolute difference between Open and Close).
        /// </summary>
        public double BodySize => Math.Abs(Close - Open);

        /// <summary>
        /// The total range (High - Low).
        /// </summary>
        public double Range => High - Low;

        /// <summary>
        /// The upper wick/shadow length.
        /// </summary>
        public double UpperWick => High - Math.Max(Open, Close);

        /// <summary>
        /// The lower wick/shadow length.
        /// </summary>
        public double LowerWick => Math.Min(Open, Close) - Low;

        /// <summary>
        /// Default constructor.
        /// </summary>
        public OHLCData()
        {
        }

        /// <summary>
        /// Constructor with all required parameters.
        /// </summary>
        public OHLCData(DateTime date, double open, double high, double low, double close, double? volume = null)
        {
            Date = date;
            Open = open;
            High = high;
            Low = low;
            Close = close;
            Volume = volume;
        }

        /// <summary>
        /// Validates that the OHLC data is consistent (High >= Low, Open/Close within range).
        /// </summary>
        /// <returns>True if data is valid, false otherwise</returns>
        public bool IsValid()
        {
            if (High < Low)
                return false;

            if (Open > High || Open < Low)
                return false;

            if (Close > High || Close < Low)
                return false;

            if (Volume.HasValue && Volume.Value < 0)
                return false;

            return true;
        }

        public override string ToString()
        {
            return $"{Date:yyyy-MM-dd} O:{Open:F2} H:{High:F2} L:{Low:F2} C:{Close:F2}" +
                   (Volume.HasValue ? $" V:{Volume:F0}" : "");
        }
    }
}
