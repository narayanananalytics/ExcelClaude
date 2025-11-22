using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelPerplexityVSTO.Helpers
{
    /// <summary>
    /// Provides methods for calculating technical indicators used in trading charts.
    /// All methods work with arrays of price data and return calculated indicator values.
    /// </summary>
    public static class TechnicalIndicators
    {
        /// <summary>
        /// Calculates Simple Moving Average (SMA).
        /// </summary>
        /// <param name="values">Array of values (typically closing prices)</param>
        /// <param name="period">Number of periods for the moving average</param>
        /// <returns>Array of SMA values (first period-1 values will be null)</returns>
        public static double?[] CalculateSMA(double[] values, int period)
        {
            if (values == null || values.Length < period)
                throw new ArgumentException($"Not enough data points. Need at least {period} values.");

            double?[] sma = new double?[values.Length];

            for (int i = period - 1; i < values.Length; i++)
            {
                double sum = 0;
                for (int j = 0; j < period; j++)
                {
                    sum += values[i - j];
                }
                sma[i] = sum / period;
            }

            return sma;
        }

        /// <summary>
        /// Calculates Exponential Moving Average (EMA).
        /// </summary>
        /// <param name="values">Array of values (typically closing prices)</param>
        /// <param name="period">Number of periods for the EMA</param>
        /// <returns>Array of EMA values</returns>
        public static double?[] CalculateEMA(double[] values, int period)
        {
            if (values == null || values.Length < period)
                throw new ArgumentException($"Not enough data points. Need at least {period} values.");

            double?[] ema = new double?[values.Length];
            double multiplier = 2.0 / (period + 1);

            // First EMA value is SMA
            double sum = 0;
            for (int i = 0; i < period; i++)
            {
                sum += values[i];
            }
            ema[period - 1] = sum / period;

            // Calculate subsequent EMA values
            for (int i = period; i < values.Length; i++)
            {
                ema[i] = (values[i] - ema[i - 1].Value) * multiplier + ema[i - 1].Value;
            }

            return ema;
        }

        /// <summary>
        /// Calculates Bollinger Bands (upper, middle, lower).
        /// </summary>
        /// <param name="values">Array of values (typically closing prices)</param>
        /// <param name="period">Number of periods for the moving average (default 20)</param>
        /// <param name="standardDeviations">Number of standard deviations (default 2)</param>
        /// <returns>Tuple containing (upper band, middle band/SMA, lower band)</returns>
        public static (double?[] upper, double?[] middle, double?[] lower) CalculateBollingerBands(
            double[] values,
            int period = 20,
            double standardDeviations = 2.0)
        {
            if (values == null || values.Length < period)
                throw new ArgumentException($"Not enough data points. Need at least {period} values.");

            double?[] upper = new double?[values.Length];
            double?[] middle = new double?[values.Length];
            double?[] lower = new double?[values.Length];

            for (int i = period - 1; i < values.Length; i++)
            {
                // Calculate SMA (middle band)
                double sum = 0;
                for (int j = 0; j < period; j++)
                {
                    sum += values[i - j];
                }
                double sma = sum / period;
                middle[i] = sma;

                // Calculate standard deviation
                double variance = 0;
                for (int j = 0; j < period; j++)
                {
                    variance += Math.Pow(values[i - j] - sma, 2);
                }
                double stdDev = Math.Sqrt(variance / period);

                // Calculate upper and lower bands
                upper[i] = sma + (standardDeviations * stdDev);
                lower[i] = sma - (standardDeviations * stdDev);
            }

            return (upper, middle, lower);
        }

        /// <summary>
        /// Calculates Relative Strength Index (RSI).
        /// </summary>
        /// <param name="values">Array of values (typically closing prices)</param>
        /// <param name="period">Number of periods for RSI calculation (default 14)</param>
        /// <returns>Array of RSI values (0-100)</returns>
        public static double?[] CalculateRSI(double[] values, int period = 14)
        {
            if (values == null || values.Length < period + 1)
                throw new ArgumentException($"Not enough data points. Need at least {period + 1} values.");

            double?[] rsi = new double?[values.Length];
            double[] gains = new double[values.Length];
            double[] losses = new double[values.Length];

            // Calculate gains and losses
            for (int i = 1; i < values.Length; i++)
            {
                double change = values[i] - values[i - 1];
                gains[i] = change > 0 ? change : 0;
                losses[i] = change < 0 ? Math.Abs(change) : 0;
            }

            // First RSI calculation using SMA
            double avgGain = 0;
            double avgLoss = 0;
            for (int i = 1; i <= period; i++)
            {
                avgGain += gains[i];
                avgLoss += losses[i];
            }
            avgGain /= period;
            avgLoss /= period;

            if (avgLoss == 0)
            {
                rsi[period] = 100;
            }
            else
            {
                double rs = avgGain / avgLoss;
                rsi[period] = 100 - (100 / (1 + rs));
            }

            // Subsequent RSI values using EMA-like smoothing
            for (int i = period + 1; i < values.Length; i++)
            {
                avgGain = ((avgGain * (period - 1)) + gains[i]) / period;
                avgLoss = ((avgLoss * (period - 1)) + losses[i]) / period;

                if (avgLoss == 0)
                {
                    rsi[i] = 100;
                }
                else
                {
                    double rs = avgGain / avgLoss;
                    rsi[i] = 100 - (100 / (1 + rs));
                }
            }

            return rsi;
        }

        /// <summary>
        /// Calculates MACD (Moving Average Convergence Divergence).
        /// </summary>
        /// <param name="values">Array of values (typically closing prices)</param>
        /// <param name="fastPeriod">Fast EMA period (default 12)</param>
        /// <param name="slowPeriod">Slow EMA period (default 26)</param>
        /// <param name="signalPeriod">Signal line period (default 9)</param>
        /// <returns>Tuple containing (MACD line, Signal line, Histogram)</returns>
        public static (double?[] macd, double?[] signal, double?[] histogram) CalculateMACD(
            double[] values,
            int fastPeriod = 12,
            int slowPeriod = 26,
            int signalPeriod = 9)
        {
            if (values == null || values.Length < slowPeriod)
                throw new ArgumentException($"Not enough data points. Need at least {slowPeriod} values.");

            // Calculate fast and slow EMAs
            double?[] fastEMA = CalculateEMA(values, fastPeriod);
            double?[] slowEMA = CalculateEMA(values, slowPeriod);

            // Calculate MACD line (fast EMA - slow EMA)
            double?[] macdLine = new double?[values.Length];
            List<double> macdValues = new List<double>();

            for (int i = 0; i < values.Length; i++)
            {
                if (fastEMA[i].HasValue && slowEMA[i].HasValue)
                {
                    macdLine[i] = fastEMA[i].Value - slowEMA[i].Value;
                    macdValues.Add(macdLine[i].Value);
                }
            }

            // Calculate signal line (EMA of MACD)
            double?[] signalLine = new double?[values.Length];
            if (macdValues.Count >= signalPeriod)
            {
                double?[] signalEMA = CalculateEMA(macdValues.ToArray(), signalPeriod);

                int signalIndex = 0;
                for (int i = 0; i < values.Length; i++)
                {
                    if (macdLine[i].HasValue)
                    {
                        if (signalIndex < signalEMA.Length && signalEMA[signalIndex].HasValue)
                        {
                            signalLine[i] = signalEMA[signalIndex];
                        }
                        signalIndex++;
                    }
                }
            }

            // Calculate histogram (MACD - Signal)
            double?[] histogram = new double?[values.Length];
            for (int i = 0; i < values.Length; i++)
            {
                if (macdLine[i].HasValue && signalLine[i].HasValue)
                {
                    histogram[i] = macdLine[i].Value - signalLine[i].Value;
                }
            }

            return (macdLine, signalLine, histogram);
        }

        /// <summary>
        /// Calculates Average True Range (ATR) for volatility measurement.
        /// </summary>
        /// <param name="high">Array of high prices</param>
        /// <param name="low">Array of low prices</param>
        /// <param name="close">Array of close prices</param>
        /// <param name="period">Number of periods (default 14)</param>
        /// <returns>Array of ATR values</returns>
        public static double?[] CalculateATR(double[] high, double[] low, double[] close, int period = 14)
        {
            if (high == null || low == null || close == null)
                throw new ArgumentNullException("Price arrays cannot be null");

            if (high.Length != low.Length || high.Length != close.Length)
                throw new ArgumentException("All price arrays must have the same length");

            if (high.Length < period + 1)
                throw new ArgumentException($"Not enough data points. Need at least {period + 1} values.");

            double?[] atr = new double?[close.Length];
            double[] trueRange = new double[close.Length];

            // Calculate True Range
            for (int i = 1; i < close.Length; i++)
            {
                double tr1 = high[i] - low[i];
                double tr2 = Math.Abs(high[i] - close[i - 1]);
                double tr3 = Math.Abs(low[i] - close[i - 1]);

                trueRange[i] = Math.Max(tr1, Math.Max(tr2, tr3));
            }

            // First ATR is simple average
            double sum = 0;
            for (int i = 1; i <= period; i++)
            {
                sum += trueRange[i];
            }
            atr[period] = sum / period;

            // Subsequent ATR values using smoothing
            for (int i = period + 1; i < close.Length; i++)
            {
                atr[i] = ((atr[i - 1].Value * (period - 1)) + trueRange[i]) / period;
            }

            return atr;
        }

        /// <summary>
        /// Calculates Stochastic Oscillator (%K and %D).
        /// </summary>
        /// <param name="high">Array of high prices</param>
        /// <param name="low">Array of low prices</param>
        /// <param name="close">Array of close prices</param>
        /// <param name="kPeriod">Period for %K calculation (default 14)</param>
        /// <param name="dPeriod">Period for %D calculation (default 3)</param>
        /// <returns>Tuple containing (%K, %D)</returns>
        public static (double?[] k, double?[] d) CalculateStochastic(
            double[] high,
            double[] low,
            double[] close,
            int kPeriod = 14,
            int dPeriod = 3)
        {
            if (high == null || low == null || close == null)
                throw new ArgumentNullException("Price arrays cannot be null");

            if (high.Length != low.Length || high.Length != close.Length)
                throw new ArgumentException("All price arrays must have the same length");

            if (high.Length < kPeriod)
                throw new ArgumentException($"Not enough data points. Need at least {kPeriod} values.");

            double?[] percentK = new double?[close.Length];

            // Calculate %K
            for (int i = kPeriod - 1; i < close.Length; i++)
            {
                double highestHigh = high[i];
                double lowestLow = low[i];

                for (int j = 0; j < kPeriod; j++)
                {
                    if (high[i - j] > highestHigh) highestHigh = high[i - j];
                    if (low[i - j] < lowestLow) lowestLow = low[i - j];
                }

                if (highestHigh == lowestLow)
                {
                    percentK[i] = 50; // Avoid division by zero
                }
                else
                {
                    percentK[i] = ((close[i] - lowestLow) / (highestHigh - lowestLow)) * 100;
                }
            }

            // Calculate %D (SMA of %K)
            double?[] percentD = new double?[close.Length];
            List<double> kValues = new List<double>();

            for (int i = 0; i < percentK.Length; i++)
            {
                if (percentK[i].HasValue)
                {
                    kValues.Add(percentK[i].Value);
                }
            }

            if (kValues.Count >= dPeriod)
            {
                double?[] dSMA = CalculateSMA(kValues.ToArray(), dPeriod);

                int dIndex = 0;
                for (int i = 0; i < close.Length; i++)
                {
                    if (percentK[i].HasValue)
                    {
                        if (dIndex < dSMA.Length && dSMA[dIndex].HasValue)
                        {
                            percentD[i] = dSMA[dIndex];
                        }
                        dIndex++;
                    }
                }
            }

            return (percentK, percentD);
        }

        /// <summary>
        /// Writes indicator values to an Excel worksheet range.
        /// </summary>
        /// <param name="worksheet">The worksheet to write to</param>
        /// <param name="startRow">Starting row (1-based)</param>
        /// <param name="column">Column number (1-based)</param>
        /// <param name="values">Array of values to write</param>
        /// <param name="headerName">Optional header name for the column</param>
        public static void WriteIndicatorToWorksheet(
            Microsoft.Office.Interop.Excel.Worksheet worksheet,
            int startRow,
            int column,
            double?[] values,
            string headerName = null)
        {
            if (worksheet == null || values == null)
                throw new ArgumentNullException();

            // Write header if provided
            if (!string.IsNullOrEmpty(headerName))
            {
                worksheet.Cells[startRow, column] = headerName;
                startRow++;
            }

            // Write values
            for (int i = 0; i < values.Length; i++)
            {
                if (values[i].HasValue)
                {
                    worksheet.Cells[startRow + i, column] = values[i].Value;
                }
            }
        }
    }
}
