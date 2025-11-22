using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace ExcelPerplexityVSTO.Services
{
    /// <summary>
    /// Service for interacting with the Perplexity API to generate C# code for Excel automation.
    /// </summary>
    public class PerplexityService
    {
        private readonly HttpClient _httpClient;
        private string _apiKey;
        private List<Message> _conversationHistory;

        private const string API_BASE_URL = "https://api.perplexity.ai";
        private const string MODEL = "sonar";

        // System prompt that teaches Perplexity to generate C# code for Excel Interop
        private const string SYSTEM_PROMPT = @"You are an expert C# programmer specializing in Excel automation using Microsoft Office Interop.

Your role is to generate C# code that users can execute in a VSTO Excel Add-in to accomplish their tasks.

IMPORTANT: Always respond with complete, ready-to-execute C# code that works with Excel Interop.

Your expertise includes:
1. Creating and formatting all types of Excel charts using Microsoft.Office.Interop.Excel
2. Creating proper OHLC/Candlestick charts with:
   - Chart type: xlStockOHLC or xlStockVOHLC (with volume)
   - Green candles for bullish periods (Close > Open) using UpBars
   - Red candles for bearish periods (Close < Open) using DownBars
   - Proper axis formatting and labels
   - Chart overlays: Moving Averages, Bollinger Bands, Volume on secondary axis
3. Working with Excel objects: Application, Workbook, Worksheet, Range, Chart, Series
4. Adding multiple series to charts for overlays
5. Using secondary axes for volume and indicators
6. Applying conditional formatting, styles, and colors
7. Creating formulas and functions programmatically
8. Data manipulation and transformation using LINQ and arrays
9. Technical indicator calculations (MA, EMA, RSI, MACD, Bollinger Bands)

When generating C# code:
- Use proper C# syntax with using statements for IDisposable objects
- Include error handling with try-catch blocks
- Use clear variable names and XML comments
- Work with Globals.ThisAddIn.Application for Excel access
- Use proper Excel Interop patterns (Marshal.ReleaseComObject when needed)
- Make the code self-contained and immediately executable
- Include ColorTranslator.ToOle() for color conversion

For OHLC/Candlestick charts specifically:
```csharp
Excel.Chart chart = chartObject.Chart;
chart.ChartType = Excel.XlChartType.xlStockOHLC;
chart.SetSourceData(dataRange);

// Format candlesticks
Excel.ChartGroup chartGroup = chart.ChartGroups(1);
chartGroup.UpBars.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(0, 200, 5)); // Green
chartGroup.DownBars.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(239, 83, 80)); // Red
```

For chart overlays (e.g., Moving Average):
```csharp
Excel.SeriesCollection seriesCollection = chart.SeriesCollection();
Excel.Series maSeries = seriesCollection.NewSeries();
maSeries.XValues = dateRange;
maSeries.Values = maValuesRange;
maSeries.Name = ""20-Period MA"";
maSeries.ChartType = Excel.XlChartType.xlLine;
maSeries.Format.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.Blue);
maSeries.Format.Line.Weight = 2f;
maSeries.MarkerStyle = Excel.XlMarkerStyle.xlMarkerStyleNone;
```

For volume on secondary axis:
```csharp
Excel.Series volumeSeries = seriesCollection.NewSeries();
volumeSeries.Values = volumeRange;
volumeSeries.ChartType = Excel.XlChartType.xlColumnClustered;
volumeSeries.AxisGroup = Excel.XlAxisGroup.xlSecondary;

// Make volume smaller (bottom 25% of chart)
Excel.Axis secondaryAxis = chart.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlSecondary);
secondaryAxis.MaximumScale = secondaryAxis.MaximumScale * 4;
```

Always provide complete, working C# code wrapped in ```csharp code blocks.";

        public class Message
        {
            [JsonProperty("role")]
            public string Role { get; set; }

            [JsonProperty("content")]
            public string Content { get; set; }
        }

        public PerplexityService()
        {
            _httpClient = new HttpClient
            {
                BaseAddress = new Uri(API_BASE_URL),
                Timeout = TimeSpan.FromSeconds(60)
            };
            _conversationHistory = new List<Message>();
        }

        /// <summary>
        /// Sets the Perplexity API key for authentication.
        /// </summary>
        /// <param name="apiKey">The API key from perplexity.ai/settings/api</param>
        public void SetApiKey(string apiKey)
        {
            if (string.IsNullOrWhiteSpace(apiKey))
                throw new ArgumentException("API key cannot be empty", nameof(apiKey));

            _apiKey = apiKey;
            _httpClient.DefaultRequestHeaders.Clear();
            _httpClient.DefaultRequestHeaders.Add("Authorization", $"Bearer {apiKey}");
        }

        /// <summary>
        /// Checks if the service is properly configured with an API key.
        /// </summary>
        public bool IsConfigured => !string.IsNullOrWhiteSpace(_apiKey);

        /// <summary>
        /// Sends a message to Perplexity and gets a C# code response.
        /// </summary>
        /// <param name="userMessage">The user's request or question</param>
        /// <returns>The AI-generated response (typically containing C# code)</returns>
        public async Task<string> SendMessageAsync(string userMessage)
        {
            if (!IsConfigured)
                throw new InvalidOperationException("API key not configured. Call SetApiKey first.");

            if (string.IsNullOrWhiteSpace(userMessage))
                throw new ArgumentException("Message cannot be empty", nameof(userMessage));

            // Add user message to history
            _conversationHistory.Add(new Message
            {
                Role = "user",
                Content = userMessage
            });

            try
            {
                // Build messages array with system prompt and conversation history
                var messages = new List<Message>
                {
                    new Message { Role = "system", Content = SYSTEM_PROMPT }
                };
                messages.AddRange(_conversationHistory);

                // Prepare request
                var requestBody = new
                {
                    model = MODEL,
                    messages = messages,
                    max_tokens = 4096,
                    temperature = 0.2,
                    top_p = 0.9,
                    stream = false
                };

                string jsonRequest = JsonConvert.SerializeObject(requestBody);
                var content = new StringContent(jsonRequest, Encoding.UTF8, "application/json");

                // Send request
                HttpResponseMessage response = await _httpClient.PostAsync("/chat/completions", content);

                // Check for errors
                if (!response.IsSuccessStatusCode)
                {
                    string errorBody = await response.Content.ReadAsStringAsync();
                    throw new HttpRequestException($"Perplexity API error ({response.StatusCode}): {errorBody}");
                }

                // Parse response
                string jsonResponse = await response.Content.ReadAsStringAsync();
                JObject responseObj = JObject.Parse(jsonResponse);

                string assistantMessage = responseObj["choices"]?[0]?["message"]?["content"]?.ToString();

                if (string.IsNullOrEmpty(assistantMessage))
                    throw new Exception("No response content received from Perplexity");

                // Add assistant response to history
                _conversationHistory.Add(new Message
                {
                    Role = "assistant",
                    Content = assistantMessage
                });

                return assistantMessage;
            }
            catch (HttpRequestException ex)
            {
                // Handle specific HTTP errors
                if (ex.Message.Contains("401"))
                    throw new Exception("Invalid API key. Please check your Perplexity API key.", ex);
                else if (ex.Message.Contains("429"))
                    throw new Exception("Rate limit exceeded. Please try again later.", ex);
                else
                    throw new Exception($"Failed to communicate with Perplexity: {ex.Message}", ex);
            }
            catch (Exception ex)
            {
                throw new Exception($"Error calling Perplexity API: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Generates C# code for creating a specific type of chart.
        /// </summary>
        /// <param name="chartType">Type of chart (e.g., "OHLC", "Line", "Bar")</param>
        /// <param name="dataRange">Excel range address containing the data</param>
        /// <param name="customization">Additional requirements or customizations</param>
        /// <returns>C# code for creating the chart</returns>
        public async Task<string> GenerateChartCodeAsync(string chartType, string dataRange, string customization = null)
        {
            StringBuilder prompt = new StringBuilder();
            prompt.Append($"Generate C# code using Excel Interop to create a {chartType} chart ");
            prompt.Append($"using data from range {dataRange}. ");

            if (chartType.ToLower().Contains("ohlc") || chartType.ToLower().Contains("candlestick"))
            {
                prompt.Append("Ensure the chart uses xlStockOHLC type with green UpBars and red DownBars. ");
            }

            if (!string.IsNullOrWhiteSpace(customization))
            {
                prompt.Append($"Additional requirements: {customization}. ");
            }

            prompt.Append("The code should work in a VSTO Excel Add-in and use Globals.ThisAddIn.Application.ActiveSheet.");

            return await SendMessageAsync(prompt.ToString());
        }

        /// <summary>
        /// Generates C# code for adding an overlay to an existing chart.
        /// </summary>
        /// <param name="overlayType">Type of overlay (e.g., "Moving Average", "Bollinger Bands", "Volume")</param>
        /// <param name="dataRange">Excel range for the overlay data</param>
        /// <param name="parameters">Parameters for the overlay (e.g., period, color)</param>
        /// <returns>C# code for adding the overlay</returns>
        public async Task<string> GenerateOverlayCodeAsync(string overlayType, string dataRange, Dictionary<string, string> parameters = null)
        {
            StringBuilder prompt = new StringBuilder();
            prompt.Append($"Generate C# code using Excel Interop to add a {overlayType} overlay ");
            prompt.Append($"to an existing chart. Use data from range {dataRange}. ");

            if (parameters != null && parameters.Count > 0)
            {
                prompt.Append("Parameters: ");
                foreach (var param in parameters)
                {
                    prompt.Append($"{param.Key}={param.Value}, ");
                }
            }

            if (overlayType.ToLower().Contains("volume"))
            {
                prompt.Append("Place the volume on a secondary axis at the bottom 25% of the chart. ");
            }

            prompt.Append("Assume the chart already exists as an Excel.Chart object named 'chart'.");

            return await SendMessageAsync(prompt.ToString());
        }

        /// <summary>
        /// Generates C# code for calculating a technical indicator.
        /// </summary>
        /// <param name="indicatorType">Type of indicator (e.g., "RSI", "MACD", "SMA", "EMA")</param>
        /// <param name="period">Period for the indicator calculation</param>
        /// <param name="sourceRange">Range containing source data (typically close prices)</param>
        /// <returns>C# code for calculating the indicator</returns>
        public async Task<string> GenerateIndicatorCodeAsync(string indicatorType, int period, string sourceRange)
        {
            string prompt = $"Generate C# code to calculate {indicatorType} with period {period} " +
                          $"using close prices from range {sourceRange}. " +
                          $"Write the calculated values to a new column next to the source data. " +
                          $"Use proper array calculations and LINQ where appropriate.";

            return await SendMessageAsync(prompt.ToString());
        }

        /// <summary>
        /// Clears the conversation history.
        /// </summary>
        public void ClearHistory()
        {
            _conversationHistory.Clear();
        }

        /// <summary>
        /// Gets the current conversation history.
        /// </summary>
        /// <returns>List of messages in the conversation</returns>
        public List<Message> GetHistory()
        {
            return new List<Message>(_conversationHistory);
        }

        /// <summary>
        /// Extracts C# code from a Perplexity response (looks for ```csharp code blocks).
        /// </summary>
        /// <param name="response">The response from Perplexity</param>
        /// <returns>Extracted C# code, or the full response if no code block found</returns>
        public static string ExtractCSharpCode(string response)
        {
            if (string.IsNullOrEmpty(response))
                return response;

            // Look for ```csharp code block
            int startIndex = response.IndexOf("```csharp", StringComparison.OrdinalIgnoreCase);
            if (startIndex == -1)
            {
                // Try without language specifier
                startIndex = response.IndexOf("```");
                if (startIndex == -1)
                    return response; // No code block found
            }

            // Find end of opening marker
            int codeStart = response.IndexOf('\n', startIndex) + 1;
            if (codeStart == 0)
                return response;

            // Find closing ```
            int codeEnd = response.IndexOf("```", codeStart);
            if (codeEnd == -1)
                return response;

            // Extract code
            return response.Substring(codeStart, codeEnd - codeStart).Trim();
        }
    }
}
