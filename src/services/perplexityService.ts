import axios, { AxiosInstance } from 'axios';

export interface PerplexityMessage {
  role: 'user' | 'assistant' | 'system';
  content: string;
}

export interface ChartRequest {
  type: 'ohlc' | 'line' | 'bar' | 'pie' | 'scatter';
  dataRange?: string;
  customization?: string;
}

class PerplexityService {
  private client: AxiosInstance | null = null;
  private conversationHistory: PerplexityMessage[] = [];
  private apiKey: string = '';

  // System prompt that teaches Perplexity to generate VBA macros
  private readonly SYSTEM_PROMPT = `You are an expert Excel VBA programmer with deep knowledge of Excel automation, charting, and data manipulation.

Your role is to generate VBA macro code that users can execute directly in Excel to accomplish their tasks.

IMPORTANT: Always respond with complete, ready-to-execute VBA code wrapped in a code block.

Your expertise includes:
1. Creating and formatting all types of Excel charts using VBA Chart objects
2. Creating proper OHLC/Candlestick charts with:
   - Chart type: xlStockOHLC or xlStockVOHLC (with volume)
   - Green candles for bullish periods (Close > Open)
   - Red candles for bearish periods (Close < Open)
   - Proper axis formatting and labels
3. Working with Range, Worksheet, and Workbook objects
4. Applying conditional formatting, styles, and colors
5. Creating formulas and functions programmatically
6. Data manipulation and transformation
7. Custom formatting and styling

When generating VBA code:
- Start with "Sub GeneratedMacro()" and end with "End Sub"
- Include error handling with "On Error Resume Next" or proper error handlers
- Use clear variable names and comments
- Work with the active worksheet unless specified otherwise
- Use Selection object when appropriate for user-selected ranges
- Include proper object cleanup (set objects = Nothing)
- Make the code self-contained and immediately executable

For OHLC/Candlestick charts specifically:
- Use Chart.ChartType = xlStockOHLC
- Data structure: Date, Open, High, Low, Close (Volume optional)
- Format UpBars and DownBars for proper coloring:
  - UpBars.Interior.Color = RGB(0, 200, 5) 'Green
  - DownBars.Interior.Color = RGB(239, 83, 80) 'Red
- Set proper axis titles and formatting
- Position chart appropriately on the worksheet

Example response format:
\`\`\`vba
Sub GeneratedMacro()
    ' Description of what this macro does
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' Your code here

    Set ws = Nothing
End Sub
\`\`\`

Always provide complete, working VBA code that can be copied and executed immediately.`;

  constructor() {
    // Client will be initialized when API key is provided
  }

  setApiKey(apiKey: string): void {
    this.apiKey = apiKey;
    this.client = axios.create({
      baseURL: 'https://api.perplexity.ai',
      headers: {
        'Authorization': `Bearer ${apiKey}`,
        'Content-Type': 'application/json'
      }
    });
  }

  isConfigured(): boolean {
    return this.client !== null && this.apiKey !== '';
  }

  async sendMessage(userMessage: string): Promise<string> {
    if (!this.client) {
      throw new Error('Perplexity API key not configured. Please set your API key first.');
    }

    // Add user message to history
    this.conversationHistory.push({
      role: 'user',
      content: userMessage
    });

    try {
      // Perplexity API format
      const messages: PerplexityMessage[] = [
        {
          role: 'system',
          content: this.SYSTEM_PROMPT
        },
        ...this.conversationHistory
      ];

      const response = await this.client.post('/chat/completions', {
        model: 'sonar',
        messages: messages,
        max_tokens: 4096,
        temperature: 0.2,
        top_p: 0.9,
        stream: false
      });

      const assistantMessage = response.data.choices[0].message.content;

      // Add assistant response to history
      this.conversationHistory.push({
        role: 'assistant',
        content: assistantMessage
      });

      return assistantMessage;
    } catch (error) {
      console.error('Error calling Perplexity API:', error);
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';

      // Check for specific error types
      if (axios.isAxiosError(error)) {
        if (error.response?.status === 401) {
          throw new Error('Invalid API key. Please check your Perplexity API key.');
        } else if (error.response?.status === 429) {
          throw new Error('Rate limit exceeded. Please try again later.');
        } else if (error.response?.data?.error?.message) {
          throw new Error(`Perplexity API error: ${error.response.data.error.message}`);
        }
      }

      throw new Error(`Failed to get response from Perplexity: ${errorMessage}`);
    }
  }

  async getChartInstructions(request: ChartRequest): Promise<string> {
    let prompt = `I need to create a ${request.type} chart in Excel.`;

    if (request.dataRange) {
      prompt += ` The data is in range ${request.dataRange}.`;
    }

    if (request.customization) {
      prompt += ` Additional requirements: ${request.customization}`;
    }

    if (request.type === 'ohlc') {
      prompt += ` Please ensure the chart follows standard OHLC formatting with green for bullish candles (close > open) and red for bearish candles (close < open).`;
    }

    prompt += ` Provide Excel JavaScript API code to create and format this chart properly.`;

    return this.sendMessage(prompt);
  }

  clearHistory(): void {
    this.conversationHistory = [];
  }

  getHistory(): PerplexityMessage[] {
    return [...this.conversationHistory];
  }
}

// Singleton instance
export const perplexityService = new PerplexityService();
