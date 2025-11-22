# Excel Perplexity Add-in

An intelligent Excel Add-in powered by Perplexity AI that generates VBA macros to automate Excel tasks. Ask for any Excel feature and get ready-to-execute VBA code!

## Features

- **VBA Macro Generation**: Perplexity AI generates complete, ready-to-execute VBA macros for any Excel task
- **Full Excel API Access**: Unlike Office.js limitations, VBA gives you access to ALL Excel features
- **OHLC/Candlestick Charts**: Generate proper OHLC charts with:
  - Green candles for bullish movements (Close > Open)
  - Red candles for bearish movements (Close < Open)
  - Full control over chart formatting and styling
- **One-Click Copy**: Copy generated VBA code to clipboard with a single click
- **Execution Instructions**: Built-in guide shows you how to run the macros in Excel
- **Conversational Interface**: Describe what you want in plain English, get working VBA code
- **Any Excel Feature**: Create charts, format data, build dashboards, automate workflows - anything Excel VBA can do!

## Prerequisites

- Node.js (v16 or higher)
- npm or yarn
- Microsoft Excel (Desktop version recommended)
- Perplexity API key from [perplexity.ai/settings/api](https://www.perplexity.ai/settings/api)

## Installation

1. Navigate to the project directory:
```bash
cd ExcelClaude
```

2. Install dependencies:
```bash
npm install
```

3. Build the project:
```bash
npm run build:dev
```

4. Generate SSL certificates for local development (required for HTTPS):
```bash
npx office-addin-dev-certs install
```

Note: If prompted, approve the certificate installation.

## Development

Start the development server:
```bash
npm run dev
```

This will start a webpack dev server at `https://localhost:3000`.

## Loading the Add-in in Excel

### Method 1: Using Office Add-in Debugging (Recommended)

1. Start the dev server:
```bash
npm run dev
```

2. In a new terminal, sideload the add-in:
```bash
npm start
```

This will automatically open Excel with your add-in loaded.

### Method 2: Manual Sideloading

1. Open Excel
2. Go to Insert > My Add-ins > Shared Folder
3. Browse to the manifest.xml file in your project directory
4. Select it to load the add-in

## Usage

### Initial Setup

1. Once the add-in is loaded, open it from the Home tab > Perplexity AI group > Perplexity Assistant button
2. Enter your Perplexity API key in the configuration section
3. Click "Set API Key" to save it

### Generating VBA Macros

#### For OHLC Charts:
1. Click "Insert Sample OHLC Data" to add sample data to your worksheet
2. Select the data range (including headers)
3. Click "Generate OHLC Chart VBA"
4. Perplexity will generate a complete VBA macro
5. Click "📋 Copy Code" to copy the macro
6. Press Alt+F11 in Excel to open the VBA Editor
7. Insert → Module, paste the code
8. Press Alt+F8, select the macro, and click Run

#### For Any Excel Task:
Simply ask in plain English:
- "Create a pivot table from my data"
- "Format all negative values in red"
- "Add conditional formatting to highlight duplicates"
- "Create a dashboard with multiple charts"
- "Automate my monthly report generation"

Perplexity will generate complete VBA code that you can copy and execute!

### Chat with Perplexity

Ask Perplexity to generate VBA code for any Excel task:

**Chart Examples:**
- "Create a waterfall chart from my data"
- "Make a combo chart with bars and line"
- "Generate a heat map using conditional formatting"

**Data Manipulation:**
- "Sort my data by column C descending"
- "Remove duplicates from the selected range"
- "Split text in column A by delimiter"

**Formatting:**
- "Apply zebra striping to my table"
- "Format dates as MM/DD/YYYY"
- "Highlight cells containing specific text"

**Automation:**
- "Create a macro to save this sheet as PDF"
- "Automate copying data from multiple sheets"
- "Build a custom toolbar with my frequent tasks"

Every response includes ready-to-execute VBA code!

### Quick Actions

- **Insert Sample OHLC Data**: Adds sample stock data to test with
- **Generate OHLC Chart VBA**: Creates VBA code for a properly formatted OHLC candlestick chart
- **Ask Perplexity about Selected Range**: Generates VBA code for your selected data range

### How to Execute Generated VBA Macros

1. Click the **📋 Copy Code** button on any generated VBA macro
2. In Excel, press **Alt+F11** to open the VBA Editor
3. Go to **Insert → Module**
4. Paste the code into the module window
5. Close the VBA Editor (**Alt+Q**)
6. Press **Alt+F8** to open the Macro dialog
7. Select your macro and click **Run**

**Enable Macros** (if needed):
- File → Options → Trust Center → Trust Center Settings
- Macro Settings → Enable all macros (or "with notification")

## Project Structure

```
ExcelClaude/
├── src/
│   ├── taskpane/
│   │   ├── components/
│   │   │   └── App.tsx          # Main React component
│   │   ├── taskpane.tsx         # Entry point
│   │   ├── taskpane.html        # HTML template
│   │   └── taskpane.css         # Styles
│   ├── commands/
│   │   ├── commands.ts          # Command functions
│   │   └── commands.html        # Commands HTML
│   ├── services/
│   │   └── perplexityService.ts # Perplexity API integration
│   └── utils/
│       └── excelHelpers.ts      # Excel API helper functions
├── assets/                      # Icons and images
├── manifest.xml                 # Add-in manifest
├── webpack.config.js            # Webpack configuration
├── tsconfig.json               # TypeScript configuration
└── package.json                # Project dependencies
```

## Key Components

### PerplexityService (`src/services/perplexityService.ts`)

Handles communication with the Perplexity API:
- Manages conversation history
- System prompt configured to generate VBA macros
- Instructs AI to create ready-to-execute VBA code
- Uses Perplexity's llama-3.1-sonar-small-128k-online model

### ExcelHelpers (`src/utils/excelHelpers.ts`)

Excel API wrapper functions for basic operations:
- `getSelectedRange()`: Gets current selection address
- `insertSampleOHLCData()`: Inserts test OHLC data
- `getRangeData()`: Retrieves range values
- `validateOHLCData()`: Validates data structure

### VBAHelpers (`src/utils/vbaHelpers.ts`)

VBA code handling utilities:
- `extractVBACode()`: Extracts VBA from Perplexity responses
- `copyToClipboard()`: Copies code to clipboard
- `generateOHLCChartVBA()`: Pre-built OHLC macro generator
- `getExecutionInstructions()`: Provides step-by-step execution guide

### App Component (`src/taskpane/components/App.tsx`)

Main UI component:
- API key configuration
- Chat interface
- Quick action buttons
- Message history management

## VBA-Generated OHLC Chart Specifications

Generated VBA macros create OHLC charts with:

- **Chart Type**: xlStockOHLC (proper candlestick/OHLC bars)
- **Bullish Candles**: Green (RGB(0, 200, 5)) - when Close > Open
- **Bearish Candles**: Red (RGB(239, 83, 80)) - when Close < Open
- **UpBars/DownBars**: Properly formatted for visual distinction
- **Dimensions**: 600 × 400 (customizable in VBA)
- **Legend**: Bottom position
- **Axes**: Formatted with titles and labels
- **Full Control**: Modify any aspect via generated VBA code

### Why VBA Instead of Office.js?

**Office.js Limitations:**
- Limited chart type support
- Restricted formatting options
- Cannot create true candlestick charts
- Limited access to advanced Excel features

**VBA Advantages:**
- Full access to ALL Excel features
- Proper OHLC/candlestick chart support
- Complete formatting control
- Can automate complex workflows
- Industry-standard Excel automation

## Security Notes

**IMPORTANT**: This add-in makes API calls directly from the browser, which is suitable for development and personal use. For production deployments:

1. Create a backend proxy service to handle Perplexity API calls
2. Never expose your API key in client-side code
3. Implement proper authentication and rate limiting
4. Use environment variables for sensitive configuration
5. Consider implementing API key rotation and usage monitoring

## Troubleshooting

### Add-in doesn't load
- Ensure you've installed SSL certificates: `npx office-addin-dev-certs install`
- Check that the dev server is running on https://localhost:3000
- Verify manifest.xml has correct URLs

### Charts not formatting correctly
- Ensure data has proper headers (Open, High, Low, Close)
- Check that numeric values are not stored as text
- Verify the selected range includes headers

### Perplexity API errors
- Verify your API key is correct
- Check your API quota at perplexity.ai/settings/api
- Ensure you have an active internet connection
- Check for rate limiting (Perplexity has usage limits based on your plan)

## Building for Production

```bash
npm run build
```

The production build will be in the `dist/` directory.

## Scripts

- `npm run dev`: Start development server
- `npm run build`: Production build
- `npm run build:dev`: Development build
- `npm start`: Sideload add-in in Excel
- `npm stop`: Stop debugging session
- `npm run validate`: Validate manifest.xml

## Technologies Used

- **TypeScript**: Type-safe development
- **React**: UI framework
- **Webpack**: Module bundler
- **Office.js**: Excel JavaScript API
- **Axios**: HTTP client for Perplexity API integration

## License

MIT

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## Support

For issues and questions:
1. Check the troubleshooting section
2. Review Office Add-ins documentation
3. Check Perplexity API documentation
4. Open an issue in this repository

## Roadmap

Future enhancements:
- [ ] Additional chart types (candlestick, volume, etc.)
- [ ] Custom color schemes
- [ ] Chart templates
- [ ] Export chart configurations
- [ ] Advanced technical indicators
- [ ] Multi-chart dashboards
- [ ] Backend proxy for API calls
- [ ] Offline mode with cached responses
