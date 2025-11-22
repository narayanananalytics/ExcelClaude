using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelPerplexityVSTO.UI
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public Ribbon()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            System.Diagnostics.Debug.WriteLine($"GetCustomUI called with ribbonID: {ribbonID}");
            string xml = GetResourceText("ExcelPerplexityVSTO.UI.Ribbon.xml");
            System.Diagnostics.Debug.WriteLine($"Ribbon XML loaded: {xml != null}");
            if (xml == null)
            {
                System.Diagnostics.Debug.WriteLine("WARNING: Ribbon XML is null!");
            }
            return xml;
        }

        #endregion

        #region Ribbon Callbacks

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            System.Diagnostics.Debug.WriteLine("Ribbon_Load called - Ribbon is loading!");
            this.ribbon = ribbonUI;
        }

        /// <summary>
        /// Opens the AI Assistant task pane.
        /// </summary>
        public void ShowTaskPane_Click(Office.IRibbonControl control)
        {
            try
            {
                Globals.ThisAddIn.ShowTaskPane();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error showing task pane: {ex.Message}",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Creates an OHLC chart from the selected data range.
        /// </summary>
        public void CreateOHLCChart_Click(Office.IRibbonControl control)
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                Excel.Worksheet activeSheet = app.ActiveSheet as Excel.Worksheet;
                Excel.Range selection = app.Selection as Excel.Range;

                if (selection == null || selection.Cells.Count < 5)
                {
                    MessageBox.Show("Please select a data range with at least 5 columns (Date, Open, High, Low, Close).",
                        "Invalid Selection", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Show overlay selection dialog
                using (var dialog = new ChartOverlayDialog())
                {
                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        var chartHelper = new Helpers.ExcelChartHelper();
                        var chart = chartHelper.CreateOHLCChart(activeSheet, selection, "OHLC Chart");

                        // Add overlays based on user selection
                        if (dialog.IncludeMA20 || dialog.IncludeMA50 || dialog.IncludeVolume)
                        {
                            MessageBox.Show("Chart created! To add overlays, first calculate the indicators in your worksheet, then use the Add Overlay button.",
                                "Chart Created", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        else
                        {
                            MessageBox.Show("OHLC chart created successfully!",
                                "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error creating OHLC chart: {ex.Message}",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Inserts sample OHLC data for testing.
        /// </summary>
        public void InsertSampleData_Click(Office.IRibbonControl control)
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                Excel.Worksheet activeSheet = app.ActiveSheet as Excel.Worksheet;

                // Create headers
                activeSheet.Cells[1, 1] = "Date";
                activeSheet.Cells[1, 2] = "Open";
                activeSheet.Cells[1, 3] = "High";
                activeSheet.Cells[1, 4] = "Low";
                activeSheet.Cells[1, 5] = "Close";
                activeSheet.Cells[1, 6] = "Volume";

                // Generate sample data (30 days)
                Random rnd = new Random();
                DateTime startDate = DateTime.Now.AddDays(-30);
                double basePrice = 100.0;

                for (int i = 0; i < 30; i++)
                {
                    int row = i + 2;
                    activeSheet.Cells[row, 1] = startDate.AddDays(i).ToShortDateString();

                    double open = basePrice + rnd.Next(-5, 6);
                    double close = open + rnd.Next(-3, 4);
                    double high = Math.Max(open, close) + rnd.Next(0, 3);
                    double low = Math.Min(open, close) - rnd.Next(0, 3);
                    double volume = rnd.Next(100000, 1000000);

                    activeSheet.Cells[row, 2] = open;
                    activeSheet.Cells[row, 3] = high;
                    activeSheet.Cells[row, 4] = low;
                    activeSheet.Cells[row, 5] = close;
                    activeSheet.Cells[row, 6] = volume;

                    basePrice = close; // Next day starts where we left off
                }

                // Format as table
                Excel.Range dataRange = activeSheet.Range[activeSheet.Cells[1, 1], activeSheet.Cells[31, 6]];
                dataRange.Columns.AutoFit();

                // Bold headers
                Excel.Range headerRange = activeSheet.Range[activeSheet.Cells[1, 1], activeSheet.Cells[1, 6]];
                headerRange.Font.Bold = true;

                MessageBox.Show("Sample OHLC data inserted successfully! Select the data range and click 'Create OHLC Chart'.",
                    "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error inserting sample data: {ex.Message}",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Adds a moving average overlay to the selected chart.
        /// </summary>
        public void AddMovingAverage_Click(Office.IRibbonControl control)
        {
            try
            {
                MessageBox.Show("This feature allows you to add moving average overlays to your chart.\n\n" +
                    "To use:\n" +
                    "1. First calculate MA values in your worksheet\n" +
                    "2. Select your chart\n" +
                    "3. Use the AI Assistant to generate code for adding the overlay\n\n" +
                    "Or use the AI Assistant to generate complete code that does everything!",
                    "Add Moving Average", MessageBoxButtons.OK, MessageBoxIcon.Information);

                Globals.ThisAddIn.ShowTaskPane();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Shows settings dialog for API key configuration.
        /// </summary>
        public void Settings_Click(Office.IRibbonControl control)
        {
            try
            {
                using (var settingsDialog = new SettingsDialog())
                {
                    settingsDialog.ShowDialog();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error showing settings: {ex.Message}",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            foreach (string name in resourceNames)
            {
                if (string.Compare(resourceName, name, StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(name)))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }

    /// <summary>
    /// Dialog for selecting chart overlay options.
    /// </summary>
    public class ChartOverlayDialog : Form
    {
        public bool IncludeMA20 { get; private set; }
        public bool IncludeMA50 { get; private set; }
        public bool IncludeBollinger { get; private set; }
        public bool IncludeVolume { get; private set; }

        private CheckBox chkMA20;
        private CheckBox chkMA50;
        private CheckBox chkBollinger;
        private CheckBox chkVolume;
        private Button btnOK;
        private Button btnCancel;

        public ChartOverlayDialog()
        {
            InitializeComponents();
        }

        private void InitializeComponents()
        {
            this.Text = "Chart Overlay Options";
            this.Size = new System.Drawing.Size(350, 250);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            Label lblTitle = new Label
            {
                Text = "Select overlays to add to the OHLC chart:",
                Location = new System.Drawing.Point(20, 20),
                Size = new System.Drawing.Size(300, 20)
            };

            chkMA20 = new CheckBox
            {
                Text = "20-Period Moving Average",
                Location = new System.Drawing.Point(40, 50),
                Size = new System.Drawing.Size(250, 20)
            };

            chkMA50 = new CheckBox
            {
                Text = "50-Period Moving Average",
                Location = new System.Drawing.Point(40, 80),
                Size = new System.Drawing.Size(250, 20)
            };

            chkBollinger = new CheckBox
            {
                Text = "Bollinger Bands (20, 2)",
                Location = new System.Drawing.Point(40, 110),
                Size = new System.Drawing.Size(250, 20)
            };

            chkVolume = new CheckBox
            {
                Text = "Volume (Secondary Axis)",
                Location = new System.Drawing.Point(40, 140),
                Size = new System.Drawing.Size(250, 20),
                Checked = true
            };

            btnOK = new Button
            {
                Text = "Create Chart",
                Location = new System.Drawing.Point(150, 175),
                Size = new System.Drawing.Size(80, 25),
                DialogResult = DialogResult.OK
            };

            btnCancel = new Button
            {
                Text = "Cancel",
                Location = new System.Drawing.Point(240, 175),
                Size = new System.Drawing.Size(80, 25),
                DialogResult = DialogResult.Cancel
            };

            btnOK.Click += (s, e) =>
            {
                IncludeMA20 = chkMA20.Checked;
                IncludeMA50 = chkMA50.Checked;
                IncludeBollinger = chkBollinger.Checked;
                IncludeVolume = chkVolume.Checked;
            };

            this.Controls.Add(lblTitle);
            this.Controls.Add(chkMA20);
            this.Controls.Add(chkMA50);
            this.Controls.Add(chkBollinger);
            this.Controls.Add(chkVolume);
            this.Controls.Add(btnOK);
            this.Controls.Add(btnCancel);

            this.AcceptButton = btnOK;
            this.CancelButton = btnCancel;
        }
    }

    /// <summary>
    /// Dialog for API key and settings configuration.
    /// </summary>
    public class SettingsDialog : Form
    {
        private TextBox txtApiKey;
        private Button btnSave;
        private Button btnCancel;

        public SettingsDialog()
        {
            InitializeComponents();
            LoadSettings();
        }

        private void InitializeComponents()
        {
            this.Text = "Perplexity Settings";
            this.Size = new System.Drawing.Size(450, 150);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            Label lblApiKey = new Label
            {
                Text = "Perplexity API Key:",
                Location = new System.Drawing.Point(20, 20),
                Size = new System.Drawing.Size(120, 20)
            };

            txtApiKey = new TextBox
            {
                Location = new System.Drawing.Point(20, 45),
                Size = new System.Drawing.Size(400, 25),
                UseSystemPasswordChar = true
            };

            Label lblInfo = new Label
            {
                Text = "Get your API key from perplexity.ai/settings/api",
                Location = new System.Drawing.Point(20, 75),
                Size = new System.Drawing.Size(400, 20),
                ForeColor = System.Drawing.Color.Gray
            };

            btnSave = new Button
            {
                Text = "Save",
                Location = new System.Drawing.Point(260, 100),
                Size = new System.Drawing.Size(75, 25)
            };

            btnCancel = new Button
            {
                Text = "Cancel",
                Location = new System.Drawing.Point(345, 100),
                Size = new System.Drawing.Size(75, 25),
                DialogResult = DialogResult.Cancel
            };

            btnSave.Click += BtnSave_Click;

            this.Controls.Add(lblApiKey);
            this.Controls.Add(txtApiKey);
            this.Controls.Add(lblInfo);
            this.Controls.Add(btnSave);
            this.Controls.Add(btnCancel);

            this.AcceptButton = btnSave;
            this.CancelButton = btnCancel;
        }

        private void LoadSettings()
        {
            string apiKey = Properties.Settings.Default.PerplexityApiKey;
            if (!string.IsNullOrEmpty(apiKey))
            {
                txtApiKey.Text = apiKey;
            }
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            try
            {
                string apiKey = txtApiKey.Text.Trim();
                if (string.IsNullOrEmpty(apiKey))
                {
                    MessageBox.Show("Please enter an API key.", "Validation Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                Properties.Settings.Default.PerplexityApiKey = apiKey;
                Properties.Settings.Default.Save();

                // Update the service
                Globals.ThisAddIn.PerplexityService.SetApiKey(apiKey);

                MessageBox.Show("API key saved successfully!", "Success",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);

                this.DialogResult = DialogResult.OK;
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error saving settings: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
