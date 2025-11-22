using System;
using System.Windows.Forms;
using Microsoft.Office.Tools;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using ExcelPerplexityVSTO.Services;
using ExcelPerplexityVSTO.UI;

namespace ExcelPerplexityVSTO
{
    public partial class ThisAddIn
    {
        private CustomTaskPane taskPane;
        private PerplexityService perplexityService;

        /// <summary>
        /// Gets the Perplexity service instance.
        /// </summary>
        public PerplexityService PerplexityService
        {
            get { return perplexityService; }
        }

        /// <summary>
        /// Called when the add-in is loaded.
        /// </summary>
        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            try
            {
                // Initialize Perplexity service
                perplexityService = new PerplexityService();

                // Load API key from settings if available
                string apiKey = Properties.Settings.Default.PerplexityApiKey;
                if (!string.IsNullOrEmpty(apiKey))
                {
                    perplexityService.SetApiKey(apiKey);
                }

                // Create task pane
                TaskPane taskPaneControl = new TaskPane();
                taskPane = this.CustomTaskPanes.Add(taskPaneControl, "Perplexity AI Assistant");
                taskPane.Width = 400;
                taskPane.Visible = false;

                // Log startup
                System.Diagnostics.Debug.WriteLine("Excel Perplexity VSTO Add-in started successfully");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error starting add-in: {ex.Message}\n\n{ex.StackTrace}",
                    "Startup Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Called when the add-in is shut down.
        /// </summary>
        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            try
            {
                // Clean up resources
                if (taskPane != null)
                {
                    taskPane.Dispose();
                    taskPane = null;
                }

                System.Diagnostics.Debug.WriteLine("Excel Perplexity VSTO Add-in shut down successfully");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error during shutdown: {ex.Message}");
            }
        }

        /// <summary>
        /// Shows the task pane.
        /// </summary>
        public void ShowTaskPane()
        {
            if (taskPane != null)
            {
                taskPane.Visible = true;
            }
        }

        /// <summary>
        /// Hides the task pane.
        /// </summary>
        public void HideTaskPane()
        {
            if (taskPane != null)
            {
                taskPane.Visible = false;
            }
        }

        /// <summary>
        /// Toggles the task pane visibility.
        /// </summary>
        public void ToggleTaskPane()
        {
            if (taskPane != null)
            {
                taskPane.Visible = !taskPane.Visible;
            }
        }

        /// <summary>
        /// Sets a prompt in the task pane input.
        /// </summary>
        /// <param name="prompt">The prompt to set</param>
        public void SetTaskPanePrompt(string prompt)
        {
            if (taskPane != null && taskPane.Control is TaskPane taskPaneControl)
            {
                taskPane.Visible = true;
                taskPaneControl.SetPrompt(prompt);
            }
        }


        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
