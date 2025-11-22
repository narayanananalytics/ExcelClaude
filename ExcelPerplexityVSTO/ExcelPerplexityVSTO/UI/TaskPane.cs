using System;
using System.Drawing;
using System.Windows.Forms;
using ExcelPerplexityVSTO.Services;

namespace ExcelPerplexityVSTO.UI
{
    /// <summary>
    /// Task pane user control for AI-powered chat interface.
    /// </summary>
    public partial class TaskPane : UserControl
    {
        private PerplexityService perplexityService;
        private RichTextBox rtbChatHistory;
        private TextBox txtInput;
        private Button btnSend;
        private Button btnClear;
        private Label lblStatus;
        private Panel pnlInput;
        private Panel pnlButtons;

        public TaskPane()
        {
            InitializeComponent();
            SetupUI();
            // Don't access Globals.ThisAddIn in constructor - it may not be initialized yet
            this.Load += TaskPane_Load;
        }

        private void TaskPane_Load(object sender, EventArgs e)
        {
            // Initialize service after control is loaded
            if (Globals.ThisAddIn != null)
            {
                perplexityService = Globals.ThisAddIn.PerplexityService;
            }
        }

        private void SetupUI()
        {
            this.SuspendLayout();

            // Main layout
            this.Size = new Size(400, 600);
            this.BackColor = Color.White;

            // Title panel
            Panel pnlTitle = new Panel
            {
                Dock = DockStyle.Top,
                Height = 80,
                BackColor = Color.FromArgb(102, 126, 234),
                Padding = new Padding(15)
            };

            Label lblTitle = new Label
            {
                Text = "Perplexity AI Assistant",
                Font = new Font("Segoe UI", 16, FontStyle.Bold),
                ForeColor = Color.White,
                AutoSize = false,
                Size = new Size(370, 30),
                Location = new Point(15, 10)
            };

            Label lblSubtitle = new Label
            {
                Text = "C# code generator for Excel automation",
                Font = new Font("Segoe UI", 9),
                ForeColor = Color.FromArgb(230, 230, 255),
                AutoSize = false,
                Size = new Size(370, 20),
                Location = new Point(15, 42)
            };

            pnlTitle.Controls.Add(lblTitle);
            pnlTitle.Controls.Add(lblSubtitle);

            // Chat history
            rtbChatHistory = new RichTextBox
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                BackColor = Color.FromArgb(245, 245, 245),
                BorderStyle = BorderStyle.None,
                Font = new Font("Segoe UI", 10),
                Padding = new Padding(10)
            };

            // Status label
            lblStatus = new Label
            {
                Dock = DockStyle.Top,
                Height = 25,
                BackColor = Color.FromArgb(248, 249, 250),
                BorderStyle = BorderStyle.FixedSingle,
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(10, 0, 0, 0),
                Font = new Font("Segoe UI", 8),
                ForeColor = Color.Gray,
                Text = "Ready"
            };

            // Input panel
            pnlInput = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 120,
                BackColor = Color.White,
                Padding = new Padding(10)
            };

            // Input textbox
            txtInput = new TextBox
            {
                Multiline = true,
                Height = 70,
                Dock = DockStyle.Top,
                Font = new Font("Segoe UI", 10),
                BorderStyle = BorderStyle.FixedSingle
            };
            txtInput.KeyPress += TxtInput_KeyPress;

            // Buttons panel
            pnlButtons = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 40,
                Padding = new Padding(0, 10, 0, 0)
            };

            btnSend = new Button
            {
                Text = "Send",
                Width = 100,
                Height = 30,
                Location = new Point(pnlInput.Width - 220, 0),
                Anchor = AnchorStyles.Top | AnchorStyles.Right,
                BackColor = Color.FromArgb(102, 126, 234),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnSend.FlatAppearance.BorderSize = 0;
            btnSend.Click += BtnSend_Click;

            btnClear = new Button
            {
                Text = "Clear Chat",
                Width = 100,
                Height = 30,
                Location = new Point(pnlInput.Width - 110, 0),
                Anchor = AnchorStyles.Top | AnchorStyles.Right,
                BackColor = Color.FromArgb(220, 53, 69),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnClear.FlatAppearance.BorderSize = 0;
            btnClear.Click += BtnClear_Click;

            pnlButtons.Controls.Add(btnSend);
            pnlButtons.Controls.Add(btnClear);

            pnlInput.Controls.Add(pnlButtons);
            pnlInput.Controls.Add(txtInput);

            // Add controls to form
            this.Controls.Add(rtbChatHistory);
            this.Controls.Add(lblStatus);
            this.Controls.Add(pnlInput);
            this.Controls.Add(pnlTitle);

            this.ResumeLayout(false);
        }

        private async void BtnSend_Click(object sender, EventArgs e)
        {
            string userMessage = txtInput.Text.Trim();
            if (string.IsNullOrEmpty(userMessage))
                return;

            // Lazy initialize service if needed
            if (perplexityService == null && Globals.ThisAddIn != null)
            {
                perplexityService = Globals.ThisAddIn.PerplexityService;
            }

            // Check if service is available
            if (perplexityService == null)
            {
                MessageBox.Show("Service not initialized. Please try again.",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Check if API key is configured
            if (!perplexityService.IsConfigured)
            {
                MessageBox.Show("Please configure your Perplexity API key in Settings first.",
                    "API Key Required", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Add user message to chat
            AppendMessage("You", userMessage, Color.FromArgb(102, 126, 234));
            txtInput.Clear();

            // Disable controls
            btnSend.Enabled = false;
            txtInput.Enabled = false;
            lblStatus.Text = "Perplexity is thinking...";
            lblStatus.ForeColor = Color.Blue;

            try
            {
                // Get response from Perplexity
                string response = await perplexityService.SendMessageAsync(userMessage);

                // Add assistant message to chat
                AppendMessage("Perplexity", response, Color.FromArgb(76, 175, 80));

                // Extract and offer to copy C# code if present
                string code = PerplexityService.ExtractCSharpCode(response);
                if (!string.IsNullOrEmpty(code) && code != response)
                {
                    AppendCodeBlock(code);
                }

                lblStatus.Text = "Ready";
                lblStatus.ForeColor = Color.Gray;
            }
            catch (Exception ex)
            {
                AppendMessage("Error", ex.Message, Color.FromArgb(220, 53, 69));
                lblStatus.Text = "Error - check your API key and connection";
                lblStatus.ForeColor = Color.Red;
            }
            finally
            {
                btnSend.Enabled = true;
                txtInput.Enabled = true;
                txtInput.Focus();
            }
        }

        private void BtnClear_Click(object sender, EventArgs e)
        {
            var result = MessageBox.Show("Clear all chat history?", "Confirm",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                rtbChatHistory.Clear();
                if (perplexityService != null)
                {
                    perplexityService.ClearHistory();
                }
                lblStatus.Text = "Chat history cleared";
            }
        }

        private void TxtInput_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Send message on Ctrl+Enter
            if (e.KeyChar == (char)10) // Ctrl+Enter
            {
                BtnSend_Click(sender, EventArgs.Empty);
                e.Handled = true;
            }
        }

        private void AppendMessage(string sender, string message, Color color)
        {
            // Add separator if not first message
            if (rtbChatHistory.Text.Length > 0)
            {
                rtbChatHistory.AppendText("\n\n");
            }

            // Add sender name
            int startIndex = rtbChatHistory.Text.Length;
            rtbChatHistory.AppendText($"[{sender}]\n");
            int endIndex = rtbChatHistory.Text.Length;

            // Format sender name
            rtbChatHistory.Select(startIndex, endIndex - startIndex);
            rtbChatHistory.SelectionFont = new Font("Segoe UI", 10, FontStyle.Bold);
            rtbChatHistory.SelectionColor = color;

            // Add message content
            rtbChatHistory.Select(endIndex, 0);
            rtbChatHistory.SelectionFont = new Font("Segoe UI", 10);
            rtbChatHistory.SelectionColor = Color.Black;
            rtbChatHistory.AppendText(message);

            // Scroll to bottom
            rtbChatHistory.SelectionStart = rtbChatHistory.Text.Length;
            rtbChatHistory.ScrollToCaret();
        }

        private void AppendCodeBlock(string code)
        {
            rtbChatHistory.AppendText("\n\n");

            // Add "Copy Code" button instruction
            int startIndex = rtbChatHistory.Text.Length;
            rtbChatHistory.AppendText("[C# Code - Right-click to copy]\n");
            int endIndex = rtbChatHistory.Text.Length;

            rtbChatHistory.Select(startIndex, endIndex - startIndex);
            rtbChatHistory.SelectionFont = new Font("Consolas", 9, FontStyle.Bold);
            rtbChatHistory.SelectionColor = Color.FromArgb(102, 126, 234);

            // Add code
            startIndex = rtbChatHistory.Text.Length;
            rtbChatHistory.AppendText(code);
            endIndex = rtbChatHistory.Text.Length;

            rtbChatHistory.Select(startIndex, endIndex - startIndex);
            rtbChatHistory.SelectionFont = new Font("Consolas", 9);
            rtbChatHistory.SelectionColor = Color.DarkBlue;
            rtbChatHistory.SelectionBackColor = Color.FromArgb(240, 240, 240);

            // Add context menu for copying code
            ContextMenuStrip contextMenu = new ContextMenuStrip();
            ToolStripMenuItem copyItem = new ToolStripMenuItem("Copy C# Code");
            copyItem.Click += (s, e) =>
            {
                Clipboard.SetText(code);
                lblStatus.Text = "Code copied to clipboard!";
                lblStatus.ForeColor = Color.Green;
            };
            contextMenu.Items.Add(copyItem);

            rtbChatHistory.ContextMenuStrip = contextMenu;

            // Scroll to bottom
            rtbChatHistory.SelectionStart = rtbChatHistory.Text.Length;
            rtbChatHistory.ScrollToCaret();
        }

        /// <summary>
        /// Pre-fills the input with a suggested prompt.
        /// </summary>
        public void SetPrompt(string prompt)
        {
            txtInput.Text = prompt;
            txtInput.Focus();
            txtInput.SelectionStart = txtInput.Text.Length;
        }
    }
}
