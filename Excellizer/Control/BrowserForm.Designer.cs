using System.Windows.Forms;
namespace Excellizer.Control
{
    partial class BrowserForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(BrowserForm));
            this.webBrowser = new System.Windows.Forms.WebBrowser();
            this.topToolStrip = new System.Windows.Forms.ToolStrip();
            this.toolStripButton_Back = new System.Windows.Forms.ToolStripButton();
            this.toolStripButton_Forward = new System.Windows.Forms.ToolStripButton();
            this.toolStripButton_Home = new System.Windows.Forms.ToolStripButton();
            this.toolStripLabel_URL = new System.Windows.Forms.ToolStripLabel();
            this.toolStripTextBox_URL = new System.Windows.Forms.ToolStripTextBox();
            this.toolStripButton_Move = new System.Windows.Forms.ToolStripButton();
            this.toolStripButton_Refresh = new System.Windows.Forms.ToolStripButton();
            this.toolStripButton_Stop = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.toolStripButton_Detect = new System.Windows.Forms.ToolStripButton();
            this.toolStripButton_Parse = new System.Windows.Forms.ToolStripButton();
            this.backgroundWorker_Init = new System.ComponentModel.BackgroundWorker();
            this.topToolStrip.SuspendLayout();
            this.SuspendLayout();
            // 
            // webBrowser
            // 
            this.webBrowser.Dock = System.Windows.Forms.DockStyle.Fill;
            this.webBrowser.Location = new System.Drawing.Point(0, 33);
            this.webBrowser.MinimumSize = new System.Drawing.Size(600, 600);
            this.webBrowser.Name = "webBrowser";
            this.webBrowser.Size = new System.Drawing.Size(855, 700);
            this.webBrowser.TabIndex = 5;
            this.webBrowser.DocumentCompleted += new System.Windows.Forms.WebBrowserDocumentCompletedEventHandler(this.webBrowser_DocumentCompleted);
            // 
            // topToolStrip
            // 
            this.topToolStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripButton_Back,
            this.toolStripButton_Forward,
            this.toolStripButton_Home,
            this.toolStripLabel_URL,
            this.toolStripTextBox_URL,
            this.toolStripButton_Move,
            this.toolStripButton_Refresh,
            this.toolStripButton_Stop,
            this.toolStripSeparator1,
            this.toolStripButton_Detect,
            this.toolStripButton_Parse});
            this.topToolStrip.Location = new System.Drawing.Point(0, 0);
            this.topToolStrip.Name = "topToolStrip";
            this.topToolStrip.Size = new System.Drawing.Size(855, 33);
            this.topToolStrip.TabIndex = 6;
            // 
            // toolStripButton_Back
            // 
            this.toolStripButton_Back.BackColor = System.Drawing.SystemColors.ControlLight;
            this.toolStripButton_Back.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.toolStripButton_Back.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButton_Back.Margin = new System.Windows.Forms.Padding(10, 5, 0, 5);
            this.toolStripButton_Back.Name = "toolStripButton_Back";
            this.toolStripButton_Back.Size = new System.Drawing.Size(23, 23);
            this.toolStripButton_Back.Text = "<";
            this.toolStripButton_Back.Click += new System.EventHandler(this.toolStripButton_Back_Click);
            // 
            // toolStripButton_Forward
            // 
            this.toolStripButton_Forward.BackColor = System.Drawing.SystemColors.ControlLight;
            this.toolStripButton_Forward.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.toolStripButton_Forward.Image = ((System.Drawing.Image)(resources.GetObject("toolStripButton_Forward.Image")));
            this.toolStripButton_Forward.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButton_Forward.Margin = new System.Windows.Forms.Padding(10, 5, 0, 5);
            this.toolStripButton_Forward.Name = "toolStripButton_Forward";
            this.toolStripButton_Forward.Size = new System.Drawing.Size(23, 23);
            this.toolStripButton_Forward.Text = ">";
            this.toolStripButton_Forward.Click += new System.EventHandler(this.toolStripButton_Forward_Click);
            // 
            // toolStripButton_Home
            // 
            this.toolStripButton_Home.BackColor = System.Drawing.SystemColors.ControlLight;
            this.toolStripButton_Home.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.toolStripButton_Home.Image = ((System.Drawing.Image)(resources.GetObject("toolStripButton_Home.Image")));
            this.toolStripButton_Home.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButton_Home.Margin = new System.Windows.Forms.Padding(10, 5, 0, 5);
            this.toolStripButton_Home.Name = "toolStripButton_Home";
            this.toolStripButton_Home.Size = new System.Drawing.Size(23, 23);
            this.toolStripButton_Home.Text = "Home";
            this.toolStripButton_Home.Click += new System.EventHandler(this.toolStripButton_Home_Click);
            // 
            // toolStripLabel_URL
            // 
            this.toolStripLabel_URL.Margin = new System.Windows.Forms.Padding(15, 5, 0, 5);
            this.toolStripLabel_URL.Name = "toolStripLabel_URL";
            this.toolStripLabel_URL.Size = new System.Drawing.Size(28, 23);
            this.toolStripLabel_URL.Text = "URL";
            // 
            // toolStripTextBox_URL
            // 
            this.toolStripTextBox_URL.AutoSize = false;
            this.toolStripTextBox_URL.Margin = new System.Windows.Forms.Padding(10, 5, 1, 5);
            this.toolStripTextBox_URL.Name = "toolStripTextBox_URL";
            this.toolStripTextBox_URL.Size = new System.Drawing.Size(400, 23);
            this.toolStripTextBox_URL.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.toolStripTextBox_URL_KeyPress);
            // 
            // toolStripButton_Move
            // 
            this.toolStripButton_Move.BackColor = System.Drawing.SystemColors.ControlLight;
            this.toolStripButton_Move.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.toolStripButton_Move.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButton_Move.Margin = new System.Windows.Forms.Padding(20, 5, 0, 5);
            this.toolStripButton_Move.Name = "toolStripButton_Move";
            this.toolStripButton_Move.Size = new System.Drawing.Size(35, 23);
            this.toolStripButton_Move.Text = "이동";
            this.toolStripButton_Move.Click += new System.EventHandler(this.toolStripButton_Move_Click);
            // 
            // toolStripButton_Refresh
            // 
            this.toolStripButton_Refresh.BackgroundImage = global::Excellizer.Properties.Resources.viewrefresh;
            this.toolStripButton_Refresh.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.toolStripButton_Refresh.ImageAlign = System.Drawing.ContentAlignment.BottomLeft;
            this.toolStripButton_Refresh.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButton_Refresh.Margin = new System.Windows.Forms.Padding(10, 5, 0, 5);
            this.toolStripButton_Refresh.Name = "toolStripButton_Refresh";
            this.toolStripButton_Refresh.Size = new System.Drawing.Size(23, 23);
            this.toolStripButton_Refresh.Text = "Refresh";
            this.toolStripButton_Refresh.Click += new System.EventHandler(this.toolStripButton_Refresh_Click);
            // 
            // toolStripButton_Stop
            // 
            this.toolStripButton_Stop.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.toolStripButton_Stop.Image = ((System.Drawing.Image)(resources.GetObject("toolStripButton_Stop.Image")));
            this.toolStripButton_Stop.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButton_Stop.Margin = new System.Windows.Forms.Padding(10, 5, 0, 5);
            this.toolStripButton_Stop.Name = "toolStripButton_Stop";
            this.toolStripButton_Stop.Size = new System.Drawing.Size(23, 23);
            this.toolStripButton_Stop.Text = "Stop";
            this.toolStripButton_Stop.Click += new System.EventHandler(this.toolStripButton_Stop_Click);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Margin = new System.Windows.Forms.Padding(10, 0, 10, 0);
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(6, 33);
            // 
            // toolStripButton_Detect
            // 
            this.toolStripButton_Detect.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.toolStripButton_Detect.Image = ((System.Drawing.Image)(resources.GetObject("toolStripButton_Detect.Image")));
            this.toolStripButton_Detect.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButton_Detect.Name = "toolStripButton_Detect";
            this.toolStripButton_Detect.Size = new System.Drawing.Size(46, 30);
            this.toolStripButton_Detect.Text = "Detect";
            this.toolStripButton_Detect.Click += new System.EventHandler(this.toolStripButton_Detect_Click);
            // 
            // toolStripButton_Parse
            // 
            this.toolStripButton_Parse.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.toolStripButton_Parse.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButton_Parse.Margin = new System.Windows.Forms.Padding(10, 1, 0, 2);
            this.toolStripButton_Parse.Name = "toolStripButton_Parse";
            this.toolStripButton_Parse.Size = new System.Drawing.Size(39, 30);
            this.toolStripButton_Parse.Text = "Parse";
            this.toolStripButton_Parse.Click += new System.EventHandler(this.parseButton_Click);
            // 
            // backgroundWorker_Init
            // 
            this.backgroundWorker_Init.WorkerReportsProgress = true;
            this.backgroundWorker_Init.WorkerSupportsCancellation = true;
            this.backgroundWorker_Init.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWorker_Init_DoWork);
            this.backgroundWorker_Init.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.backgroundWorker_Init_ProgressChanged);
            this.backgroundWorker_Init.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.backgroundWorker_Init_RunWorkerCompleted);
            // 
            // BrowserForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(855, 733);
            this.Controls.Add(this.webBrowser);
            this.Controls.Add(this.topToolStrip);
            this.Name = "BrowserForm";
            this.Text = "Excellizer Browser";
            this.Load += new System.EventHandler(this.BrowserForm_Load);
            this.topToolStrip.ResumeLayout(false);
            this.topToolStrip.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        void webBrowser_DockChanged(object sender, System.EventArgs e)
        {
            return;
            //throw new System.NotImplementedException();
        }

        void webBrowser_LocationChanged(object sender, System.EventArgs e)
        {
            return;
            //throw new System.NotImplementedException();
        }

        void webBrowser_Navigated(object sender, System.Windows.Forms.WebBrowserNavigatedEventArgs e)
        {
             return;
            //throw new System.NotImplementedException();
        }

        #endregion

        private WebBrowser webBrowser;
        private System.Windows.Forms.ToolStrip topToolStrip;
        private System.Windows.Forms.ToolStripLabel toolStripLabel_URL;
        private System.Windows.Forms.ToolStripTextBox toolStripTextBox_URL;
        private System.Windows.Forms.ToolStripButton toolStripButton_Move;
        private System.Windows.Forms.ToolStripButton toolStripButton_Stop;
        private System.Windows.Forms.ToolStripButton toolStripButton_Home;
        private System.Windows.Forms.ToolStripButton toolStripButton_Forward;
        private System.Windows.Forms.ToolStripButton toolStripButton_Back;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripButton toolStripButton_Refresh;
        private System.Windows.Forms.ToolStripButton toolStripButton_Parse;
        private System.ComponentModel.BackgroundWorker backgroundWorker_Init;
        private ToolStripButton toolStripButton_Detect;


    }
}