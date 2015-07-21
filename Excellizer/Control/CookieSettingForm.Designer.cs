namespace Excellizer.Control
{
    partial class CookieSettingForm
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
            this.FirefoxButton = new System.Windows.Forms.Button();
            this.IEButton = new System.Windows.Forms.Button();
            this.ChromeButton = new System.Windows.Forms.Button();
            this.ChromePictureBox = new System.Windows.Forms.PictureBox();
            this.IEPictureBox = new System.Windows.Forms.PictureBox();
            this.FirefoxPictureBox = new System.Windows.Forms.PictureBox();
            this.OKButton = new System.Windows.Forms.Button();
            this.CancelButton = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.ChromePictureBox)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.IEPictureBox)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.FirefoxPictureBox)).BeginInit();
            this.SuspendLayout();
            // 
            // FirefoxButton
            // 
            this.FirefoxButton.Location = new System.Drawing.Point(100, 146);
            this.FirefoxButton.Name = "FirefoxButton";
            this.FirefoxButton.Size = new System.Drawing.Size(75, 23);
            this.FirefoxButton.TabIndex = 12;
            this.FirefoxButton.Text = "Firefox";
            this.FirefoxButton.UseVisualStyleBackColor = true;
            this.FirefoxButton.Click += new System.EventHandler(this.FirefoxButton_Click);
            // 
            // IEButton
            // 
            this.IEButton.Location = new System.Drawing.Point(100, 92);
            this.IEButton.Name = "IEButton";
            this.IEButton.Size = new System.Drawing.Size(75, 23);
            this.IEButton.TabIndex = 11;
            this.IEButton.Text = "IE";
            this.IEButton.UseVisualStyleBackColor = true;
            this.IEButton.Click += new System.EventHandler(this.IEButton_Click);
            // 
            // ChromeButton
            // 
            this.ChromeButton.Location = new System.Drawing.Point(100, 39);
            this.ChromeButton.Name = "ChromeButton";
            this.ChromeButton.Size = new System.Drawing.Size(75, 23);
            this.ChromeButton.TabIndex = 10;
            this.ChromeButton.Text = "Chrome";
            this.ChromeButton.UseVisualStyleBackColor = true;
            this.ChromeButton.Click += new System.EventHandler(this.ChromeButton_Click);
            // 
            // ChromePictureBox
            // 
            this.ChromePictureBox.Image = global::Excellizer.Properties.Resources.ChromeImage;
            this.ChromePictureBox.Location = new System.Drawing.Point(39, 34);
            this.ChromePictureBox.Name = "ChromePictureBox";
            this.ChromePictureBox.Size = new System.Drawing.Size(30, 30);
            this.ChromePictureBox.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.ChromePictureBox.TabIndex = 13;
            this.ChromePictureBox.TabStop = false;
            // 
            // IEPictureBox
            // 
            this.IEPictureBox.Image = global::Excellizer.Properties.Resources.IEImage;
            this.IEPictureBox.Location = new System.Drawing.Point(39, 87);
            this.IEPictureBox.Name = "IEPictureBox";
            this.IEPictureBox.Size = new System.Drawing.Size(30, 30);
            this.IEPictureBox.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.IEPictureBox.TabIndex = 14;
            this.IEPictureBox.TabStop = false;
            // 
            // FirefoxPictureBox
            // 
            this.FirefoxPictureBox.Image = global::Excellizer.Properties.Resources.FirefoxImage;
            this.FirefoxPictureBox.Location = new System.Drawing.Point(39, 141);
            this.FirefoxPictureBox.Name = "FirefoxPictureBox";
            this.FirefoxPictureBox.Size = new System.Drawing.Size(30, 30);
            this.FirefoxPictureBox.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.FirefoxPictureBox.TabIndex = 15;
            this.FirefoxPictureBox.TabStop = false;
            // 
            // OKButton
            // 
            this.OKButton.Location = new System.Drawing.Point(21, 208);
            this.OKButton.Name = "OKButton";
            this.OKButton.Size = new System.Drawing.Size(75, 23);
            this.OKButton.TabIndex = 16;
            this.OKButton.Text = "OK";
            this.OKButton.UseVisualStyleBackColor = true;
            this.OKButton.Click += new System.EventHandler(this.OKButton_Click);
            // 
            // CancelButton
            // 
            this.CancelButton.Location = new System.Drawing.Point(111, 208);
            this.CancelButton.Name = "CancelButton";
            this.CancelButton.Size = new System.Drawing.Size(75, 23);
            this.CancelButton.TabIndex = 17;
            this.CancelButton.Text = "Cancel";
            this.CancelButton.UseVisualStyleBackColor = true;
            this.CancelButton.Click += new System.EventHandler(this.CancelButton_Click);
            // 
            // CookieSettingForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(211, 243);
            this.Controls.Add(this.CancelButton);
            this.Controls.Add(this.OKButton);
            this.Controls.Add(this.FirefoxPictureBox);
            this.Controls.Add(this.IEPictureBox);
            this.Controls.Add(this.ChromePictureBox);
            this.Controls.Add(this.FirefoxButton);
            this.Controls.Add(this.IEButton);
            this.Controls.Add(this.ChromeButton);
            this.MaximizeBox = false;
            this.Name = "CookieSettingForm";
            this.Text = "CookieSettingForm";
            ((System.ComponentModel.ISupportInitialize)(this.ChromePictureBox)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.IEPictureBox)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.FirefoxPictureBox)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button FirefoxButton;
        private System.Windows.Forms.Button IEButton;
        private System.Windows.Forms.Button ChromeButton;
        private System.Windows.Forms.PictureBox ChromePictureBox;
        private System.Windows.Forms.PictureBox IEPictureBox;
        private System.Windows.Forms.PictureBox FirefoxPictureBox;
        private System.Windows.Forms.Button OKButton;
        private System.Windows.Forms.Button CancelButton;
    }
}