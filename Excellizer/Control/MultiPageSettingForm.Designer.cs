namespace Excellizer.Control
{
    partial class MultiPageSettingForm
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
            this.FrontURLTextBox = new System.Windows.Forms.TextBox();
            this.FrontURLLabel = new System.Windows.Forms.Label();
            this.PageIndexLabel = new System.Windows.Forms.Label();
            this.StartIndexTextBox = new System.Windows.Forms.TextBox();
            this.EndIndexTextBox = new System.Windows.Forms.TextBox();
            this.OkButton = new System.Windows.Forms.Button();
            this.CancelButton = new System.Windows.Forms.Button();
            this.BackURLLabel = new System.Windows.Forms.Label();
            this.BackURLTextBox = new System.Windows.Forms.TextBox();
            this.mulgeolLabel = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // FrontURLTextBox
            // 
            this.FrontURLTextBox.Location = new System.Drawing.Point(120, 21);
            this.FrontURLTextBox.Name = "FrontURLTextBox";
            this.FrontURLTextBox.Size = new System.Drawing.Size(333, 21);
            this.FrontURLTextBox.TabIndex = 0;
            // 
            // FrontURLLabel
            // 
            this.FrontURLLabel.AutoSize = true;
            this.FrontURLLabel.Location = new System.Drawing.Point(29, 26);
            this.FrontURLLabel.Name = "FrontURLLabel";
            this.FrontURLLabel.Size = new System.Drawing.Size(74, 12);
            this.FrontURLLabel.TabIndex = 1;
            this.FrontURLLabel.Text = "URL(앞부분)";
            // 
            // PageIndexLabel
            // 
            this.PageIndexLabel.AutoSize = true;
            this.PageIndexLabel.Location = new System.Drawing.Point(34, 70);
            this.PageIndexLabel.Name = "PageIndexLabel";
            this.PageIndexLabel.Size = new System.Drawing.Size(69, 12);
            this.PageIndexLabel.TabIndex = 2;
            this.PageIndexLabel.Text = "페이지 범위";
            // 
            // StartIndexTextBox
            // 
            this.StartIndexTextBox.Location = new System.Drawing.Point(120, 65);
            this.StartIndexTextBox.Name = "StartIndexTextBox";
            this.StartIndexTextBox.Size = new System.Drawing.Size(69, 21);
            this.StartIndexTextBox.TabIndex = 3;
            // 
            // EndIndexTextBox
            // 
            this.EndIndexTextBox.Location = new System.Drawing.Point(215, 65);
            this.EndIndexTextBox.Name = "EndIndexTextBox";
            this.EndIndexTextBox.Size = new System.Drawing.Size(69, 21);
            this.EndIndexTextBox.TabIndex = 5;
            // 
            // OkButton
            // 
            this.OkButton.Location = new System.Drawing.Point(297, 158);
            this.OkButton.Name = "OkButton";
            this.OkButton.Size = new System.Drawing.Size(75, 23);
            this.OkButton.TabIndex = 10;
            this.OkButton.Text = "OK";
            this.OkButton.UseVisualStyleBackColor = true;
            this.OkButton.Click += new System.EventHandler(this.OkButton_Click);
            // 
            // CancelButton
            // 
            this.CancelButton.Location = new System.Drawing.Point(378, 158);
            this.CancelButton.Name = "CancelButton";
            this.CancelButton.Size = new System.Drawing.Size(75, 23);
            this.CancelButton.TabIndex = 11;
            this.CancelButton.Text = "Cancel";
            this.CancelButton.UseVisualStyleBackColor = true;
            this.CancelButton.Click += new System.EventHandler(this.CancelButton_Click);
            // 
            // BackURLLabel
            // 
            this.BackURLLabel.AutoSize = true;
            this.BackURLLabel.Location = new System.Drawing.Point(29, 114);
            this.BackURLLabel.Name = "BackURLLabel";
            this.BackURLLabel.Size = new System.Drawing.Size(74, 12);
            this.BackURLLabel.TabIndex = 13;
            this.BackURLLabel.Text = "URL(뒷부분)";
            // 
            // BackURLTextBox
            // 
            this.BackURLTextBox.Location = new System.Drawing.Point(120, 109);
            this.BackURLTextBox.Name = "BackURLTextBox";
            this.BackURLTextBox.Size = new System.Drawing.Size(333, 21);
            this.BackURLTextBox.TabIndex = 12;
            // 
            // mulgeolLabel
            // 
            this.mulgeolLabel.AutoSize = true;
            this.mulgeolLabel.Location = new System.Drawing.Point(195, 70);
            this.mulgeolLabel.Name = "mulgeolLabel";
            this.mulgeolLabel.Size = new System.Drawing.Size(14, 12);
            this.mulgeolLabel.TabIndex = 14;
            this.mulgeolLabel.Text = "~";
            // 
            // MultiPageSettingForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(487, 195);
            this.Controls.Add(this.mulgeolLabel);
            this.Controls.Add(this.BackURLLabel);
            this.Controls.Add(this.BackURLTextBox);
            this.Controls.Add(this.CancelButton);
            this.Controls.Add(this.OkButton);
            this.Controls.Add(this.EndIndexTextBox);
            this.Controls.Add(this.StartIndexTextBox);
            this.Controls.Add(this.PageIndexLabel);
            this.Controls.Add(this.FrontURLLabel);
            this.Controls.Add(this.FrontURLTextBox);
            this.MaximizeBox = false;
            this.Name = "MultiPageSettingForm";
            this.Text = "MultiPageSettingForm";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox FrontURLTextBox;
        private System.Windows.Forms.Label FrontURLLabel;
        private System.Windows.Forms.Label PageIndexLabel;
        private System.Windows.Forms.TextBox StartIndexTextBox;
        private System.Windows.Forms.TextBox EndIndexTextBox;
        private System.Windows.Forms.Button OkButton;
        private System.Windows.Forms.Button CancelButton;
        private System.Windows.Forms.Label BackURLLabel;
        private System.Windows.Forms.TextBox BackURLTextBox;
        private System.Windows.Forms.Label mulgeolLabel;
    }
}