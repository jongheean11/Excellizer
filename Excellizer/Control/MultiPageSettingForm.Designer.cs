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
            this.URLTextBox = new System.Windows.Forms.TextBox();
            this.URLLabel = new System.Windows.Forms.Label();
            this.StartIndexLabel = new System.Windows.Forms.Label();
            this.StartIndexTextBox = new System.Windows.Forms.TextBox();
            this.EndIndexLabel = new System.Windows.Forms.Label();
            this.EndIndexTextBox = new System.Windows.Forms.TextBox();
            this.OkButton = new System.Windows.Forms.Button();
            this.CancelButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // URLTextBox
            // 
            this.URLTextBox.Location = new System.Drawing.Point(136, 25);
            this.URLTextBox.Name = "URLTextBox";
            this.URLTextBox.Size = new System.Drawing.Size(333, 21);
            this.URLTextBox.TabIndex = 0;
            // 
            // URLLabel
            // 
            this.URLLabel.AutoSize = true;
            this.URLLabel.Location = new System.Drawing.Point(30, 30);
            this.URLLabel.Name = "URLLabel";
            this.URLLabel.Size = new System.Drawing.Size(98, 12);
            this.URLLabel.TabIndex = 1;
            this.URLLabel.Text = "URL(시작페이지)";
            // 
            // StartIndexLabel
            // 
            this.StartIndexLabel.AutoSize = true;
            this.StartIndexLabel.Location = new System.Drawing.Point(27, 74);
            this.StartIndexLabel.Name = "StartIndexLabel";
            this.StartIndexLabel.Size = new System.Drawing.Size(103, 12);
            this.StartIndexLabel.TabIndex = 2;
            this.StartIndexLabel.Text = "시작 페이지(번호)";
            // 
            // StartIndexTextBox
            // 
            this.StartIndexTextBox.Location = new System.Drawing.Point(136, 69);
            this.StartIndexTextBox.Name = "StartIndexTextBox";
            this.StartIndexTextBox.Size = new System.Drawing.Size(69, 21);
            this.StartIndexTextBox.TabIndex = 3;
            // 
            // EndIndexLabel
            // 
            this.EndIndexLabel.AutoSize = true;
            this.EndIndexLabel.Location = new System.Drawing.Point(39, 115);
            this.EndIndexLabel.Name = "EndIndexLabel";
            this.EndIndexLabel.Size = new System.Drawing.Size(91, 12);
            this.EndIndexLabel.TabIndex = 4;
            this.EndIndexLabel.Text = "끝 페이지(번호)";
            // 
            // EndIndexTextBox
            // 
            this.EndIndexTextBox.Location = new System.Drawing.Point(136, 110);
            this.EndIndexTextBox.Name = "EndIndexTextBox";
            this.EndIndexTextBox.Size = new System.Drawing.Size(69, 21);
            this.EndIndexTextBox.TabIndex = 5;
            // 
            // OkButton
            // 
            this.OkButton.Location = new System.Drawing.Point(313, 162);
            this.OkButton.Name = "OkButton";
            this.OkButton.Size = new System.Drawing.Size(75, 23);
            this.OkButton.TabIndex = 10;
            this.OkButton.Text = "OK";
            this.OkButton.UseVisualStyleBackColor = true;
            this.OkButton.Click += new System.EventHandler(this.OkButton_Click);
            // 
            // CancelButton
            // 
            this.CancelButton.Location = new System.Drawing.Point(394, 162);
            this.CancelButton.Name = "CancelButton";
            this.CancelButton.Size = new System.Drawing.Size(75, 23);
            this.CancelButton.TabIndex = 11;
            this.CancelButton.Text = "Cancel";
            this.CancelButton.UseVisualStyleBackColor = true;
            this.CancelButton.Click += new System.EventHandler(this.CancelButton_Click);
            // 
            // MultiPageSettingForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(511, 209);
            this.Controls.Add(this.CancelButton);
            this.Controls.Add(this.OkButton);
            this.Controls.Add(this.EndIndexTextBox);
            this.Controls.Add(this.EndIndexLabel);
            this.Controls.Add(this.StartIndexTextBox);
            this.Controls.Add(this.StartIndexLabel);
            this.Controls.Add(this.URLLabel);
            this.Controls.Add(this.URLTextBox);
            this.MaximizeBox = false;
            this.Name = "MultiPageSettingForm";
            this.Text = "MultiPageSettingForm";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox URLTextBox;
        private System.Windows.Forms.Label URLLabel;
        private System.Windows.Forms.Label StartIndexLabel;
        private System.Windows.Forms.TextBox StartIndexTextBox;
        private System.Windows.Forms.Label EndIndexLabel;
        private System.Windows.Forms.TextBox EndIndexTextBox;
        private System.Windows.Forms.Button OkButton;
        private System.Windows.Forms.Button CancelButton;
    }
}