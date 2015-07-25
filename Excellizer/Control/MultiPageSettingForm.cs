using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Excellizer.Control
{
    public partial class MultiPageSettingForm : Form
    {
        public MultiPageSettingForm(Form owner)
        {
            this.Owner = owner;
            InitializeComponent();
        }

        private void OkButton_Click(object sender, EventArgs e)
        {
            if(FrontURLTextBox.Text.Contains("http"))
            {
                ((BrowserForm)Owner).frontURL =  FrontURLTextBox.Text;
                ((BrowserForm)Owner).startIndex = int.Parse(StartIndexTextBox.Text);
                ((BrowserForm)Owner).endIndex = int.Parse(EndIndexTextBox.Text);
                ((BrowserForm)Owner).backURL = BackURLTextBox.Text;

                ((BrowserForm)Owner).MultiPageParse();

                this.Close();
            }
            else
            {
                MessageBox.Show("올바른 URL 폼을 입력하세요.");
            }
        }

        private void CancelButton_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
