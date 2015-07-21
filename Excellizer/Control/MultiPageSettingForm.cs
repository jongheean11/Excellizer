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
        public MultiPageSettingForm()
        {
            InitializeComponent();
        }

        private void OkButton_Click(object sender, EventArgs e)
        {
            String url = URLTextBox.Text,
                startIndex = StartIndexTextBox.Text,
                endIndex = EndIndexTextBox.Text;
        }

        private void CancelButton_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
