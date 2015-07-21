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
    public partial class CookieSettingForm : Form
    {
        bool chromeButtonChosen_Current = false, IEButtonChosen_Current = false, firefoxButtonChosen_Current = false;

        public CookieSettingForm()
        {
            InitializeComponent();
            chromeButtonChosen_Current = Globals.ThisAddIn.chromeButtonChosen;
            IEButtonChosen_Current = Globals.ThisAddIn.IEButtonChosen;
            firefoxButtonChosen_Current = Globals.ThisAddIn.firefoxButtonChosen;

            if (chromeButtonChosen_Current)
                ChromeButton.BackColor = Color.LightGreen;
            else
                ChromeButton.BackColor = Color.FromArgb(235, 235, 235);
            if (IEButtonChosen_Current)
                IEButton.BackColor = Color.LightGreen;
            else
                IEButton.BackColor = Color.FromArgb(235, 235, 235);
            if (firefoxButtonChosen_Current)
                FirefoxButton.BackColor = Color.LightGreen;
            else
                FirefoxButton.BackColor = Color.FromArgb(235, 235, 235);
            OKButton.BackColor = Color.FromArgb(235, 235, 235);
            CancelButton.BackColor = Color.FromArgb(235, 235, 235);
        }

        private void ChromeButton_Click(object sender, EventArgs e)
        {
            if (ChromeButton.BackColor == Color.LightGreen)
            {
                ChromeButton.BackColor = IEButton.BackColor;
                chromeButtonChosen_Current = false;
            }
            else
            {
                IEButton.BackColor = Color.FromArgb(235, 235, 235);
                FirefoxButton.BackColor = Color.FromArgb(235, 235, 235);
                ChromeButton.BackColor = Color.LightGreen;
                chromeButtonChosen_Current = true;
                IEButtonChosen_Current = false;
                firefoxButtonChosen_Current = false;
            }
        }

        private void IEButton_Click(object sender, EventArgs e)
        {
            if (IEButton.BackColor == Color.LightGreen)
            {
                IEButton.BackColor = ChromeButton.BackColor;
                IEButtonChosen_Current = false;
            }
            else
            {
                ChromeButton.BackColor = Color.FromArgb(235, 235, 235);
                FirefoxButton.BackColor = Color.FromArgb(235, 235, 235);
                IEButton.BackColor = Color.LightGreen;
                IEButtonChosen_Current = true;
                chromeButtonChosen_Current = false;
                firefoxButtonChosen_Current = false;
            }
        }

        private void FirefoxButton_Click(object sender, EventArgs e)
        {
            if (FirefoxButton.BackColor == Color.LightGreen)
            {
                FirefoxButton.BackColor = ChromeButton.BackColor;
                firefoxButtonChosen_Current = false;
            }
            else
            {
                ChromeButton.BackColor = Color.FromArgb(235, 235, 235);
                IEButton.BackColor = Color.FromArgb(235, 235, 235);
                FirefoxButton.BackColor = Color.LightGreen;
                firefoxButtonChosen_Current = true;
                chromeButtonChosen_Current = false;
                IEButtonChosen_Current = false;
            }
        }

        private void CancelButton_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void OKButton_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.chromeButtonChosen = chromeButtonChosen_Current;
            Globals.ThisAddIn.IEButtonChosen = IEButtonChosen_Current;
            Globals.ThisAddIn.firefoxButtonChosen = firefoxButtonChosen_Current;
            this.Close();
        }
    }
}
