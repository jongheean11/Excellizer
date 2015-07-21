using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace Excellizer.Control
{
    class Plexiglass : Form
    {
        public int _x = 0;
        public int _y = 0;
        public int _width = 0;
        public int _height = 0;
        public Plexiglass(Form tocover)
        {
            this.BackColor = Color.LightBlue;
            this.Opacity = 0.30;      // Tweak as desired
            this.FormBorderStyle = FormBorderStyle.None;
            this.ControlBox = false;
            this.ShowInTaskbar = false;
            this.StartPosition = FormStartPosition.Manual;
            this.AutoScaleMode = AutoScaleMode.None;
            this.Location = tocover.PointToScreen(Point.Empty);
            this.ClientSize = tocover.ClientSize;
            tocover.LocationChanged += Cover_LocationChanged;
            tocover.ClientSizeChanged += Cover_ClientSizeChanged;
            this.Show(tocover);
            tocover.Focus();
            // Disable Aero transitions, the plexiglass gets too visible
            if (Environment.OSVersion.Version.Major >= 6)
            {
                int value = 1;
                DwmSetWindowAttribute(tocover.Handle, DWMWA_TRANSITIONS_FORCEDISABLED, ref value, 4);
            }
        }

        public Plexiglass(Form tocover, int x, int y, int width, int height)
        {
            this.BackColor = Color.LightSkyBlue;
            this.Opacity = 0.30;      // Tweak as desired
            this.FormBorderStyle = FormBorderStyle.None;
            this.ControlBox = false;
            this.ShowInTaskbar = false;
            this.StartPosition = FormStartPosition.Manual;
            this.AutoScaleMode = AutoScaleMode.None;
            Point parentPtr = tocover.PointToScreen(Point.Empty);
            this.Location = new Point(x + parentPtr.X, y + parentPtr.Y);
            this.ClientSize = new Size(width, height);
            tocover.LocationChanged += Cover_LocationChanged;
            tocover.ClientSizeChanged += Cover_ClientSizeChanged;
            this.Show(tocover);
            tocover.Focus();
            // Disable Aero transitions, the plexiglass gets too visible
            if (Environment.OSVersion.Version.Major >= 6)
            {
                int value = 1;
                DwmSetWindowAttribute(tocover.Handle, DWMWA_TRANSITIONS_FORCEDISABLED, ref value, 4);
            }

            this._x = x;
            this._y = y;
            this._width = width;
            this._height = height;
        }
        private void Cover_LocationChanged(object sender, EventArgs e)
        {
            // Ensure the plexiglass follows the owner
            Point parentPtr = this.Owner.PointToScreen(Point.Empty);
            this.Location = new Point(_x + parentPtr.X, _y + parentPtr.Y);
            //this.Location = this.Owner.PointToScreen(Point.Empty);
        }
        private void Cover_ClientSizeChanged(object sender, EventArgs e)
        {
            // Ensure the plexiglass keeps the owner covered
            int ownerWidth = this.Owner.ClientSize.Width,
                ownerHeight = this.Owner.ClientSize.Height;
            
            if((_x + _width) > ownerWidth)
            {
                this.ClientSize = new Size(ownerWidth - _x, this.ClientSize.Height);
            }
            else
            {
                this.ClientSize = new Size(_width, this.ClientSize.Height);
            }
            if((_y + _height + 33) > ownerHeight)
            {
                this.ClientSize = new Size(this.ClientSize.Width, ownerHeight - (_y + 33));
            }
            else
            {
                this.ClientSize = new Size(this.ClientSize.Width, _height);
            }


            //this.ClientSize = this.Owner.ClientSize;
        }
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            // Restore owner
            this.Owner.LocationChanged -= Cover_LocationChanged;
            this.Owner.ClientSizeChanged -= Cover_ClientSizeChanged;
            if (!this.Owner.IsDisposed && Environment.OSVersion.Version.Major >= 6)
            {
                int value = 1;
                DwmSetWindowAttribute(this.Owner.Handle, DWMWA_TRANSITIONS_FORCEDISABLED, ref value, 4);
            }
            base.OnFormClosing(e);
        }
        protected override void OnActivated(EventArgs e)
        {
            // Always keep the owner activated instead
            this.BeginInvoke(new Action(() => this.Owner.Activate()));
        }
        private const int DWMWA_TRANSITIONS_FORCEDISABLED = 3;
        [DllImport("dwmapi.dll")]
        private static extern int DwmSetWindowAttribute(IntPtr hWnd, int attr, ref int value, int attrLen);

        protected override void WndProc(ref Message m)
        {
            const int WM_NCHITTEST = 0x0084;
            const int HTTRANSPARENT = (-1);

            if (m.Msg == WM_NCHITTEST)
            {
                m.Result = (IntPtr)HTTRANSPARENT;
            }
            else
            {
                base.WndProc(ref m);
            }
        }
    }
}
