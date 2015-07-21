using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Excellizer.Control
{
    public partial class TransparentControl : Label //Control, Label
    {
        public TransparentControl()
        {
            InitializeComponent();
            InitializeUi();
        }

        void InitializeUi()
        {
            base.SetStyle(ControlStyles.OptimizedDoubleBuffer | ControlStyles.SupportsTransparentBackColor | ControlStyles.UserPaint, true);
            base.SetStyle(ControlStyles.Selectable | ControlStyles.FixedHeight, false);
            base.SetStyle(ControlStyles.ResizeRedraw, true);

            this.TabStop = false;
            this.BackColor = Color.Transparent;
        }

        protected override void OnPaintBackground(PaintEventArgs pevent)
        {
            base.OnPaintBackground(pevent);
        }

        protected override void OnPaint(PaintEventArgs pe)
        {
            Graphics g = pe.Graphics;

            //Frame
            Pen pen = new Pen(Color.FromArgb(100, Color.OrangeRed), 3);
            g.DrawRectangle(pen, 0, 0, Width - 1, Height - 1);

            //Fill
            SolidBrush br = new SolidBrush(Color.FromArgb(100, Color.OrangeRed));
            for (int i = 0; i < this.Height; i += 10)
            {
                if ((i / 10 & 0x0001) != 0x0001)
                    g.FillRectangle(br, 0, i, this.Width, 10);
            }

            base.OnPaint(pe);
        }
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
        /*protected override CreateParams CreateParams
        {
            get
            {
                CreateParams createParams = base.CreateParams;
                createParams.ExStyle |= 0x00000020; // WS_EX_TRANSPARENT
                return createParams;
            }
        }*/
    }
}