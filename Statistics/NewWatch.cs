using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Priem
{
    public partial class NewWatch : Form
    {
        int iMax = 0;
        public NewWatch()
        {
            InitializeComponent();
            pBar.Value = 0;
            pBar.Step = 1;
            pBar.Maximum = 100;
        }

        public NewWatch(int max)
        {
            InitializeComponent();
            pBar.Maximum = max;
            iMax = max;
            pBar.Value = 1;
            pBar.Step = 1;
        }

        public void PerformStep()
        {
            if (iMax == 0)
                iMax = 100;
            if (iMax < pBar.Maximum)
                iMax++;

            pBar.Maximum = iMax;
            pBar.PerformStep();
        }

        public void SetText(string text)
        {
            lblStatus.Text = text;
            this.Refresh();
        }

        public void SetMax(int newMax)
        {
            if (newMax > iMax)
                iMax = newMax;
            else
                iMax++;
        }

        public void ZeroCount()
        {
            pBar.Value = 0;
        }
    }
}
