using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace Priem
{
    public partial class Watch : Form
    {
        public Watch(int max_value)
        {
            InitializeComponent();
            CenterToParent();
            pBar.Minimum = 1;
            pBar.Maximum = max_value;
            pBar.Value = 1;
            pBar.Step = 1;
        }

        public void PerformStep()
        {
            pBar.PerformStep();
        }        
    }
}