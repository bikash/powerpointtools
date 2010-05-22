using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace PowerPointLaTeX {
    public partial class LogForm : Form {
        public LogForm(string log) {
            InitializeComponent();

            logBox.Text = log;
        }
    }
}
