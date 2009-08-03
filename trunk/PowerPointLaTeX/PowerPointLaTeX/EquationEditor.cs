using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.PowerPoint;
using System.Diagnostics;

namespace PowerPointLaTeX {
    public partial class EquationEditor : Form {
        private LaTeXTool Tool {
            get {
                return Globals.ThisAddIn.Tool;
            }
        }

        private System.Timers.Timer updatePreviewTimer;

        public String LaTeXCode {
            get { return formulaText.Text; }
        }

        public EquationEditor(String LaTeXCode ) {
            InitializeComponent();

            formulaText.Text = LaTeXCode;

            updatePreviewTimer = new System.Timers.Timer();
            updatePreviewTimer.Interval = 0.5 * 1000;
            updatePreviewTimer.AutoReset = false;
            updatePreviewTimer.Elapsed += new System.Timers.ElapsedEventHandler(updatePreviewTimer_Elapsed);
        }

        void worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e) {
        }

        void updatePreviewTimer_Elapsed(object sender, System.Timers.ElapsedEventArgs e) {
            updatePreview();
        }

        private void updatePreview() {
            UseWaitCursor = true;

            Image previewImage = Tool.GetImageForLaTeXCode(LaTeXCode);
            formulaPreview.Invoke( new MethodInvoker( delegate { formulaPreview.Image = previewImage; } ) );

            UseWaitCursor = false;
        }

        private void formulaText_TextChanged(object sender, EventArgs e) {
            updatePreviewTimer.Stop();
            updatePreviewTimer.Start();
        }       
    }
}
