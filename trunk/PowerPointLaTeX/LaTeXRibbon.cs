using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.PowerPoint;
using System.Diagnostics;
using System.Windows.Forms;
using System.IO;
using System.Drawing;

namespace PowerPointLaTeX
{
    public partial class LaTeXRibbon : OfficeRibbon
    {
        public LaTeXRibbon()
        {
            InitializeComponent();
        }

        private void LaTeXRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Slide slide = ((Slide) Globals.ThisAddIn.Application.ActiveWindow.View.Slide);
            Globals.ThisAddIn.Tool.CompileSlide(slide);
        }
        /*
                   string latexCode = @"a \le b";
                    string hexData = "";
                           foreach (byte c in data.content) {
                               hexData += c.ToString("X2");
                           }
                           //{\*\shppict }
                           Clipboard.SetText(@"{\pict \pngblip " + hexData + "}", TextDataFormat.Rtf);
                           //range.Paste();
                           TextRange picture = range.PasteSpecial(PpPasteDataType.ppPasteRTF,Microsoft.Office.Core.MsoTriState.msoFalse,"",0,"",Microsoft.Office.Core.MsoTriState.msoFalse);
 
        */
    }
}
