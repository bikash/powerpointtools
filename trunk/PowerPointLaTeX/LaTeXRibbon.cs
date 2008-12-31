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

        private Slide GetActiveSlide()
        {
            return ((Slide) Globals.ThisAddIn.Application.ActiveWindow.View.Slide);
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Tool.CompileSlide(GetActiveSlide());
            
            
/*
            Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
            Microsoft.Vbe.Interop.VBComponent component = presentation.VBProject.VBComponents.Add(Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_StdModule);
            component.CodeModule.AddFromString(
@"
Sub Test__()
    MsgBox ""Hello World""
End Sub
"
                );

            //Globals.ThisAddIn.Tool.CompileSlide(slide);
            Shape shape = slide.Shapes[1];
            shape.TextFrame.TextRange.Text = "Hello World";
            ActionSetting setting = shape.ActionSettings[PpMouseActivation.ppMouseClick];
            setting.Action = PpActionType.ppActionRunMacro;
            setting.Run = "Test__";*/

        }

        private void DecompileSlide_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Tool.DecompileSlide(GetActiveSlide());)
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
