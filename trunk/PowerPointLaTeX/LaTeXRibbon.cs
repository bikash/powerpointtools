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
        private LaTeXTool Tool
        {
            get
            {
                return Globals.ThisAddIn.Tool;
            }
        }

        private Microsoft.Office.Interop.PowerPoint.Application Application
        {
            get
            {
                return Globals.ThisAddIn.Application;
            }
        }

        private Settings Settings
        {
            get
            {
                return Globals.ThisAddIn.Settings;
            }
        }

        public LaTeXRibbon()
        {
            InitializeComponent();
        }

        void Application_WindowSelectionChange(Selection Sel)
        {
            if (Sel.Type == PpSelectionType.ppSelectionNone)
            {
                CompileSlide.Label = "Compile Slide";
                DecompileSlide.Label = "Decompile Slide";
            }
            else
            {
                CompileSlide.Label = "Compile Selection";
                DecompileSlide.Label = "Decompile Selection";
            }
        }

        void Settings_onAutomaticCompilationChanged(bool isChecked)
        {
            CompileSlide.Enabled = !isChecked;
            DecompileSlide.Enabled = !isChecked;
        }

        private void LaTeXRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            Settings.onAutomaticCompilationChanged += new Settings.ToggleChangedEventHandler(Settings_onAutomaticCompilationChanged);
            Application.WindowSelectionChange += new EApplication_WindowSelectionChangeEventHandler(Application_WindowSelectionChange);
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Application.ActiveWindow.Selection.Unselect();
            Tool.CompileSlide(Tool.ActiveSlide);


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
            Tool.DecompileSlide(Tool.ActiveSlide);
        }

        private void automaticCompilationToggle_Click(object sender, RibbonControlEventArgs e)
        {
            Settings.AutomaticCompilation = Settings.AutomaticCompilation;
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
