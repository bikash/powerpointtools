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
using Microsoft.Office.Core;

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

            Properties.Settings.Default.PropertyChanged += new System.ComponentModel.PropertyChangedEventHandler(Default_PropertyChanged);
        }

        void Default_PropertyChanged(object sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            if (e.PropertyName == "ShowDeveloperTaskPane")
            {
                DeveloperTaskPaneToggle.Checked = (bool) ((Properties.Settings) sender)[e.PropertyName];
            }
        }

        void Application_WindowSelectionChange(Selection Sel)
        {
            if (Sel.Type == PpSelectionType.ppSelectionNone)
            {
                CompileButton.Label.Replace("Selection", "Slide");
                DecompileButton.Label.Replace("Selection", "Slide");
            }
            else
            {
                CompileButton.Label.Replace("Slide", "Selection");
                DecompileButton.Label.Replace("Slide", "Selection");
            }
        }

        void Settings_onAutomaticCompilationChanged(bool isChecked)
        {
            CompileButton.Enabled = !isChecked;
            DecompileButton.Enabled = !isChecked;
        }

        private void LaTeXRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            Settings.onAutomaticCompilationChanged += new Settings.ToggleChangedEventHandler(Settings_onAutomaticCompilationChanged);
            Application.WindowSelectionChange += new EApplication_WindowSelectionChangeEventHandler(Application_WindowSelectionChange);
        }

        private void CompileButton_Click(object sender, RibbonControlEventArgs e)
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

        private void DecompileButton_Click(object sender, RibbonControlEventArgs e)
        {
            Tool.DecompileSlide(Tool.ActiveSlide);
        }

        private void AutomaticCompilationToggle_Click(object sender, RibbonControlEventArgs e)
        {
            Settings.AutomaticCompilation = Settings.AutomaticCompilation;
        }

        private void PresentationModeToggle_Click(object sender, RibbonControlEventArgs e)
        {
            Settings.OfflineMode = Settings.OfflineMode;
        }

        private void FinalizeButton_Click(object sender, RibbonControlEventArgs e)
        {
            DialogResult result;

            // make sure the user really wants to do this
            // TODO: resources.. [1/1/2009 Andreas]
            result = MessageBox.Show("Do you really want to finalize all formulas? You won't be able to edit them afterwards.", "PowerPointLaTeX", MessageBoxButtons.YesNo);
            if (result == DialogResult.No)
            {
                return;
            }

            // ask if a backup is wanted
            result = MessageBox.Show("Do you want to create a backup of your presentation?", "PowerPointLaTeX", MessageBoxButtons.YesNoCancel);
            if (result == DialogResult.Cancel)
            {
                return;
            }
            else if (result == DialogResult.Yes)
            {
                Microsoft.Office.Core.FileDialog dialog = Application.get_FileDialog(MsoFileDialogType.msoFileDialogSaveAs);
                dialog.AllowMultiSelect = false;

                if (dialog.Show() == 0)
                {
                    // cancel
                    return;
                }
                Debug.Assert(dialog.SelectedItems.Count == 1);
                string filename = dialog.SelectedItems.Item(1);
                // dont embed fonts..
                Tool.ActivePresentation.SaveCopyAs(filename, PpSaveAsFileType.ppSaveAsDefault, MsoTriState.msoFalse);
            }

            // finalize the presentation
            Tool.FinalizePresentation(Tool.ActivePresentation);
        }

        private void DeveloperTaskPaneToggle_Click(object sender, RibbonControlEventArgs e)
        {
            Properties.Settings.Default.ShowDeveloperTaskPane = DeveloperTaskPaneToggle.Checked;
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
