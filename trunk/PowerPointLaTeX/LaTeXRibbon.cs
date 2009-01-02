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

        public LaTeXRibbon()
        {
            InitializeComponent();

            Properties.Settings.Default.PropertyChanged += new System.ComponentModel.PropertyChangedEventHandler(Default_PropertyChanged);

            SettingsTags.ManualPreviewChanged += new SettingsTags.ToggleChangedEventHandler(SettingsTags_ManualPreviewChanged);
            SettingsTags.PresentationModeChanged += new SettingsTags.ToggleChangedEventHandler(SettingsTags_PresentationModeChanged);
        }

        void SettingsTags_PresentationModeChanged(bool enabled)
        {
            PresentationModeToggle.Checked = enabled;
        }

        void SettingsTags_ManualPreviewChanged(bool enabled)
        {
            CompileButton.Enabled = enabled;
            DecompileButton.Enabled = enabled;

            AutomaticCompilationToggle.Checked = !enabled;
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

        private void LaTeXRibbon_Load(object sender, RibbonUIEventArgs e)
        {
        }

        private void CompileButton_Click(object sender, RibbonControlEventArgs e)
        {
            // an exception is thrown otherwise >_>
            if (!Tool.EnableAddIn)
            {
                return;
            }

            // unselect the current selection to avoid decompiling it straight away..
            Application.ActiveWindow.Selection.Unselect();

            Slide slide = Tool.ActiveSlide;
            if (slide != null)
                Tool.CompileSlide(slide);
        }

        private void DecompileButton_Click(object sender, RibbonControlEventArgs e)
        {
            // an exception is thrown otherwise >_>
            if (!Tool.EnableAddIn)
            {
                return;
            }

            Slide slide = Tool.ActiveSlide;
            if (slide != null)
                Tool.DecompileSlide(slide);
        }

        private void AutomaticCompilationToggle_Click(object sender, RibbonControlEventArgs e)
        {
            Tool.ActivePresentation.SettingsTags().ManualPreview.value = !AutomaticCompilationToggle.Checked;
        }

        private void PresentationModeToggle_Click(object sender, RibbonControlEventArgs e)
        {
            Tool.ActivePresentation.SettingsTags().PresentationMode.value = PresentationModeToggle.Checked;
        }

        private void FinalizeButton_Click(object sender, RibbonControlEventArgs e)
        {
            // an exception is thrown otherwise >_>
            if (!Tool.EnableAddIn)
            {
                return;
            }

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

        private void PreferencesButton_Click(object sender, RibbonControlEventArgs e)
        {
            Preferences preferences = new Preferences();
            preferences.ShowDialog();
        }
    }
}
