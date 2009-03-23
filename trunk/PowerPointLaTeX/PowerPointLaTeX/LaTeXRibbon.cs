#region Copyright Notice
// This file is part of PowerPoint LaTeX.
// 
// Copyright (C) 2008/2009 Andreas Kirsch
// 
// PowerPoint LaTeX is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
// 
// PowerPoint LaTeX is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
// 
// You should have received a copy of the GNU General Public License
// along with this program.  If not, see <http://www.gnu.org/licenses/>.
#endregion

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
            SettingsTags.ManualEquationEditingChanged += new SettingsTags.ToggleChangedEventHandler(SettingsTags_ManualEquationEditingChanged);
        }

        public void RegisterApplicationEvents()
        {
            // hook our(the ribbon) event listeners into the application
            Application.WindowSelectionChange += new EApplication_WindowSelectionChangeEventHandler(Application_WindowSelectionChange);
            Application.WindowActivate += new EApplication_WindowActivateEventHandler(Application_WindowActivate);
        }

        private void SettingsTags_ManualEquationEditingChanged(bool enabled)
        {
            AutoEditEquationToggle.Checked = !enabled;
        }

        private void SettingsTags_PresentationModeChanged(bool enabled)
        {
            PresentationModeToggle.Checked = enabled;
        }

        private void SettingsTags_ManualPreviewChanged(bool enabled)
        {
            CompileButton.Enabled = enabled;
            DecompileButton.Enabled = enabled;

            AutomaticCompilationToggle.Checked = !enabled;
        }

        private void Default_PropertyChanged(object sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            if (e.PropertyName == "ShowDeveloperTaskPane")
            {
                DeveloperTaskPaneToggle.Checked = (bool) ((Properties.Settings) sender)[e.PropertyName];
            }
        }

        private void Application_WindowSelectionChange(Selection Sel)
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

        private void Application_WindowActivate(Presentation Pres, DocumentWindow Wn)
        {
            SettingsTags tags = Pres.SettingsTags();
            AutoEditEquationToggle.Checked = !tags.ManualEquationEditing;
            AutomaticCompilationToggle.Checked = !tags.ManualPreview;
            PresentationModeToggle.Checked = tags.PresentationMode;
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

            Selection selection = Application.ActiveWindow.Selection;
            List<Microsoft.Office.Interop.PowerPoint.Shape> shapes = selection.GetShapes();

            // deselect the current selection to avoid decompiling it straight away..
            Application.ActiveWindow.Selection.Unselect();

            Slide slide = Tool.ActiveSlide;
            if (slide != null)
            {
                if (shapes.Count == 0)
                {
                    Tool.CompileSlide(slide);
                }
                else
                {
                    foreach (Microsoft.Office.Interop.PowerPoint.Shape shape in shapes)
                    {
                        Tool.CompileShape(slide, shape);
                    }
                }
            }
        }

        private void DecompileButton_Click(object sender, RibbonControlEventArgs e)
        {
            // an exception is thrown otherwise >_>
            if (!Tool.EnableAddIn)
            {
                return;
            }

            // TODO: this is copy from CompileButton - find a way to merge the two [3/12/2009 Andreas]
            Selection selection = Application.ActiveWindow.Selection;
            List<Microsoft.Office.Interop.PowerPoint.Shape> shapes = selection.GetShapes();

            Slide slide = Tool.ActiveSlide;
            if (slide != null)
            {
                if (shapes.Count == 0)
                {
                    Tool.DecompileSlide(slide);
                }
                else
                {
                    foreach (Microsoft.Office.Interop.PowerPoint.Shape shape in shapes)
                    {
                        Tool.DecompileShape(slide, shape);
                    }
                }
            }
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
                // don't embed fonts..
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

        private void CreateFormula_Click(object sender, RibbonControlEventArgs e)
        {
            Tool.CreateEmptyEquation();
        }

        private void ShowEquationCode_Click(object sender, RibbonControlEventArgs e)
        {
            // get the currently selected shape
            Selection selection = Application.ActiveWindow.Selection;
            List<Microsoft.Office.Interop.PowerPoint.Shape> shapes = selection.GetShapes();
            foreach (Microsoft.Office.Interop.PowerPoint.Shape shape in shapes)
            {
                if( shape.LaTeXTags().Type == EquationType.Equation && !LaTeXTool.IsShapeUncompiledEquation(shape)) {
                    Tool.ShowEquationSource(shape);
                }
            }
        }

        private void AutoEditEquationToggle_Click(object sender, RibbonControlEventArgs e)
        {
            Tool.ActivePresentation.SettingsTags().ManualEquationEditing.value = !AutoEditEquationToggle.Checked;
        }
    }
}
