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

        public void RegisterApplicationEvents()
        {
            // hook our(the ribbon) event listeners into the application
            Application.WindowSelectionChange += new EApplication_WindowSelectionChangeEventHandler(Application_WindowSelectionChange);
            Application.WindowActivate += new EApplication_WindowActivateEventHandler(Application_WindowActivate);
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

                EditEquationCode.Enabled = false;
            }
            else
            {
                CompileButton.Label.Replace("Slide", "Selection");
                DecompileButton.Label.Replace("Slide", "Selection");

                if( Sel.Type == PpSelectionType.ppSelectionShapes ) {
                    if( Sel.ShapeRange.Count == 1 && Sel.ShapeRange[1].IsEquation() ) {
                        EditEquationCode.Enabled = true;
                    }
                    else {
                        EditEquationCode.Enabled = false;
                    }
                }
            }
        }

        private void Application_WindowActivate(Presentation Pres, DocumentWindow Wn)
        {
            SettingsTags tags = Pres.SettingsTags();
            AutomaticCompilationToggle.Checked = !tags.ManualPreview;
            PresentationModeToggle.Checked = tags.PresentationMode;
        }

        private void LaTeXRibbon_Load(object sender, RibbonUIEventArgs e)
        {
        }

        private void CompileButton_Click(object sender, RibbonControlEventArgs e)
        {
            // an exception is thrown otherwise >_>
            if (!LaTeXTool.AddInEnabled)
            {
                return;
            }

            Selection selection = Application.ActiveWindow.Selection;
            if( selection.Type == PpSelectionType.ppSelectionSlides ) {
                foreach( Slide slide in selection.SlideRange ) {
                    LaTeXTool.CompileSlide(slide);
                }
            }
            else {
                List<Microsoft.Office.Interop.PowerPoint.Shape> shapes = selection.GetShapes();

                // deselect the current selection to avoid decompiling it straight away..
                Application.ActiveWindow.Selection.Unselect();

                Slide slide = LaTeXTool.ActiveSlide;
                if( slide != null ) {
                    if( shapes.Count == 0 ) {
                        LaTeXTool.CompileSlide( slide );
                    }
                    else {
                        foreach( Microsoft.Office.Interop.PowerPoint.Shape shape in shapes ) {
                            InlineFormulas.Embedding.CompileShape( slide, shape );
                        }
                    }
                }
            }
        }

        private void DecompileButton_Click(object sender, RibbonControlEventArgs e)
        {
            // an exception is thrown otherwise >_>
            if (!LaTeXTool.AddInEnabled)
            {
                return;
            }

            // TODO: this is copy from CompileButton - find a way to merge the two [3/12/2009 Andreas]
            Selection selection = Application.ActiveWindow.Selection;
            if( selection.Type == PpSelectionType.ppSelectionSlides ) {
                foreach( Slide slide in selection.SlideRange ) {
                    LaTeXTool.DecompileSlide( slide );
                }
            }
            else {
                List<Microsoft.Office.Interop.PowerPoint.Shape> shapes = selection.GetShapes();

                Slide slide = LaTeXTool.ActiveSlide;
                if( slide != null ) {
                    if( shapes.Count == 0 ) {
                        LaTeXTool.DecompileSlide( slide );
                    }
                    else {
                        foreach( Microsoft.Office.Interop.PowerPoint.Shape _shape in shapes ) {
                            Microsoft.Office.Interop.PowerPoint.Shape shape = _shape.SafeThis();
                            if( shape != null )
                                InlineFormulas.Embedding.DecompileShape( slide, shape );
                        }
                    }
                }
            }
        }

        private void AutomaticCompilationToggle_Click(object sender, RibbonControlEventArgs e)
        {
            LaTeXTool.ActivePresentation.SettingsTags().ManualPreview.value = !AutomaticCompilationToggle.Checked;
        }

        private void PresentationModeToggle_Click(object sender, RibbonControlEventArgs e)
        {
            LaTeXTool.ActivePresentation.SettingsTags().PresentationMode.value = PresentationModeToggle.Checked;
        }

        private void FinalizeButton_Click(object sender, RibbonControlEventArgs e)
        {
            // an exception is thrown otherwise >_>
            if (!LaTeXTool.AddInEnabled)
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

                if (dialog.Show() == 0 /*cancel button*/)
                {
                    return;
                }
                Debug.Assert(dialog.SelectedItems.Count == 1);
                string filename = dialog.SelectedItems.Item(1);
                // don't embed fonts..
                LaTeXTool.ActivePresentation.SaveCopyAs(filename, PpSaveAsFileType.ppSaveAsDefault, MsoTriState.msoFalse);
            }

            // finalize the presentation
            LaTeXTool.FinalizePresentation(LaTeXTool.ActivePresentation);
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
            Microsoft.Office.Interop.PowerPoint.Shape equation = EquationHandling.CreateEmptyEquation(LaTeXTool.ActiveSlide);
            bool cancelled;
            equation = EquationHandling.EditEquation(equation, out cancelled);
            if( !cancelled ) {
                equation.Select( MsoTriState.msoTrue );
            }
            else {
                equation.Delete();
            }
        }

        private void EditEquationCode_Click(object sender, RibbonControlEventArgs e)
        {
            // get the currently selected shape
            Selection selection = Application.ActiveWindow.Selection;
            List<Microsoft.Office.Interop.PowerPoint.Shape> shapes = selection.GetShapes();
            if( shapes.Count == 1 ) {
                Microsoft.Office.Interop.PowerPoint.Shape shape = shapes[0];
                if( shape.IsEquation() ) {
                    bool unused_cancelled;
                    shape = EquationHandling.EditEquation(shape, out unused_cancelled);
                    if( shape != null ) {
                        shape.Select(MsoTriState.msoCTrue);
                    }
                }
            }
        }

        private void ClearCache_Click(object sender, RibbonControlEventArgs e) {
            if (MessageBox.Show("Do you really want to clear the cache?", "PowerPoint LaTeX", MessageBoxButtons.YesNo ) == DialogResult.Yes ) {
                LaTeXTool.ActivePresentation.CacheTags().PurgeAll();
            }
        }

        private void showLastLogButton_Click( object sender, RibbonControlEventArgs e ) {
            LogForm logForm = new LogForm( Globals.ThisAddIn.LaTeXRenderingServices.Service.GetLastErrorReport() );
            logForm.ShowDialog();
        }

    }
}
