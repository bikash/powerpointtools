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

namespace PowerPointLaTeX
{
    partial class LaTeXRibbon
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(LaTeXRibbon));
            this.LaTeX = new Microsoft.Office.Tools.Ribbon.RibbonTab();
            this.generalGroup = new Microsoft.Office.Tools.Ribbon.RibbonGroup();
            this.PreferencesButton = new Microsoft.Office.Tools.Ribbon.RibbonButton();
            this.ClearCache = new Microsoft.Office.Tools.Ribbon.RibbonButton();
            this.PresentationModeToggle = new Microsoft.Office.Tools.Ribbon.RibbonToggleButton();
            this.separator1 = new Microsoft.Office.Tools.Ribbon.RibbonSeparator();
            this.FinalizeButton = new Microsoft.Office.Tools.Ribbon.RibbonButton();
            this.inlineGroup = new Microsoft.Office.Tools.Ribbon.RibbonGroup();
            this.AutomaticCompilationToggle = new Microsoft.Office.Tools.Ribbon.RibbonToggleButton();
            this.DecompileButton = new Microsoft.Office.Tools.Ribbon.RibbonButton();
            this.CompileButton = new Microsoft.Office.Tools.Ribbon.RibbonButton();
            this.equationGroup = new Microsoft.Office.Tools.Ribbon.RibbonGroup();
            this.AutoEditEquationToggle = new Microsoft.Office.Tools.Ribbon.RibbonToggleButton();
            this.CreateFormula = new Microsoft.Office.Tools.Ribbon.RibbonButton();
            this.ShowEquationCode = new Microsoft.Office.Tools.Ribbon.RibbonButton();
            this.tab1 = new Microsoft.Office.Tools.Ribbon.RibbonTab();
            this.LaTeXGroup = new Microsoft.Office.Tools.Ribbon.RibbonGroup();
            this.DeveloperTaskPaneToggle = new Microsoft.Office.Tools.Ribbon.RibbonToggleButton();
            this.LaTeX.SuspendLayout();
            this.generalGroup.SuspendLayout();
            this.inlineGroup.SuspendLayout();
            this.equationGroup.SuspendLayout();
            this.tab1.SuspendLayout();
            this.LaTeXGroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // LaTeX
            // 
            this.LaTeX.Groups.Add(this.generalGroup);
            this.LaTeX.Groups.Add(this.inlineGroup);
            this.LaTeX.Groups.Add(this.equationGroup);
            this.LaTeX.Label = "LaTeX";
            this.LaTeX.Name = "LaTeX";
            // 
            // generalGroup
            // 
            this.generalGroup.Items.Add(this.PreferencesButton);
            this.generalGroup.Items.Add(this.ClearCache);
            this.generalGroup.Items.Add(this.PresentationModeToggle);
            this.generalGroup.Items.Add(this.separator1);
            this.generalGroup.Items.Add(this.FinalizeButton);
            this.generalGroup.Label = "General";
            this.generalGroup.Name = "generalGroup";
            // 
            // PreferencesButton
            // 
            this.PreferencesButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.PreferencesButton.Label = "Preferences";
            this.PreferencesButton.Name = "PreferencesButton";
            this.PreferencesButton.OfficeImageId = "MessageOptions";
            this.PreferencesButton.ShowImage = true;
            this.PreferencesButton.Visible = false;
            this.PreferencesButton.Click += new System.EventHandler<Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs>(this.PreferencesButton_Click);
            // 
            // ClearCache
            // 
            this.ClearCache.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.ClearCache.Label = "Clear Cache";
            this.ClearCache.Name = "ClearCache";
            this.ClearCache.OfficeImageId = "Delete";
            this.ClearCache.ScreenTip = "Clear the LaTeX Cache";
            this.ClearCache.ShowImage = true;
            this.ClearCache.SuperTip = resources.GetString("ClearCache.SuperTip");
            this.ClearCache.Click += new System.EventHandler<Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs>(this.ClearCache_Click);
            // 
            // PresentationModeToggle
            // 
            this.PresentationModeToggle.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.PresentationModeToggle.Description = "Compiles everything and protects it from changes";
            this.PresentationModeToggle.Label = "Protect Formulas";
            this.PresentationModeToggle.Name = "PresentationModeToggle";
            this.PresentationModeToggle.OfficeImageId = "ProtectDocument";
            this.PresentationModeToggle.ScreenTip = "Protect Formulas";
            this.PresentationModeToggle.ShowImage = true;
            this.PresentationModeToggle.SuperTip = "Compile all LaTeX $$formulas$$ and protect them from changes (locks all text area" +
                "s that contain formulas, etc.)";
            this.PresentationModeToggle.Click += new System.EventHandler<Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs>(this.PresentationModeToggle_Click);
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // FinalizeButton
            // 
            this.FinalizeButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.FinalizeButton.Description = "Compile everything and remove all meta-information";
            this.FinalizeButton.Label = "Finalize";
            this.FinalizeButton.Name = "FinalizeButton";
            this.FinalizeButton.OfficeImageId = "MenuPublish";
            this.FinalizeButton.ScreenTip = "Finalize";
            this.FinalizeButton.ShowImage = true;
            this.FinalizeButton.SuperTip = "Compile all LaTeX $$formulas$$ and remove all metainformation from this addin fro" +
                "m the presentation.";
            this.FinalizeButton.Click += new System.EventHandler<Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs>(this.FinalizeButton_Click);
            // 
            // inlineGroup
            // 
            this.inlineGroup.Items.Add(this.AutomaticCompilationToggle);
            this.inlineGroup.Items.Add(this.DecompileButton);
            this.inlineGroup.Items.Add(this.CompileButton);
            this.inlineGroup.Label = "Inline Formulas";
            this.inlineGroup.Name = "inlineGroup";
            // 
            // AutomaticCompilationToggle
            // 
            this.AutomaticCompilationToggle.Checked = true;
            this.AutomaticCompilationToggle.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.AutomaticCompilationToggle.Description = "Automatically compile LaTeX shortcodes";
            this.AutomaticCompilationToggle.Image = ((System.Drawing.Image) (resources.GetObject("AutomaticCompilationToggle.Image")));
            this.AutomaticCompilationToggle.Label = "Automatic Preview";
            this.AutomaticCompilationToggle.Name = "AutomaticCompilationToggle";
            this.AutomaticCompilationToggle.ScreenTip = "Automatic Preview";
            this.AutomaticCompilationToggle.ShowImage = true;
            this.AutomaticCompilationToggle.SuperTip = "Compiles and Decompiles LaTeX $$formulas$$ automatically.";
            this.AutomaticCompilationToggle.Click += new System.EventHandler<Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs>(this.AutomaticCompilationToggle_Click);
            // 
            // DecompileButton
            // 
            this.DecompileButton.Label = "Decompile Selection";
            this.DecompileButton.Name = "DecompileButton";
            this.DecompileButton.OfficeImageId = "MailMergeGoToPreviousRecord";
            this.DecompileButton.ScreenTip = "Decompile Selection/Slide";
            this.DecompileButton.ShowImage = true;
            this.DecompileButton.SuperTip = "Decompiles all LaTeX $$formulas$$ in the selection or in the active slide, if not" +
                "hing is selected.";
            this.DecompileButton.Click += new System.EventHandler<Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs>(this.DecompileButton_Click);
            // 
            // CompileButton
            // 
            this.CompileButton.Label = "Compile Selection";
            this.CompileButton.Name = "CompileButton";
            this.CompileButton.OfficeImageId = "MacroPlay";
            this.CompileButton.ScreenTip = "Compile Selection/Slide";
            this.CompileButton.ShowImage = true;
            this.CompileButton.SuperTip = "Compiles all LaTeX $$formulas$$ in the selection or in the active slide, if nothi" +
                "ng is selected.";
            this.CompileButton.Click += new System.EventHandler<Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs>(this.CompileButton_Click);
            // 
            // equationGroup
            // 
            this.equationGroup.Items.Add(this.AutoEditEquationToggle);
            this.equationGroup.Items.Add(this.CreateFormula);
            this.equationGroup.Items.Add(this.ShowEquationCode);
            this.equationGroup.Label = "Equations";
            this.equationGroup.Name = "equationGroup";
            // 
            // AutoEditEquationToggle
            // 
            this.AutoEditEquationToggle.Checked = true;
            this.AutoEditEquationToggle.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.AutoEditEquationToggle.Image = ((System.Drawing.Image) (resources.GetObject("AutoEditEquationToggle.Image")));
            this.AutoEditEquationToggle.Label = "Auto-Edit Equations";
            this.AutoEditEquationToggle.Name = "AutoEditEquationToggle";
            this.AutoEditEquationToggle.OfficeImageId = "FunctionWizard";
            this.AutoEditEquationToggle.ShowImage = true;
            this.AutoEditEquationToggle.Click += new System.EventHandler<Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs>(this.AutoEditEquationToggle_Click);
            // 
            // CreateFormula
            // 
            this.CreateFormula.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.CreateFormula.Label = "New Formula";
            this.CreateFormula.Name = "CreateFormula";
            this.CreateFormula.OfficeImageId = "FunctionWizard";
            this.CreateFormula.ShowImage = true;
            this.CreateFormula.Click += new System.EventHandler<Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs>(this.CreateFormula_Click);
            // 
            // ShowEquationCode
            // 
            this.ShowEquationCode.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.ShowEquationCode.Label = "Show Equation Code";
            this.ShowEquationCode.Name = "ShowEquationCode";
            this.ShowEquationCode.OfficeImageId = "WordArtEditTextClassic";
            this.ShowEquationCode.ShowImage = true;
            this.ShowEquationCode.Click += new System.EventHandler<Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs>(this.ShowEquationCode_Click);
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.ControlId.OfficeId = "TabDeveloper";
            this.tab1.Groups.Add(this.LaTeXGroup);
            this.tab1.Label = "Developer";
            this.tab1.Name = "tab1";
            // 
            // LaTeXGroup
            // 
            this.LaTeXGroup.Items.Add(this.DeveloperTaskPaneToggle);
            this.LaTeXGroup.Label = "LaTeX";
            this.LaTeXGroup.Name = "LaTeXGroup";
            // 
            // DeveloperTaskPaneToggle
            // 
            this.DeveloperTaskPaneToggle.Checked = global::PowerPointLaTeX.Properties.Settings.Default.ShowDeveloperTaskPane;
            this.DeveloperTaskPaneToggle.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.DeveloperTaskPaneToggle.Label = "Developer Info";
            this.DeveloperTaskPaneToggle.Name = "DeveloperTaskPaneToggle";
            this.DeveloperTaskPaneToggle.OfficeImageId = "FileDocumentInspect";
            this.DeveloperTaskPaneToggle.ShowImage = true;
            this.DeveloperTaskPaneToggle.Click += new System.EventHandler<Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs>(this.DeveloperTaskPaneToggle_Click);
            // 
            // LaTeXRibbon
            // 
            this.Name = "LaTeXRibbon";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.LaTeX);
            this.Tabs.Add(this.tab1);
            this.Load += new System.EventHandler<Microsoft.Office.Tools.Ribbon.RibbonUIEventArgs>(this.LaTeXRibbon_Load);
            this.LaTeX.ResumeLayout(false);
            this.LaTeX.PerformLayout();
            this.generalGroup.ResumeLayout(false);
            this.generalGroup.PerformLayout();
            this.inlineGroup.ResumeLayout(false);
            this.inlineGroup.PerformLayout();
            this.equationGroup.ResumeLayout(false);
            this.equationGroup.PerformLayout();
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.LaTeXGroup.ResumeLayout(false);
            this.LaTeXGroup.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab LaTeX;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup inlineGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton CompileButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton DecompileButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup generalGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton PresentationModeToggle;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton AutomaticCompilationToggle;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton FinalizeButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton PreferencesButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup LaTeXGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton DeveloperTaskPaneToggle;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup equationGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton CreateFormula;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton AutoEditEquationToggle;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ShowEquationCode;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ClearCache;
    }

    partial class ThisRibbonCollection : Microsoft.Office.Tools.Ribbon.RibbonReadOnlyCollection
    {
        internal LaTeXRibbon LaTeXRibbon
        {
            get { return this.GetRibbon<LaTeXRibbon>(); }
        }
    }
}
