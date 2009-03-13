#region Copyright Notice
// This file is part of PowerPoint Language Painter.
// 
// Copyright (C) 2008/2009 Andreas Kirsch
// 
// PowerPoint Language Painter is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
// 
// PowerPoint Language Painter is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
// 
// You should have received a copy of the GNU General Public License
// along with this program.  If not, see <http://www.gnu.org/licenses/>.
#endregion

namespace PowerPointLanguagePainter
{
    partial class PainterRibbon
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
            Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
            this.paintToggleButton = new Microsoft.Office.Tools.Ribbon.RibbonSplitButton();
            this.languageIDEnglishUS = new Microsoft.Office.Tools.Ribbon.RibbonButton();
            this.languageIDGerman = new Microsoft.Office.Tools.Ribbon.RibbonButton();
            this.languageIDFrench = new Microsoft.Office.Tools.Ribbon.RibbonButton();
            this.languageIDSpanish = new Microsoft.Office.Tools.Ribbon.RibbonButton();
            this.tab1 = new Microsoft.Office.Tools.Ribbon.RibbonTab();
            group1 = new Microsoft.Office.Tools.Ribbon.RibbonGroup();
            group1.SuspendLayout();
            this.tab1.SuspendLayout();
            this.SuspendLayout();
            // 
            // group1
            // 
            group1.Items.Add(this.paintToggleButton);
            group1.Label = "Language Painter";
            group1.Name = "group1";
            group1.Position = Microsoft.Office.Tools.Ribbon.RibbonPosition.AfterOfficeId("GroupEditing");
            // 
            // paintToggleButton
            // 
            this.paintToggleButton.ButtonType = Microsoft.Office.Tools.Ribbon.RibbonButtonType.ToggleButton;
            this.paintToggleButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.paintToggleButton.Items.Add(this.languageIDEnglishUS);
            this.paintToggleButton.Items.Add(this.languageIDGerman);
            this.paintToggleButton.Items.Add(this.languageIDFrench);
            this.paintToggleButton.Items.Add(this.languageIDSpanish);
            this.paintToggleButton.Label = "Paint Language";
            this.paintToggleButton.Name = "paintToggleButton";
            this.paintToggleButton.OfficeImageId = "Spelling";
            this.paintToggleButton.ScreenTip = "Language Painter";
            this.paintToggleButton.SuperTip = "If toggled, automatically sets the language of the current text to the one specif" +
                "ied (in the drop-box below).";
            this.paintToggleButton.Click += new System.EventHandler<Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs>(this.paintToggleButton_Click);
            // 
            // languageIDEnglishUS
            // 
            this.languageIDEnglishUS.Label = "English (US)";
            this.languageIDEnglishUS.Name = "languageIDEnglishUS";
            this.languageIDEnglishUS.ShowImage = true;
            this.languageIDEnglishUS.Click += new System.EventHandler<Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs>(this.languageIDEnglishUS_Click);
            // 
            // languageIDGerman
            // 
            this.languageIDGerman.Label = "German";
            this.languageIDGerman.Name = "languageIDGerman";
            this.languageIDGerman.ShowImage = true;
            this.languageIDGerman.Click += new System.EventHandler<Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs>(this.languageIDGerman_Click);
            // 
            // languageIDFrench
            // 
            this.languageIDFrench.Label = "French";
            this.languageIDFrench.Name = "languageIDFrench";
            this.languageIDFrench.ShowImage = true;
            this.languageIDFrench.Click += new System.EventHandler<Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs>(this.languageIDFrench_Click);
            // 
            // languageIDSpanish
            // 
            this.languageIDSpanish.Label = "Spanish";
            this.languageIDSpanish.Name = "languageIDSpanish";
            this.languageIDSpanish.ShowImage = true;
            this.languageIDSpanish.Click += new System.EventHandler<Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs>(this.languageIDSpanish_Click);
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.ControlId.OfficeId = "TabHome";
            this.tab1.Groups.Add(group1);
            this.tab1.Label = "TabHome";
            this.tab1.Name = "tab1";
            // 
            // PainterRibbon
            // 
            this.Name = "PainterRibbon";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.tab1);
            group1.ResumeLayout(false);
            group1.PerformLayout();
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton languageIDEnglishUS;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton languageIDGerman;
        private Microsoft.Office.Tools.Ribbon.RibbonSplitButton paintToggleButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton languageIDFrench;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton languageIDSpanish;
    }

    partial class ThisRibbonCollection : Microsoft.Office.Tools.Ribbon.RibbonReadOnlyCollection
    {
        internal PainterRibbon PainterRibbon
        {
            get { return this.GetRibbon<PainterRibbon>(); }
        }
    }
}
