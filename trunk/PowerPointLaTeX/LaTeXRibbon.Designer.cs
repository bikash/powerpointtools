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
            this.LaTeX = new Microsoft.Office.Tools.Ribbon.RibbonTab();
            this.inlineGroup = new Microsoft.Office.Tools.Ribbon.RibbonGroup();
            this.DecompileSlide = new Microsoft.Office.Tools.Ribbon.RibbonButton();
            this.CompileSlide = new Microsoft.Office.Tools.Ribbon.RibbonButton();
            this.generalGroup = new Microsoft.Office.Tools.Ribbon.RibbonGroup();
            this.offlineModeToggle = new Microsoft.Office.Tools.Ribbon.RibbonToggleButton();
            this.automaticCompilationToggle = new Microsoft.Office.Tools.Ribbon.RibbonToggleButton();
            this.bakeButton = new Microsoft.Office.Tools.Ribbon.RibbonButton();
            this.LaTeX.SuspendLayout();
            this.inlineGroup.SuspendLayout();
            this.generalGroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // LaTeX
            // 
            this.LaTeX.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.LaTeX.Groups.Add(this.generalGroup);
            this.LaTeX.Groups.Add(this.inlineGroup);
            this.LaTeX.Label = "TabAddIns";
            this.LaTeX.Name = "LaTeX";
            // 
            // inlineGroup
            // 
            this.inlineGroup.Items.Add(this.automaticCompilationToggle);
            this.inlineGroup.Items.Add(this.DecompileSlide);
            this.inlineGroup.Items.Add(this.CompileSlide);
            this.inlineGroup.Label = "Inline Formulas";
            this.inlineGroup.Name = "inlineGroup";
            // 
            // DecompileSlide
            // 
            this.DecompileSlide.Label = "Decompile Selection";
            this.DecompileSlide.Name = "DecompileSlide";
            this.DecompileSlide.ShowImage = true;
            this.DecompileSlide.Click += new System.EventHandler<Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs>(this.DecompileSlide_Click);
            // 
            // CompileSlide
            // 
            this.CompileSlide.Label = "Compile Selection";
            this.CompileSlide.Name = "CompileSlide";
            this.CompileSlide.ShowImage = true;
            this.CompileSlide.Click += new System.EventHandler<Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs>(this.button1_Click);
            // 
            // generalGroup
            // 
            this.generalGroup.Items.Add(this.offlineModeToggle);
            this.generalGroup.Items.Add(this.bakeButton);
            this.generalGroup.Label = "General";
            this.generalGroup.Name = "generalGroup";
            // 
            // offlineModeToggle
            // 
            this.offlineModeToggle.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.offlineModeToggle.Label = "Offline Mode";
            this.offlineModeToggle.Name = "offlineModeToggle";
            this.offlineModeToggle.ShowImage = true;
            // 
            // automaticCompilationToggle
            // 
            this.automaticCompilationToggle.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.automaticCompilationToggle.Label = "Automatic Compilation";
            this.automaticCompilationToggle.Name = "automaticCompilationToggle";
            this.automaticCompilationToggle.ShowImage = true;
            // 
            // bakeButton
            // 
            this.bakeButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.bakeButton.Label = "Bake LaTeX";
            this.bakeButton.Name = "bakeButton";
            this.bakeButton.ShowImage = true;
            // 
            // LaTeXRibbon
            // 
            this.Name = "LaTeXRibbon";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.LaTeX);
            this.Load += new System.EventHandler<Microsoft.Office.Tools.Ribbon.RibbonUIEventArgs>(this.LaTeXRibbon_Load);
            this.LaTeX.ResumeLayout(false);
            this.LaTeX.PerformLayout();
            this.inlineGroup.ResumeLayout(false);
            this.inlineGroup.PerformLayout();
            this.generalGroup.ResumeLayout(false);
            this.generalGroup.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab LaTeX;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup inlineGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton CompileSlide;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton DecompileSlide;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup generalGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton offlineModeToggle;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton automaticCompilationToggle;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton bakeButton;
    }

    partial class ThisRibbonCollection : Microsoft.Office.Tools.Ribbon.RibbonReadOnlyCollection
    {
        internal LaTeXRibbon LaTeXRibbon
        {
            get { return this.GetRibbon<LaTeXRibbon>(); }
        }
    }
}
